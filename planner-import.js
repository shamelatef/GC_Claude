/**
 * Microsoft Planner Excel Import
 * Handles the native "Export plan to Excel" format from Microsoft Planner.
 *
 * Planner export columns (case-insensitive, fuzzy matched):
 *   Task ID, Task Name, Bucket Name, Progress, Priority,
 *   Assigned To, Start Date, Due Date, Created Date,
 *   Completed Date, Notes, Checklist Item (subtask rows)
 *
 * The export is a FLAT table — every row is either a parent task or a
 * checklist/subtask row that belongs to the preceding parent.
 * We identify subtask rows because they have no Start/Due date OR because
 * a "Checklist Item" / "Subtask" column is populated.
 */

// ─── Header normalisation ────────────────────────────────────────────────────
function normKey(s) {
    return String(s || '').toLowerCase().replace(/[\s_\-\.%]+/g, '');
}

/**
 * Map raw column headers to canonical field keys.
 * Returns { taskName, bucket, progress, startDate, dueDate,
 *           priority, assignedTo, completedDate, notes, checklistItem,
 *           checklistStatus, parentTask, percentComplete }
 */
function buildPlannerHeaderMap(headers) {
    const map = {};
    const rules = [
        // canonical key → list of normalised header patterns that match
        ['taskId',          ['taskid','id']],
        ['taskName',        ['taskname','name','title','task']],
        ['bucket',          ['bucketname','bucket','sprint','epic','group','phase','category']],
        ['progress',        ['progress','state','status','taskstatus']],
        ['percentComplete', ['percentcomplete','complete','completion','percentdone','done','pctcomplete']],
        ['startDate',       ['startdate','start','startdatetime','scheduledstart']],
        ['dueDate',         ['duedate','due','duedatetime','enddate','end','finishdate','finish','deadline']],
        ['priority',        ['priority','urgency']],
        ['assignedTo',      ['assignedto','assignee','owner','responsible']],
        ['completedDate',   ['completeddate','completiondatetime','closedate']],
        ['notes',           ['notes','description','details','comment']],
        ['checklistItem',   ['checklistitem','checklist','subtask','subtaskname','checklistitemtitle','Checklist Item']],
        ['checklistStatus', ['checklistitemstate','checkliststatus','subtaskstatus','subtaskstate']],
        ['parentTask',      ['parenttask','parenttaskname','parentid']],
    ];

    for (const h of headers) {
        const nk = normKey(h);
        for (const [field, patterns] of rules) {
            if (!map[field] && patterns.some(p => nk === p || nk.startsWith(p))) {
                map[field] = h;
                break;
            }
        }
    }
    return map;
}

// ─── Progress parsing ────────────────────────────────────────────────────────
/**
 * Parse various progress representations into a 0-100 integer.
 * Handles: "50%", "0.5", "50", "In Progress", "Completed", "Not started" etc.
 */
function parseProgress(raw) {
    if (raw === null || raw === undefined || raw === '') return null;
    const s = String(raw).trim();

    // Numeric with %
    const pctMatch = s.match(/^(\d+(?:\.\d+)?)\s*%$/);
    if (pctMatch) return Math.round(Math.min(100, Math.max(0, parseFloat(pctMatch[1]))));

    // Pure number
    const numMatch = s.match(/^(\d+(?:\.\d+)?)$/);
    if (numMatch) {
        const v = parseFloat(numMatch[1]);
        return Math.round(v <= 1 ? v * 100 : Math.min(100, v));
    }

    // Planner progress text values
    const lower = s.toLowerCase().replace(/[\s_\-]+/g, '');
    if (['notstarted','notyetstarted','new'].includes(lower)) return 0;
    if (['inprogress','ongoing','active','started'].includes(lower)) return 50;
    if (['completed','done','closed','finished','complete'].includes(lower)) return 100;
    if (['late','overdue','delayed'].includes(lower)) return 25;

    return null;
}

// ─── Status mapping ──────────────────────────────────────────────────────────
/**
 * Map Planner progress text / percent to one of the 6 Gantt statuses.
 */
function plannerProgressToStatus(progressRaw, pctComplete) {
    const pct = pctComplete != null ? pctComplete : parseProgress(progressRaw);
    const lower = String(progressRaw || '').toLowerCase().replace(/[\s_\-]+/g, '');

    if (lower === 'completed' || lower === 'done' || pct === 100) return 'Completed';
    if (lower === 'late' || lower === 'overdue') return 'Delayed';
    if (lower === 'blocked') return 'Blocked';
    if (lower === 'inprogress' || lower === 'ongoing' || lower === 'active' || (pct != null && pct > 0 && pct < 100)) return 'In Progress';
    if (lower === 'notstarted' || lower === 'notyetstarted' || pct === 0) return 'Not Started';
    return 'Not Started';
}

// ─── Date parsing ────────────────────────────────────────────────────────────
function parsePlannerDate(raw) {
    if (!raw && raw !== 0) return null;
    const s = String(raw).trim();
    if (!s || s === '0' || s.toLowerCase() === 'n/a' || s.toLowerCase() === 'none') return null;

    // Excel serial date (number)
    const serial = parseFloat(s);
    if (!isNaN(serial) && serial > 1000 && serial < 100000) {
        // Excel epoch: Dec 30 1899
        const ms = (serial - 25569) * 86400000;
        const d = new Date(ms);
        if (!isNaN(d)) {
            return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`;
        }
    }

    // ISO yyyy-mm-dd
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;

    // m/d/yyyy or m/d/yy
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
        let y = parseInt(m[3]);
        if (y < 100) y += y < 50 ? 2000 : 1900;
        return `${y}-${String(m[1]).padStart(2,'0')}-${String(m[2]).padStart(2,'0')}`;
    }

    // d.m.yyyy (European)
    m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;

    // Try native Date parse as fallback
    try {
        const d = new Date(s);
        if (!isNaN(d)) {
            return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
        }
    } catch(_) {}

    return null;
}

// ─── Main import function ────────────────────────────────────────────────────
/**
 * Process rows from a Planner Excel export and return { imported, skipped, warnings }.
 *
 * Mapping:
 *   "Task Name" column          →  Group name
 *   "Checklist Item" column     →  Semicolon-separated task names  "item1;item2;item3"
 *                                  Each token becomes one Gantt task under that group.
 *   No checklist / empty cell   →  One fallback task named after the group.
 *
 * All tasks in a group inherit the parent row's Start Date, Due Date,
 * Progress, and Status.
 *
 * Side-effects: mutates global `tasks`, `groups`, `groupStates`, `groupOrder`.
 */
function importPlannerRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) {
        return { imported: 0, skipped: 0, warnings: ['No rows found in file.'] };
    }

    const headers = Object.keys(rows[0]);
    const hm      = buildPlannerHeaderMap(headers);
    const warnings = [];

    if (!hm.taskName) {
        return { imported: 0, skipped: 0, warnings: ['Could not find a Task Name column. Make sure you are importing the original Planner Excel export.'] };
    }

    const get = (row, field) => hm[field] ? String(row[hm[field]] ?? '').trim() : '';

    let imported = 0;
    let skipped  = 0;
    let idBase   = Date.now();

    for (const row of rows) {
        const rawName      = get(row, 'taskName');
        const rawStart     = get(row, 'startDate');
        const rawDue       = get(row, 'dueDate');
        const rawProg      = get(row, 'progress');
        const rawPct       = get(row, 'percentComplete');
        const rawChecklist = get(row, 'checklistItem');   // e.g. "Design;Build;Test"

        if (!rawName) { skipped++; continue; }

        // ── Dates ─────────────────────────────────────────────────────────
        const startISO = parsePlannerDate(rawStart);
        const dueISO   = parsePlannerDate(rawDue);

        const today    = new Date();
        const todayISO = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
        const tmr      = new Date(today); tmr.setDate(tmr.getDate() + 1);
        const tmrISO   = `${tmr.getFullYear()}-${String(tmr.getMonth()+1).padStart(2,'0')}-${String(tmr.getDate()).padStart(2,'0')}`;

        const effectiveStart = startISO || todayISO;
        const effectiveDue   = dueISO   || (startISO ? startISO : tmrISO);

        if (new Date(effectiveStart) > new Date(effectiveDue)) {
            warnings.push(`"${rawName}": start date is after due date – swapped.`);
        }
        const finalStart = new Date(effectiveStart) <= new Date(effectiveDue) ? effectiveStart : effectiveDue;
        const finalEnd   = new Date(effectiveStart) <= new Date(effectiveDue) ? effectiveDue   : effectiveStart;

        // ── Progress / status ─────────────────────────────────────────────
        const pctRaw = rawPct !== '' ? parseProgress(rawPct) : parseProgress(rawProg);
        const pct    = pctRaw !== null ? pctRaw : null;
        const status = plannerProgressToStatus(rawProg, pct);
        const color  = (typeof STATUS_COLORS !== 'undefined' ? STATUS_COLORS[status] : null) || '#808080';

        // ── Group = Task Name ─────────────────────────────────────────────
        const groupName = rawName;

        if (!groups[groupName]) {
            groups[groupName]      = { name: groupName, color };
            groupStates[groupName] = true;
            groupOrder.push(groupName);
        }

        // ── Tasks = semicolon-split checklist items ────────────────────────
        // Split on ";" and strip whitespace; filter out empty tokens.
        const checklistItems = rawChecklist
            ? rawChecklist.split(';').map(s => s.trim()).filter(Boolean)
            : [];

        if (checklistItems.length > 0) {
            for (const itemName of checklistItems) {
                tasks.push({
                    id:        idBase++,
                    name:      itemName,
                    group:     groupName,
                    startDate: finalStart,
                    endDate:   finalEnd,
                    status,
                    progress:  pct,
                    color,
                });
                imported++;
            }
        } else {
            // No checklist → one task named after the group
            tasks.push({
                id:        idBase++,
                name:      groupName,
                group:     groupName,
                startDate: finalStart,
                endDate:   finalEnd,
                status,
                progress:  pct,
                color,
            });
            imported++;
        }

        if (typeof applyAutoGroupStatus === 'function') applyAutoGroupStatus(groupName);
    }

    return { imported, skipped, warnings };
}

// ─── File entry point ────────────────────────────────────────────────────────
/**
 * importFromPlanner(inputEl)
 * Called by the file input's onchange. Reads the chosen .xlsx/.csv/.xls file
 * and delegates to importPlannerRows().
 */
function importFromPlanner(input) {
    const file = input.files[0];
    if (!file) return;
    const name = file.name.toLowerCase();

    const onRows = (rows) => {
        const { imported, skipped, warnings } = importPlannerRows(rows);

        if (typeof markAsChanged === 'function') markAsChanged();
        if (typeof updateGroupsList === 'function') updateGroupsList();
        if (typeof updateChart === 'function') updateChart();
        if (typeof updateGroupSuggestions === 'function') updateGroupSuggestions();

        if (warnings.length) {
            console.warn('[Planner Import] Warnings:', warnings);
        }

        if (imported > 0) {
            const note = warnings.length ? ` (${warnings.length} warning${warnings.length>1?'s':''})` : '';
            if (typeof showNotification === 'function')
                showNotification(`Imported ${imported} task${imported !== 1 ? 's' : ''} from Planner${note}`, 'success');
        } else {
            if (typeof showNotification === 'function')
                showNotification('No tasks could be imported. ' + (warnings[0] || 'Check column names.'), 'error');
        }

        input.value = '';
    };

    if (name.endsWith('.csv')) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const rows = typeof parseCSV === 'function' ? parseCSV(e.target.result) : [];
                onRows(rows);
            } catch(err) {
                console.error(err);
                if (typeof showNotification === 'function') showNotification('Error reading CSV', 'error');
                input.value = '';
            }
        };
        reader.readAsText(file);
    } else {
        // xlsx / xls
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const wb   = XLSX.read(data, { type: 'array', cellDates: false });
                const ws   = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
                onRows(rows.map(r => ({ ...r })));
            } catch(err) {
                console.error(err);
                if (typeof showNotification === 'function') showNotification('Error reading Excel file', 'error');
                input.value = '';
            }
        };
        reader.readAsArrayBuffer(file);
    }
}
