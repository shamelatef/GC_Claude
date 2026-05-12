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
        ['taskId',          ['taskid','id','outlinenumber','outline','wbs','outlinenum']],
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
 * Returns true when the file uses the native Planner hierarchy:
 *   whole-number IDs = parent tasks, decimal IDs (1.1, 2.3) = subtasks.
 */
function _hasPlannerHierarchy(rows, hm) {
    if (!hm.taskId) return false;
    return rows.some(r => /^\d+\.\d+/.test(String(r[hm.taskId] ?? '').trim()));
}

/**
 * @param {object[]} rows
 * @param {{ groupCol?: string, taskCol?: string }} [overrides]
 *   groupCol – raw header to use as group name (column-picker mode only)
 *   taskCol  – raw header to use as task/checklist (column-picker mode only)
 *
 * Two modes (auto-selected):
 *   Hierarchy mode  – Task ID column contains decimal IDs like 1.1, 2.3.
 *                     Whole-number ID rows → groups; decimal ID rows → tasks under last parent.
 *   Column-picker   – No Task ID hierarchy detected. Group column and task/checklist
 *                     column are chosen by the user (or auto-detected).
 *
 * Side-effects: mutates global `tasks`, `groups`, `groupStates`, `groupOrder`.
 */
function importPlannerRows(rows, overrides = {}) {
    if (!Array.isArray(rows) || rows.length === 0) {
        return { imported: 0, skipped: 0, warnings: ['No rows found in file.'] };
    }

    const headers = Object.keys(rows[0]);
    const hm      = buildPlannerHeaderMap(headers);
    const warnings = [];

    if (overrides.groupCol) hm.taskName      = overrides.groupCol;
    if (overrides.taskCol)  hm.checklistItem = overrides.taskCol;

    if (!hm.taskName) {
        return { imported: 0, skipped: 0, warnings: ['Could not find a Task Name column. Make sure you are importing the original Planner Excel export.'] };
    }

    const get = (row, field) => hm[field] ? String(row[hm[field]] ?? '').trim() : '';

    let imported = 0;
    let skipped  = 0;
    let idBase   = Date.now();

    // Shared helper: parse dates + status + color from one row.
    // Returns { startISO, dueISO, finalStart, finalEnd, status, pct, color, hasOwnDates, hasOwnProgress }
    const parseRow = (row, nameForWarning) => {
        const rawStart = get(row, 'startDate');
        const rawDue   = get(row, 'dueDate');
        const rawProg  = get(row, 'progress');
        const rawPct   = get(row, 'percentComplete');

        const startISO = parsePlannerDate(rawStart);
        const dueISO   = parsePlannerDate(rawDue);

        const today    = new Date();
        const todayISO = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
        const tmr      = new Date(today); tmr.setDate(tmr.getDate() + 1);
        const tmrISO   = `${tmr.getFullYear()}-${String(tmr.getMonth()+1).padStart(2,'0')}-${String(tmr.getDate()).padStart(2,'0')}`;

        const effectiveStart = startISO || todayISO;
        const effectiveDue   = dueISO   || (startISO ? startISO : tmrISO);

        if (nameForWarning && new Date(effectiveStart) > new Date(effectiveDue)) {
            warnings.push(`"${nameForWarning}": start date is after due date – swapped.`);
        }
        const finalStart = new Date(effectiveStart) <= new Date(effectiveDue) ? effectiveStart : effectiveDue;
        const finalEnd   = new Date(effectiveStart) <= new Date(effectiveDue) ? effectiveDue   : effectiveStart;

        const pctRaw = rawPct !== '' ? parseProgress(rawPct) : parseProgress(rawProg);
        const pct    = pctRaw !== null ? pctRaw : null;
        const status = plannerProgressToStatus(rawProg, pct);
        const color  = (typeof STATUS_COLORS !== 'undefined' ? STATUS_COLORS[status] : null) || '#808080';

        return {
            startISO, dueISO, finalStart, finalEnd, status, pct, color,
            hasOwnDates:    !!(startISO || dueISO),
            hasOwnProgress: rawProg !== '' || rawPct !== '',
        };
    };

    const pushTask = (name, group, p) => {
        tasks.push({ id: idBase++, name, group, startDate: p.finalStart, endDate: p.finalEnd, status: p.status, progress: p.pct, color: p.color });
        if (typeof applyAutoGroupStatus === 'function') applyAutoGroupStatus(group);
        imported++;
    };

    // Track which groups were added during this import and (in hierarchy mode) their numeric Task ID.
    const importStartIdx  = groupOrder.length;
    const groupTaskIdMap  = {};   // groupName → numeric Task ID (hierarchy mode only)

    const ensureGroup = (name, color, numericId) => {
        if (!groups[name]) {
            groups[name]      = { name, color };
            groupStates[name] = true;
            groupOrder.push(name);
        }
        if (numericId != null) groupTaskIdMap[name] = numericId;
    };

    // ── Mode 1: Task ID hierarchy (1 → group, 1.1 → subtask) ────────────────
    if (_hasPlannerHierarchy(rows, hm)) {
        let parent     = null;   // { name, ...parseRow result }
        let hasSubs    = false;

        const flushParent = () => {
            if (parent && !hasSubs) {
                // Parent with no subtasks → one fallback task
                pushTask(parent.name, parent.name, parent);
            }
        };

        for (const row of rows) {
            const rawId   = get(row, 'taskId');
            const rawName = get(row, 'taskName');
            if (!rawName) { skipped++; continue; }

            const isSubtask = /^\d+\.\d+/.test(rawId);

            if (!isSubtask) {
                // ── Parent row ─────────────────────────────────────────────
                flushParent();
                hasSubs = false;
                const p = parseRow(row, rawName);
                const numericId = parseFloat(rawId) || 0;
                ensureGroup(rawName, p.color, numericId);
                parent = { name: rawName, ...p };
            } else {
                // ── Subtask row ────────────────────────────────────────────
                if (!parent) { skipped++; continue; }

                const sub = parseRow(row, null);
                // Prefer subtask's own dates / status; fall back to parent's
                const p = {
                    finalStart: sub.hasOwnDates    ? sub.finalStart : parent.finalStart,
                    finalEnd:   sub.hasOwnDates    ? sub.finalEnd   : parent.finalEnd,
                    status:     sub.hasOwnProgress ? sub.status     : parent.status,
                    pct:        sub.hasOwnProgress ? sub.pct        : parent.pct,
                    color:      sub.hasOwnProgress ? sub.color      : parent.color,
                };
                pushTask(rawName, parent.name, p);
                hasSubs = true;
            }
        }
        flushParent();

        // Sort newly added groups by their numeric Task ID (ascending) so that
        // Task 1 appears before Task 2 regardless of Excel row order.
        const added = groupOrder.splice(importStartIdx);
        added.sort((a, b) => (groupTaskIdMap[a] || 0) - (groupTaskIdMap[b] || 0));
        groupOrder.push(...added);

    // ── Mode 2: Column-picker (semicolon checklist split) ────────────────────
    } else {
        for (const row of rows) {
            const rawName      = get(row, 'taskName');
            const rawChecklist = get(row, 'checklistItem');
            if (!rawName) { skipped++; continue; }

            const p = parseRow(row, rawName);
            ensureGroup(rawName, p.color);

            const items = rawChecklist
                ? rawChecklist.split(';').map(s => s.trim()).filter(Boolean)
                : [];

            if (items.length > 0) {
                for (const itemName of items) pushTask(itemName, rawName, p);
            } else {
                pushTask(rawName, rawName, p);
            }
        }

        // Planner exports newest-first; reverse so the oldest task appears at the top.
        const added = groupOrder.splice(importStartIdx);
        added.reverse();
        groupOrder.push(...added);
    }

    return { imported, skipped, warnings };
}

// ─── Column-picker modal state ───────────────────────────────────────────────
let _plannerPendingRows  = null;
let _plannerPendingInput = null;

/**
 * Populate the two dropdowns, pre-selecting the best guesses from the header map.
 */
function _populatePlannerDropdowns(headers, hm) {
    const groupSel = document.getElementById('plannerGroupCol');
    const taskSel  = document.getElementById('plannerTaskCol');
    if (!groupSel || !taskSel) return;

    groupSel.innerHTML = '';
    taskSel.innerHTML  = '<option value="">(none – use group name as task)</option>';

    for (const h of headers) {
        groupSel.appendChild(new Option(h, h));
        taskSel.appendChild(new Option(h, h));
    }

    // Smart defaults based on header map
    if (hm.taskName)      groupSel.value = hm.taskName;
    if (hm.checklistItem) taskSel.value  = hm.checklistItem;
}

function closePlannerColumnModal() {
    const modal = document.getElementById('plannerColumnModal');
    if (modal) modal.style.display = 'none';
    if (_plannerPendingInput) _plannerPendingInput.value = '';
    _plannerPendingRows  = null;
    _plannerPendingInput = null;
}

function confirmPlannerColumnImport() {
    const groupCol = document.getElementById('plannerGroupCol')?.value || '';
    const taskCol  = document.getElementById('plannerTaskCol')?.value  || '';

    const modal = document.getElementById('plannerColumnModal');
    if (modal) modal.style.display = 'none';

    const rows  = _plannerPendingRows;
    const input = _plannerPendingInput;
    _plannerPendingRows  = null;
    _plannerPendingInput = null;

    if (!rows) return;

    const { imported, skipped, warnings } = importPlannerRows(rows, { groupCol, taskCol });

    if (typeof markAsChanged === 'function') markAsChanged();
    if (typeof updateGroupsList === 'function') updateGroupsList();
    if (typeof updateChart === 'function') updateChart();
    if (typeof updateGroupSuggestions === 'function') updateGroupSuggestions();

    if (warnings.length) console.warn('[Planner Import] Warnings:', warnings);

    if (imported > 0) {
        const note = warnings.length ? ` (${warnings.length} warning${warnings.length>1?'s':''})` : '';
        if (typeof showNotification === 'function')
            showNotification(`Imported ${imported} task${imported !== 1 ? 's' : ''} from Planner${note}`, 'success');
    } else {
        if (typeof showNotification === 'function')
            showNotification('No tasks could be imported. ' + (warnings[0] || 'Check column names.'), 'error');
    }

    if (input) input.value = '';
}

// ─── Metadata-skip helper ────────────────────────────────────────────────────
/**
 * Given raw row arrays (from sheet_to_json with header:1), find the index of
 * the real column-header row by scoring each candidate row.
 *
 * A cell scores as a header if its normalised value STARTS WITH a known
 * header keyword (e.g. "task number" → "task", "start" → "start").
 * "Project name" starts with "project" (not in the keyword list) so it
 * scores 0, while row 9 of an MS-Project export scores 6.
 * The row with the highest score (minimum 2 matches) is chosen.
 */
function _findHeaderRowIndex(rawArrays) {
    const headerWords = new Set([
        'task','name','title','outline','start','finish','due','end',
        'assigned','assign','status','progress','wbs','id','checklist',
        'bucket','priority','complete','notes','description','percent',
    ]);
    const cellScore = (c) => {
        const norm = String(c ?? '').toLowerCase().trim();
        if (!norm) return 0;
        for (const w of headerWords) {
            if (norm === w || norm.startsWith(w + ' ') || norm.startsWith(w + '_')) return 1;
        }
        return 0;
    };

    let bestIdx   = 0;
    let bestScore = 0;
    for (let i = 0; i < Math.min(rawArrays.length, 20); i++) {
        const row     = rawArrays[i];
        const matches = row.reduce((s, c) => s + cellScore(c), 0);
        if (matches >= 2 && matches > bestScore) {
            bestScore = matches;
            bestIdx   = i;
        }
    }
    return bestIdx;
}

/**
 * Convert raw-array rows (from sheet_to_json header:1) into objects using the
 * header row at headerIdx, skipping blank rows.
 */
function _arraysToObjects(rawArrays, headerIdx) {
    const headers = rawArrays[headerIdx].map(c => String(c ?? ''));
    return rawArrays.slice(headerIdx + 1)
        .map(arr => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = arr[i] ?? ''; });
            return obj;
        })
        .filter(row => headers.some(h => h && String(row[h] ?? '').trim() !== ''));
}

// ─── File entry point ────────────────────────────────────────────────────────
/**
 * onPlannerFileSelected(inputEl)
 * Called by the file input's onchange. Reads the file, parses to rows,
 * then opens the column-picker modal so the user can choose which columns
 * map to groups and tasks before committing the import.
 */
function onPlannerFileSelected(input) {
    const file = input.files[0];
    if (!file) return;
    const name = file.name.toLowerCase();

    const onRows = (rows) => {
        if (!rows || rows.length === 0) {
            if (typeof showNotification === 'function') showNotification('No rows found in file.', 'error');
            input.value = '';
            return;
        }
        const headers = Object.keys(rows[0]);
        const hm      = buildPlannerHeaderMap(headers);

        // Native Planner hierarchy detected (Task IDs like 1, 1.1, 1.2) → auto-import
        if (_hasPlannerHierarchy(rows, hm)) {
            const { imported, skipped, warnings } = importPlannerRows(rows);
            if (typeof markAsChanged          === 'function') markAsChanged();
            if (typeof updateGroupsList       === 'function') updateGroupsList();
            if (typeof updateChart            === 'function') updateChart();
            if (typeof updateGroupSuggestions === 'function') updateGroupSuggestions();
            if (warnings.length) console.warn('[Planner Import] Warnings:', warnings);
            if (imported > 0) {
                const note = warnings.length ? ` (${warnings.length} warning${warnings.length>1?'s':''})` : '';
                if (typeof showNotification === 'function')
                    showNotification(`Imported ${imported} task${imported !== 1 ? 's' : ''} from Planner${note}`, 'success');
            } else {
                if (typeof showNotification === 'function')
                    showNotification('No tasks could be imported. ' + (warnings[0] || 'Check column names.'), 'error');
            }
            input.value = '';
            return;
        }

        // No hierarchy → show column-picker modal
        _plannerPendingRows  = rows;
        _plannerPendingInput = input;
        _populatePlannerDropdowns(headers, hm);
        const modal = document.getElementById('plannerColumnModal');
        if (modal) modal.style.display = 'flex';
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
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data      = new Uint8Array(e.target.result);
                const wb        = XLSX.read(data, { type: 'array', cellDates: false });
                const ws        = wb.Sheets[wb.SheetNames[0]];
                // Use raw arrays so we can detect and skip metadata rows (e.g. MS Project exports)
                const rawArrays = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
                const headerIdx = _findHeaderRowIndex(rawArrays);
                const rows      = _arraysToObjects(rawArrays, headerIdx);
                onRows(rows);
            } catch(err) {
                console.error(err);
                if (typeof showNotification === 'function') showNotification('Error reading Excel file', 'error');
                input.value = '';
            }
        };
        reader.readAsArrayBuffer(file);
    }
}
