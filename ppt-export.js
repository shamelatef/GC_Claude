/**
 * Export Gantt Chart to PowerPoint — Office Timeline style.
 *
 * Layout per slide (16:9, 13.33" × 7.5"):
 *
 *   ┌──────────────────────────────────────────────────────┐
 *   │ SECTION 1   ▶───────────▶   May 19 – Jun 2           │
 *   │              ◇ Sub-task ◇  ◇ Sub-task 2 ◇            │
 *   │     SECTION 2     ▶─────▶  Jun 3 – Jun 17            │
 *   │                                                       │
 *   │ ┌──────── Week 1 │ 2 │ 3 │ 4 │ 5 │ 6 │ 7 ────────┐   │
 *   │ 2025                                                  │
 *   │       ★    ◆      ◆           ◆               ★      │
 *   │     Mon May 19   Fri May 23  Tue Jun 3   …            │
 *   │     Milestone 1  Milestone 2  Milestone 3 …           │
 *   └──────────────────────────────────────────────────────┘
 *
 * Auto-paginates when the section + task list overflows.
 * Auto-picks week / month / quarter axis cells based on duration.
 *
 * Requires PptxGenJS (pptxgen.bundle.js) and the following globals from
 * the host page: tasks, milestones, groups, groupOrder, projects,
 * activeProjectIndex, parseYMD, showNotification.
 */
function exportToPPT(mode) {
    // mode: 'tasks' (default) or 'tasks-with-subtasks'
    mode = mode === 'tasks-with-subtasks' ? 'tasks-with-subtasks' : 'tasks';
    const includeSubtasks = mode === 'tasks-with-subtasks';

    if (!tasks || tasks.length === 0) {
        showNotification('No tasks to export', 'error');
        return;
    }
    showNotification('Building PPTX…', 'info');

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33" × 7.5"

    // ── Slide constants (inches) ──────────────────────────────────────────
    const W = 13.33, H = 7.5;
    const MARGIN_L = 0.7, MARGIN_R = 0.3, MARGIN_T = 0.45, MARGIN_B = 0.35;
    const LABEL_COL_W = 1.45;   // left column for SECTION / sub-task names
    const SIDE_COL_W  = 1.30;   // right column for date ranges / overflow names

    const TX = MARGIN_L + LABEL_COL_W;
    const TR = W - MARGIN_R - SIDE_COL_W;
    const TW = TR - TX;

    const PILL_H   = 0.32;
    const MS_AREA  = 1.05;       // milestone area below the pill
    const PILL_TOP = H - MARGIN_B - MS_AREA - PILL_H;
    const CONTENT_TOP = MARGIN_T + 0.05;
    const CONTENT_BOT = PILL_TOP - 0.18;
    const CONTENT_H   = CONTENT_BOT - CONTENT_TOP;

    // ── Colors ────────────────────────────────────────────────────────────
    const C = {
        bg:        'FFFFFF',
        pill:      '475569',
        guide:     'D8DDE3',
        ink:       '1A1A1A',
        muted:     '9AA0A6',
        msLine:    'CFD4DA',
        year:      'E58A3A',
        sideText:  '374151',
    };

    // ── Date helpers ──────────────────────────────────────────────────────
    function ymd(s) { return (s instanceof Date) ? s : parseYMD(s); }
    function days(a, b) { return Math.round((b - a) / 86400000); }
    function startOfWeekMon(d) {
        const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
        const wd = (x.getDay() + 6) % 7;
        x.setDate(x.getDate() - wd);
        return x;
    }
    function fmtMD(d) {
        const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return `${m[d.getMonth()]} ${d.getDate()}`;
    }
    function fmtDOWMD(d) {
        const dow = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
        const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return `${dow[d.getDay()]} ${m[d.getMonth()]} ${d.getDate()}`;
    }

    // ── Date range ────────────────────────────────────────────────────────
    const taskDates = tasks.flatMap(t => [ymd(t.startDate), ymd(t.endDate)]).filter(Boolean);
    const msDates = (milestones || []).map(m => ymd(m.date)).filter(Boolean);
    const allDates = [...taskDates, ...msDates];
    if (!allDates.length) { showNotification('No date data to export', 'error'); return; }

    const minDate = new Date(Math.min(...allDates));
    const maxDate = new Date(Math.max(...allDates));

    // ── Pick axis unit based on duration ──────────────────────────────────
    const totalDays = Math.max(1, days(minDate, maxDate));
    let unit;
    if (totalDays <= 80) unit = 'week';
    else if (totalDays <= 400) unit = 'month';
    else unit = 'quarter';

    // ── Build axis cells ──────────────────────────────────────────────────
    function buildAxis() {
        const cells = [];
        if (unit === 'week') {
            const start = startOfWeekMon(minDate);
            let cur = new Date(start), i = 1;
            while (cur <= maxDate) {
                const next = new Date(cur); next.setDate(next.getDate() + 7);
                cells.push({ label: i === 1 ? 'Week 1' : String(i), start: new Date(cur), end: new Date(next) });
                cur = next; i++;
            }
            return { axisStart: start, axisEnd: cells[cells.length - 1].end, cells };
        }
        if (unit === 'month') {
            const start = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
            const mn = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            let cur = new Date(start);
            while (cur <= maxDate) {
                const next = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
                cells.push({ label: mn[cur.getMonth()], start: new Date(cur), end: new Date(next) });
                cur = next;
            }
            return { axisStart: start, axisEnd: cells[cells.length - 1].end, cells };
        }
        // quarter
        const qStartMo = Math.floor(minDate.getMonth() / 3) * 3;
        const start = new Date(minDate.getFullYear(), qStartMo, 1);
        let cur = new Date(start);
        while (cur <= maxDate) {
            const next = new Date(cur.getFullYear(), cur.getMonth() + 3, 1);
            const q = Math.floor(cur.getMonth() / 3) + 1;
            cells.push({ label: `Q${q} ${cur.getFullYear()}`, start: new Date(cur), end: new Date(next) });
            cur = next;
        }
        return { axisStart: start, axisEnd: cells[cells.length - 1].end, cells };
    }
    const axis = buildAxis();
    const axisMs = axis.axisEnd - axis.axisStart;

    function dToX(raw) {
        const d = ymd(raw); if (!d) return null;
        const t = Math.max(axis.axisStart.getTime(), Math.min(axis.axisEnd.getTime(), d.getTime()));
        return TX + ((t - axis.axisStart.getTime()) / axisMs) * TW;
    }

    // ── Active groups in display order ────────────────────────────────────
    const activeGroups = (groupOrder.length ? groupOrder : Object.keys(groups))
        .filter(g => tasks.some(t => t.group === g));

    // ── Pagination ────────────────────────────────────────────────────────
    // Each section contributes 1 header row + N task rows.
    // Pick row height that fits everything if possible; otherwise paginate.
    // Subtask rows count toward total only when included.
    function rowsForTask(t) {
        if (!includeSubtasks) return 1;
        const subs = Array.isArray(t.subtasks) ? t.subtasks : [];
        return 1 + subs.length;
    }
    function rowsForGroup(g) {
        // 1 (section header) + sum of task rows
        return 1 + tasks.filter(t => t.group === g)
                        .reduce((acc, t) => acc + rowsForTask(t), 0);
    }

    function planLayout() {
        const totalRows = activeGroups.reduce((acc, g) => acc + rowsForGroup(g), 0);

        const ROW_MIN = 0.18;
        const ROW_MAX = 0.42;
        const fitH = CONTENT_H / Math.max(1, totalRows);
        const rowH = Math.max(ROW_MIN, Math.min(ROW_MAX, fitH));
        const rowsPerPage = Math.floor(CONTENT_H / rowH);

        // Build a flat list of "units" per group, where one unit is
        // {task, includeSubs:bool} representing 1 or (1 + subtasks) rows.
        // We paginate by group -> by units, splitting only at unit boundaries
        // so a parent task and its subtasks stay together when possible.
        const pages = [];
        let page = []; // [{group, units:[unitIdx...], isContinuation}]
        let used = 0;

        for (const g of activeGroups) {
            const gTasks = tasks.filter(t => t.group === g);
            const units = gTasks.map(t => ({
                task: t,
                rows: rowsForTask(t)
            }));
            let unitFrom = 0;
            while (unitFrom < units.length || unitFrom === 0) {
                if (unitFrom >= units.length && unitFrom > 0) break;
                // section header costs 1 row
                let room = rowsPerPage - used - 1;
                if (room < 1 && (page.length > 0 || used > 0)) {
                    pages.push(page); page = []; used = 0;
                    continue;
                }
                // Take as many full units as fit; if a single unit is bigger
                // than the page itself, force-take it (and accept overflow).
                let taken = 0, takenRows = 0;
                while (unitFrom + taken < units.length) {
                    const ur = units[unitFrom + taken].rows;
                    if (takenRows + ur > room && taken > 0) break;
                    takenRows += ur;
                    taken++;
                    if (takenRows >= room) break;
                }
                // Empty group with no tasks: still emit the header row
                page.push({
                    group: g,
                    unitFrom,
                    unitTo: unitFrom + taken,
                    units,
                    isContinuation: unitFrom > 0
                });
                used += 1 + takenRows;
                unitFrom += taken;
                if (unitFrom >= units.length) break;
                // section continues
                pages.push(page); page = []; used = 0;
            }
        }
        if (page.length > 0) pages.push(page);
        return { pages, rowH };
    }

    const { pages, rowH } = planLayout();

    // ── Slide chrome (axis pill + year + milestones) ──────────────────────
    function chrome(slide, pageLabel) {
        // White background
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: 0, w: W, h: H,
            fill: { color: C.bg }, line: { color: C.bg }
        });

        // Project title (small, top-left)
        const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';
        slide.addText(projName, {
            x: MARGIN_L, y: 0.05, w: 7, h: 0.30,
            fontSize: 13, bold: true, color: C.ink,
            fontFace: 'Segoe UI', valign: 'middle'
        });
        if (pageLabel) {
            slide.addText(pageLabel, {
                x: W - 1.6, y: 0.05, w: 1.4, h: 0.30,
                fontSize: 10, color: C.muted,
                fontFace: 'Segoe UI', valign: 'middle', align: 'right'
            });
        }

        // Brand watermark on far left, rotated
        slide.addText('Made with', {
            x: -0.05, y: (CONTENT_TOP + CONTENT_BOT)/2 - 0.1, w: 0.5, h: 0.2,
            fontSize: 9, color: C.muted, rotate: -90,
            fontFace: 'Segoe UI', align: 'center', valign: 'middle'
        });

        // Vertical guide lines per axis cell
        axis.cells.forEach((c, i) => {
            if (i === 0) return;
            const gx = dToX(c.start);
            slide.addShape(pptx.shapes.LINE, {
                x: gx, y: CONTENT_TOP - 0.05, w: 0, h: PILL_TOP - (CONTENT_TOP - 0.05),
                line: { color: C.guide, width: 0.75 }
            });
        });

        // Dark navy pill
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: TX, y: PILL_TOP, w: TW, h: PILL_H,
            rectRadius: 0.06,
            fill: { color: C.pill }, line: { color: C.pill }
        });
        // Pill cells
        const cellW = TW / axis.cells.length;
        axis.cells.forEach((c, i) => {
            const cx = TX + i * cellW;
            // Cell divider
            if (i > 0) {
                slide.addShape(pptx.shapes.LINE, {
                    x: cx, y: PILL_TOP + 0.04, w: 0, h: PILL_H - 0.08,
                    line: { color: 'FFFFFF', width: 0.6, transparency: 60 }
                });
            }
            slide.addText(c.label, {
                x: cx, y: PILL_TOP, w: cellW, h: PILL_H,
                fontSize: i === 0 ? 11 : 12, bold: true, color: 'FFFFFF',
                fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
        });

        // Year label, hanging off the left of the pill
        slide.addText(String(axis.axisStart.getFullYear()), {
            x: MARGIN_L - 0.15, y: PILL_TOP - 0.02, w: LABEL_COL_W, h: PILL_H + 0.04,
            fontSize: 22, bold: true, color: C.year,
            fontFace: 'Segoe UI', align: 'left', valign: 'middle'
        });

        // ── Milestones below the pill, staggered ──
        const msTop = PILL_TOP + PILL_H + 0.14;
        const placed = []; // {x, lane}
        const minSpacing = 1.25; // inches
        (milestones || []).forEach(m => {
            const d = ymd(m.date);
            if (!d || d < axis.axisStart || d > axis.axisEnd) return;
            const mx = dToX(d);
            let lane = 0;
            while (placed.some(p => p.lane === lane && Math.abs(p.x - mx) < minSpacing)) lane++;
            placed.push({ x: mx, lane });
            const baseY = msTop + lane * 0.46;

            // Connector line up from milestone to bottom of pill
            slide.addShape(pptx.shapes.LINE, {
                x: mx, y: PILL_TOP + PILL_H, w: 0, h: baseY - (PILL_TOP + PILL_H),
                line: { color: C.msLine, width: 0.75, dashType: 'dash' }
            });

            // Marker (diamond or star)
            const sz = 0.16;
            const color = (m.color || '#3B8FD9').replace('#', '');
            if (m.shape === 'star') {
                slide.addShape(pptx.shapes.STAR_5, {
                    x: mx - sz/2, y: baseY - sz/2, w: sz, h: sz,
                    fill: { color }, line: { color }
                });
            } else {
                slide.addShape(pptx.shapes.DIAMOND, {
                    x: mx - sz/2, y: baseY - sz/2, w: sz, h: sz,
                    fill: { color }, line: { color }
                });
            }

            // Date and name (two lines below marker)
            slide.addText(fmtDOWMD(d), {
                x: mx - 0.7, y: baseY + sz/2 + 0.02, w: 1.4, h: 0.18,
                fontSize: 8, color: C.muted, fontFace: 'Segoe UI',
                align: 'center', valign: 'middle'
            });
            slide.addText(m.name || '', {
                x: mx - 0.7, y: baseY + sz/2 + 0.18, w: 1.4, h: 0.20,
                fontSize: 9.5, bold: true, color: C.ink, fontFace: 'Segoe UI',
                align: 'center', valign: 'middle'
            });
        });

        // Footer
        const ts = new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        slide.addText(`VOIS  ·  Generated ${ts}`, {
            x: 0, y: H - 0.22, w: W, h: 0.2,
            fontSize: 7, color: C.muted,
            fontFace: 'Segoe UI', align: 'center', valign: 'middle'
        });
    }

    // ── Draw a section row (header + arrow + date range) ──────────────────
    function drawSection(slide, gName, isCont, y, rh) {
        const groupColor = (groups[gName]?.color || '#475569').replace('#', '');
        const gTasks = tasks.filter(t => t.group === gName);
        if (!gTasks.length) return;
        const gStart = new Date(Math.min(...gTasks.map(t => ymd(t.startDate))));
        const gEnd   = new Date(Math.max(...gTasks.map(t => ymd(t.endDate))));
        const x1 = dToX(gStart), x2 = dToX(gEnd);
        const arrowH = Math.min(0.34, rh * 0.78);
        const ay = y + (rh - arrowH) / 2;
        const aw = Math.max(0.5, x2 - x1);

        // Right-pointing pentagon arrow for the section
        slide.addShape(pptx.shapes.PENTAGON, {
            x: x1, y: ay, w: aw, h: arrowH,
            fill: { color: groupColor }, line: { color: groupColor }
        });

        // Label on the LEFT (section name) — right-aligned to butt up to arrow
        const labelText = (gName + (isCont ? ' (cont.)' : '')).toUpperCase();
        slide.addText(labelText, {
            x: MARGIN_L, y: y, w: LABEL_COL_W - 0.08, h: rh,
            fontSize: Math.min(12, rh * 28),
            bold: true, color: C.ink, fontFace: 'Segoe UI',
            align: 'right', valign: 'middle', wrap: false, shrinkText: true
        });

        // Date range to the RIGHT of the arrow tip, in the section color
        const drText = `${fmtMD(gStart)} – ${fmtMD(gEnd)}`;
        slide.addText(drText, {
            x: x2 + 0.08, y: y, w: (W - MARGIN_R) - (x2 + 0.08), h: rh,
            fontSize: Math.min(11, rh * 26), bold: true,
            color: groupColor, fontFace: 'Segoe UI',
            align: 'left', valign: 'middle', wrap: false
        });
    }

    // ── Draw a sub-task row ───────────────────────────────────────────────
    function drawTask(slide, task, y, rh, isFirstInGroup) {
        const tColor = (task.color || '#475569').replace('#', '');
        const ts = ymd(task.startDate), te = ymd(task.endDate);
        const x1 = dToX(ts), x2 = dToX(te);
        if (x1 === null || x2 === null) return;
        const arrowH = Math.min(0.24, rh * 0.62);
        const ay = y + (rh - arrowH) / 2;
        const w = Math.max(0, x2 - x1);
        const isZero = w < 0.04; // < ~3px — render diamond instead of arrow

        // Approx character width in inches at the chosen font size
        const fs = Math.min(10, arrowH * 36);
        const charW = fs * 0.0075; // very rough but works for Segoe UI
        const padIn = 0.16;
        const minWForInside = task.name.length * charW + padIn;

        if (isZero) {
            // Diamond marker
            const sz = Math.min(0.16, arrowH);
            slide.addShape(pptx.shapes.DIAMOND, {
                x: x1 - sz/2, y: y + (rh - sz)/2, w: sz, h: sz,
                fill: { color: tColor }, line: { color: tColor }
            });
            // Name to the right
            slide.addText(task.name || '', {
                x: x1 + sz/2 + 0.04, y: y, w: (W - MARGIN_R) - (x1 + sz/2 + 0.04), h: rh,
                fontSize: Math.min(10, rh * 24), color: C.sideText,
                fontFace: 'Segoe UI', bold: true,
                align: 'left', valign: 'middle', wrap: false
            });
            return;
        }

        // Use CHEVRON for sub-tasks if not first (gives <— shape on the left),
        // PENTAGON if first or arrow is small. Both are right-pointing.
        const shape = isFirstInGroup
            ? pptx.shapes.PENTAGON
            : pptx.shapes.CHEVRON;
        slide.addShape(shape, {
            x: x1, y: ay, w: Math.max(0.32, w), h: arrowH,
            fill: { color: tColor }, line: { color: tColor }
        });

        // Text inside if it fits, else to the right
        if (w >= minWForInside && w >= 0.5) {
            slide.addText(task.name || '', {
                x: x1 + 0.08, y: ay, w: Math.max(0.16, w - 0.24), h: arrowH,
                fontSize: fs, bold: true, color: 'FFFFFF',
                fontFace: 'Segoe UI', align: 'center', valign: 'middle', wrap: false
            });
        } else {
            slide.addText(task.name || '', {
                x: x2 + 0.08, y: y, w: (W - MARGIN_R) - (x2 + 0.08), h: rh,
                fontSize: Math.min(10, rh * 24), color: C.sideText,
                fontFace: 'Segoe UI', bold: true,
                align: 'left', valign: 'middle', wrap: false
            });
        }
    }

    // ── Draw a subtask row (checklist item under a parent task) ──────────
    function drawSubtask(slide, parentTask, sub, y, rh) {
        const pColor = (parentTask.color || '#475569').replace('#', '');
        const ts = ymd(parentTask.startDate), te = ymd(parentTask.endDate);
        const x1 = dToX(ts), x2 = dToX(te);
        if (x1 === null || x2 === null) return;
        const w = Math.max(0, x2 - x1);

        // Thin track for subtask, slightly inset, lighter feel
        const trackH = Math.max(0.06, Math.min(0.10, rh * 0.32));
        const ty = y + (rh - trackH) / 2;
        const inset = Math.min(0.12, w * 0.08);
        const tx = x1 + inset;
        const tw = Math.max(0.10, w - inset * 2);

        // Light-tinted track in the parent's color
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: tx, y: ty, w: tw, h: trackH,
            rectRadius: 0.04,
            fill: { color: pColor, transparency: 70 },
            line: { color: pColor, transparency: 50, width: 0.5 }
        });

        // Status marker on the LEFT — filled dot if done, ring if not
        const dotSize = Math.min(0.16, rh * 0.55);
        const dotX = x1 - dotSize - 0.04;
        const dotY = y + (rh - dotSize) / 2;
        if (sub.done) {
            slide.addShape(pptx.shapes.OVAL, {
                x: dotX, y: dotY, w: dotSize, h: dotSize,
                fill: { color: pColor }, line: { color: pColor }
            });
            slide.addText('✓', {
                x: dotX, y: dotY, w: dotSize, h: dotSize,
                fontSize: Math.max(7, dotSize * 36),
                bold: true, color: 'FFFFFF',
                fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
        } else {
            slide.addShape(pptx.shapes.OVAL, {
                x: dotX, y: dotY, w: dotSize, h: dotSize,
                fill: { color: 'FFFFFF' },
                line: { color: pColor, width: 1 }
            });
        }

        // Subtask name — inside the track if there is room, else to the right.
        const fs = Math.min(9.5, rh * 22);
        const charW = fs * 0.0075;
        const minWForInside = (sub.name || '').length * charW + 0.18;
        if (tw >= minWForInside && tw >= 0.6) {
            slide.addText(sub.name || '', {
                x: tx + 0.06, y: y, w: tw - 0.12, h: rh,
                fontSize: fs,
                color: sub.done ? C.muted : pColor,
                bold: !sub.done, italic: !!sub.done, strike: !!sub.done,
                fontFace: 'Segoe UI', align: 'left', valign: 'middle', wrap: false
            });
        } else {
            slide.addText(sub.name || '', {
                x: x2 + 0.08, y: y, w: (W - MARGIN_R) - (x2 + 0.08), h: rh,
                fontSize: fs,
                color: sub.done ? C.muted : C.sideText,
                italic: !!sub.done, strike: !!sub.done,
                fontFace: 'Segoe UI', align: 'left', valign: 'middle', wrap: false
            });
        }
    }

    // ── Build slides ──────────────────────────────────────────────────────
    pages.forEach((pageRows, pi) => {
        const slide = pptx.addSlide();
        const lbl = pages.length > 1 ? `Page ${pi + 1} / ${pages.length}` : '';
        chrome(slide, lbl);

        let y = CONTENT_TOP;
        pageRows.forEach(({ group, unitFrom, unitTo, units, isContinuation }) => {
            // Section row
            drawSection(slide, group, isContinuation, y, rowH);
            y += rowH;

            for (let ui = unitFrom; ui < unitTo; ui++) {
                const t = units[ui].task;
                drawTask(slide, t, y, rowH, ui === unitFrom);
                y += rowH;

                if (includeSubtasks && Array.isArray(t.subtasks) && t.subtasks.length > 0) {
                    t.subtasks.forEach(st => {
                        drawSubtask(slide, t, st, y, rowH);
                        y += rowH;
                    });
                }
            }
        });
    });

    // ── Save ──────────────────────────────────────────────────────────────
    const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';
    const safeName = projName.replace(/[^\w\-]+/g, '_');
    const modeSuffix = includeSubtasks ? '_with-subtasks' : '_tasks';
    pptx.writeFile({ fileName: `${safeName}${modeSuffix}_${new Date().toISOString().slice(0, 10)}.pptx` })
        .then(() => showNotification('PPTX exported', 'success'))
        .catch(err => { console.error(err); showNotification('PPTX export failed', 'error'); });
}
