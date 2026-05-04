/**
 * Export Gantt Chart to PowerPoint (.pptx)
 *
 * Slide structure:
 *   • 1+ Summary slides  — one row per group (auto-paginated)
 *   • 1+ Detail slides   — one row per task within each group (auto-paginated)
 *
 * Requires PptxGenJS loaded on the page.
 */
function exportToPPT() {
    if (!tasks || tasks.length === 0) {
        showNotification('No tasks to export', 'error');
        return;
    }

    showNotification('Building PPTX…', 'info');

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33" × 7.5"

    // ── Palette ──────────────────────────────────────────────────────────
    const BG      = '1A1A1A';
    const TITLE_B = 'E60000';
    const HDR_BG  = '252525';
    const ODD_BG  = '222222';
    const GRID_M  = '323232';
    const GRID_Q  = '7A0000';
    const TODAY_C = 'FF4444';
    const TXT_W   = 'FFFFFF';
    const TXT_DIM = '777777';
    const TXT_MUT = '3D3D3D';

    // ── Layout (inches) ──────────────────────────────────────────────────
    const W        = 13.33;
    const H        = 7.5;
    const LW       = 2.4;          // label column width
    const TX       = LW;            // timeline x origin
    const TW       = W - TX - 0.05; // timeline width
    const TITLE_H  = 0.45;
    const HDR_Y    = 0.50;
    const HDR_H    = 0.30;
    const ROW_Y0   = HDR_Y + HDR_H + 0.05;
    const ROW_H    = 0.34;
    const BAR_H    = 0.22;
    const BAR_VPAD = (ROW_H - BAR_H) / 2;
    const FOOTER_H = 0.26;
    // Reserve row 0 on detail slides for group summary; tasks fill rows 1..N
    const ROWS_PER_SLIDE = 14;

    // ── Timeline bounds ──────────────────────────────────────────────────
    const taskDates = tasks
        .flatMap(t => [parseYMD(t.startDate), parseYMD(t.endDate)])
        .filter(Boolean);
    const msDates = (milestones || []).map(m => parseYMD(m.date)).filter(Boolean);
    const allDates = [...taskDates, ...msDates];

    if (!allDates.length) {
        showNotification('No date data to export', 'error');
        return;
    }

    const minDate  = new Date(Math.min(...allDates));
    const maxDate  = new Date(Math.max(...allDates));
    const months   = getMonthRange(minDate, maxDate);
    const qGroups  = groupMonthsByQuarters(months);
    const totalM   = months.length;
    const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';

    const orderedGroups = (groupOrder.length ? groupOrder : Object.keys(groups))
        .filter(g => tasks.some(t => t.group === g));

    // ── Date → x position on the timeline ───────────────────────────────
    function dToX(raw) {
        const d = (raw instanceof Date) ? raw : parseYMD(raw);
        if (!d) return null;
        const y = d.getFullYear(), mo = d.getMonth() + 1, day = d.getDate();
        const idx = months.findIndex(m => m.year === y && m.month === mo);
        if (idx < 0) return null;
        const daysInMo = new Date(y, mo, 0).getDate();
        return TX + ((idx + (day - 1) / daysInMo) / totalM) * TW;
    }

    // ── Compute group-level summary (start, end, avg progress) ──────────
    function groupSummary(gName) {
        const gt = tasks.filter(t => t.group === gName);
        if (!gt.length) return null;
        const dates = gt.flatMap(t => [parseYMD(t.startDate), parseYMD(t.endDate)]).filter(Boolean);
        const gStart = new Date(Math.min(...dates));
        const gEnd   = new Date(Math.max(...dates));
        const withProg = gt.filter(t => typeof t.progress === 'number');
        const gPct = withProg.length
            ? Math.round(withProg.reduce((s, t) => s + t.progress, 0) / withProg.length)
            : Math.round(gt.filter(t => t.status === 'Completed').length / gt.length * 100);
        return { gStart, gEnd, gPct, count: gt.length };
    }

    // ── Draw slide chrome (bg, title bar, timeline header, grid, footer) ─
    function chrome(slide, title, subtitle) {
        const gridBottom = H - FOOTER_H - 0.04;

        // Background
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: 0, w: W, h: H,
            fill: { color: BG }, line: { color: BG }
        });

        // Title bar
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: 0, w: W, h: TITLE_H,
            fill: { color: TITLE_B }, line: { color: TITLE_B }
        });
        slide.addText(projName, {
            x: 0.16, y: 0, w: 7, h: TITLE_H,
            fontSize: 15, bold: true, color: TXT_W, fontFace: 'Segoe UI', valign: 'middle'
        });
        slide.addText(title, {
            x: 7.2, y: 0, w: W - 7.2 - 0.1, h: TITLE_H,
            fontSize: 12, color: 'FFBBBB', fontFace: 'Segoe UI', valign: 'middle', align: 'right'
        });

        // Timeline header background
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: HDR_Y, w: W, h: HDR_H,
            fill: { color: HDR_BG }, line: { color: HDR_BG }
        });

        // Header label column text
        if (subtitle) {
            slide.addText(subtitle, {
                x: 0.1, y: HDR_Y, w: LW - 0.15, h: HDR_H,
                fontSize: 7.5, color: TXT_DIM, fontFace: 'Segoe UI', valign: 'middle'
            });
        }

        // Quarter labels (top half of header)
        qGroups.forEach(q => {
            const qi = months.findIndex(m => m.year === q.months[0].year && m.month === q.months[0].month);
            if (qi < 0) return;
            const qx = TX + (qi / totalM) * TW;
            const qw = (q.months.length / totalM) * TW;
            slide.addText(`${q.fyLabel}  Q${q.quarter}`, {
                x: qx + 0.03, y: HDR_Y, w: qw - 0.06, h: HDR_H * 0.55,
                fontSize: 6.5, bold: true, color: 'FF6666', fontFace: 'Segoe UI', valign: 'middle'
            });
        });

        // Month labels (bottom half of header)
        months.forEach((m, i) => {
            const mx = TX + (i / totalM) * TW;
            const mw = (1 / totalM) * TW;
            slide.addText(getMonthShortName(m.month), {
                x: mx, y: HDR_Y + HDR_H * 0.55, w: mw, h: HDR_H * 0.45,
                fontSize: 6, color: TXT_DIM, fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
        });

        // Vertical grid lines
        months.slice(1).forEach((m, i) => {
            const gx = TX + ((i + 1) / totalM) * TW;
            const isQ = [4, 7, 10, 1].includes(m.month);
            slide.addShape(pptx.shapes.LINE, {
                x: gx, y: HDR_Y, w: 0, h: gridBottom - HDR_Y,
                line: { color: isQ ? GRID_Q : GRID_M, width: isQ ? 1.2 : 0.5, dashType: isQ ? 'solid' : 'lgDash' }
            });
        });

        // Today marker
        const todayX = dToX(new Date());
        if (todayX !== null && todayX >= TX && todayX <= TX + TW) {
            slide.addShape(pptx.shapes.LINE, {
                x: todayX, y: HDR_Y, w: 0, h: gridBottom - HDR_Y,
                line: { color: TODAY_C, width: 1.5, dashType: 'dash' }
            });
            slide.addText('TODAY', {
                x: todayX - 0.3, y: gridBottom + 0.01, w: 0.6, h: 0.16,
                fontSize: 5.5, bold: true, color: TODAY_C, fontFace: 'Segoe UI', align: 'center'
            });
        }

        // Milestones
        (milestones || []).forEach(ms => {
            const mx = dToX(ms.date);
            if (mx === null || mx < TX || mx > TX + TW) return;
            slide.addShape(pptx.shapes.DIAMOND, {
                x: mx - 0.09, y: HDR_Y - 0.01, w: 0.18, h: 0.18,
                fill: { color: 'E60000' }, line: { color: 'FFFFFF', width: 0.75 }
            });
            slide.addText(ms.name || '', {
                x: mx - 0.55, y: HDR_Y - 0.22, w: 1.1, h: 0.2,
                fontSize: 5.5, bold: true, color: 'FF8888', fontFace: 'Segoe UI', align: 'center'
            });
        });

        // Footer
        const ts = new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        slide.addText(`VOIS  ·  Generated ${ts}`, {
            x: 0, y: H - FOOTER_H, w: W, h: FOOTER_H,
            fontSize: 6.5, color: TXT_MUT, fontFace: 'Segoe UI', align: 'center', valign: 'middle'
        });
    }

    // ── Draw one Gantt row ────────────────────────────────────────────────
    function addRow(slide, ri, label, color, startDate, endDate, pct, isGroup) {
        const ry   = ROW_Y0 + ri * ROW_H;
        const hexC = color.replace('#', '');

        // Zebra stripe on odd rows
        if (ri % 2 === 0) {
            slide.addShape(pptx.shapes.RECTANGLE, {
                x: 0, y: ry, w: W, h: ROW_H,
                fill: { color: ODD_BG }, line: { color: ODD_BG }
            });
        }

        // Left colour accent strip for group rows
        if (isGroup) {
            slide.addShape(pptx.shapes.RECTANGLE, {
                x: 0, y: ry, w: 0.06, h: ROW_H,
                fill: { color: hexC }, line: { color: hexC }
            });
        }

        // Row label
        const lx = isGroup ? 0.12 : 0.22;
        slide.addText(label, {
            x: lx, y: ry, w: LW - lx - 0.05, h: ROW_H,
            fontSize: isGroup ? 9 : 8,
            bold: isGroup,
            color: isGroup ? TXT_W : 'CCCCCC',
            fontFace: 'Segoe UI',
            valign: 'middle',
            wrap: false
        });

        // Bar
        const x1 = dToX(startDate);
        const x2 = dToX(endDate);
        if (x1 === null || x2 === null) return;

        const bx = Math.min(x1, x2);
        const bw = Math.max(Math.abs(x2 - x1), 0.08);
        const by = ry + BAR_VPAD;

        slide.addShape(pptx.shapes.RECTANGLE, {
            x: bx, y: by, w: bw, h: BAR_H,
            fill: { color: hexC },
            line: { color: hexC }
        });

        // Progress fill (white semi-transparent overlay on left portion)
        if (pct !== null && pct > 0) {
            slide.addShape(pptx.shapes.RECTANGLE, {
                x: bx, y: by, w: bw * (pct / 100), h: BAR_H,
                fill: { color: 'FFFFFF', transparency: 65 },
                line: { color: 'FFFFFF', transparency: 100 }
            });
        }

        // Percentage label inside bar
        if (pct !== null && bw > 0.35) {
            slide.addText(`${pct}%`, {
                x: bx + 0.04, y: by, w: bw - 0.08, h: BAR_H,
                fontSize: isGroup ? 8 : 7,
                bold: true,
                color: 'FFFFFF',
                fontFace: 'Segoe UI',
                align: 'center',
                valign: 'middle'
            });
        }
    }

    // ════════════════════════════════════════════════════════════════════
    // SUMMARY SLIDE(S)  — one row per group, auto-paginated
    // ════════════════════════════════════════════════════════════════════
    const groupChunks = [];
    for (let i = 0; i < orderedGroups.length; i += ROWS_PER_SLIDE) {
        groupChunks.push(orderedGroups.slice(i, i + ROWS_PER_SLIDE));
    }

    groupChunks.forEach((chunk, ci) => {
        const slide = pptx.addSlide();
        const pg = groupChunks.length > 1 ? ` (${ci + 1}/${groupChunks.length})` : '';
        chrome(
            slide,
            `Summary${pg}`,
            `${orderedGroups.length} sprints  ·  ${tasks.length} tasks`
        );

        chunk.forEach((gName, ri) => {
            const s = groupSummary(gName);
            if (!s) return;
            const c = groups[gName]?.color || '#808080';
            addRow(slide, ri, gName, c, s.gStart, s.gEnd, s.gPct, true);

            // Task count badge to the right of the label
            slide.addText(`${s.count} tasks`, {
                x: LW - 0.72, y: ROW_Y0 + ri * ROW_H, w: 0.67, h: ROW_H,
                fontSize: 6.5, color: TXT_DIM, fontFace: 'Segoe UI', align: 'right', valign: 'middle'
            });
        });
    });

    // ════════════════════════════════════════════════════════════════════
    // DETAIL SLIDES  — one slide per group, tasks auto-paginated
    //   Row 0 = group summary bar
    //   Rows 1..N = individual tasks
    // ════════════════════════════════════════════════════════════════════
    const tasksPerPage = ROWS_PER_SLIDE - 1; // row 0 is always the group header

    orderedGroups.forEach(gName => {
        const gt = tasks.filter(t => t.group === gName);
        if (!gt.length) return;

        const gColor = groups[gName]?.color || '#808080';
        const s = groupSummary(gName);

        const taskChunks = [];
        for (let i = 0; i < gt.length; i += tasksPerPage) {
            taskChunks.push(gt.slice(i, i + tasksPerPage));
        }

        taskChunks.forEach((chunk, ci) => {
            const slide = pptx.addSlide();
            const pg = taskChunks.length > 1 ? `  —  page ${ci + 1}/${taskChunks.length}` : '';
            chrome(slide, `${gName}${pg}`, `${gt.length} tasks`);

            // Group summary row (always row 0)
            if (s) {
                addRow(slide, 0, gName, gColor, s.gStart, s.gEnd, s.gPct, true);
            }

            // Individual task rows
            chunk.forEach((task, ri) => {
                const tc = task.color || gColor;
                let pct = typeof task.progress === 'number' ? task.progress : null;
                if (pct === null) {
                    if (task.status === 'Completed')        pct = 100;
                    else if (task.status === 'Not Started') pct = 0;
                }
                addRow(slide, ri + 1, task.name, tc, task.startDate, task.endDate, pct, false);
            });
        });
    });

    // ── Save ─────────────────────────────────────────────────────────────
    const safeName = projName.replace(/[^\w\-]+/g, '_');
    const dateStr  = new Date().toISOString().slice(0, 10);
    pptx.writeFile({ fileName: `${safeName}_${dateStr}.pptx` })
        .then(() => showNotification('PPTX exported', 'success'))
        .catch(err => {
            console.error('PPTX export failed:', err);
            showNotification('PPTX export failed', 'error');
        });
}
