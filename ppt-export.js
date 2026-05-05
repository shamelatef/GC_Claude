/**
 * Export Gantt Chart to PowerPoint — Swimlane layout matching reference design.
 *
 * Layout per slide:
 *   [VOIS red title bar]
 *   [Quarter markers row  — coloured shapes + labels + Today arrow]
 *   [Dark pill timeline   — month names inside]
 *   [Swimlane rows        — coloured label left | rounded task bars right]
 *   [Footer]
 *
 * Requires PptxGenJS (pptxgen.bundle.js) on the page.
 */
function exportToPPT() {
    if (!tasks || tasks.length === 0) {
        showNotification('No tasks to export', 'error');
        return;
    }
    showNotification('Building PPTX…', 'info');

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33" × 7.5"

    // ── Colours ──────────────────────────────────────────────────────────
    const C = {
        bg:       'FFFFFF',
        titleBg:  'E60000',        // VOIS red
        pill:     '2C3454',        // dark navy pill
        lane1:    'FFFFFF',
        lane2:    'F4F4F4',
        border:   'E2E2E2',
        today:    'E60000',
        txtW:     'FFFFFF',
        txtD:     '2C2C2C',
        txtMut:   'AAAAAA',
        // cycling colours for quarter marker shapes
        qShapes:  ['7B68C8', '3AAFA9', 'C0392B', '2980B9']
    };

    // ── Layout constants (inches) ─────────────────────────────────────────
    const W        = 13.33;
    const H        = 7.5;
    const TITLE_H  = 0.42;         // VOIS red bar

    const QTR_Y    = TITLE_H + 0.08;   // quarter marker row top
    const QTR_H    = 0.28;

    const PILL_Y   = QTR_Y + QTR_H + 0.04;  // dark pill top
    const PILL_H   = 0.38;

    const ROW_Y0   = PILL_Y + PILL_H + 0.06; // first swimlane top

    const BAR_H    = 0.22;
    const BAR_GAP  = 0.06;
    const ROW_PAD  = 0.10;         // top & bottom padding inside each swimlane
    const FOOTER_H = 0.20;
    const AVAIL_H  = H - ROW_Y0 - FOOTER_H - 0.04;

    // ── Timeline data ────────────────────────────────────────────────────
    const taskDates = tasks.flatMap(t => [parseYMD(t.startDate), parseYMD(t.endDate)]).filter(Boolean);
    const msDates   = (milestones || []).map(m => parseYMD(m.date)).filter(Boolean);
    const allDates  = [...taskDates, ...msDates];
    if (!allDates.length) { showNotification('No date data to export', 'error'); return; }

    const minDate  = new Date(Math.min(...allDates));
    const maxDate  = new Date(Math.max(...allDates));
    const months   = getMonthRange(minDate, maxDate);
    const qGroups  = groupMonthsByQuarters(months);
    const totalM   = months.length;
    const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';

    // Groups that actually have tasks, in display order
    const activeGroups = (groupOrder.length ? groupOrder : Object.keys(groups))
        .filter(g => tasks.some(t => t.group === g));
    const showLabels = activeGroups.length > 1;
    const LW = showLabels ? 1.65 : 0;   // label column width
    const TX = LW;
    const TW = W - TX - 0.05;

    // ── Helpers ───────────────────────────────────────────────────────────
    function dToX(raw) {
        const d = (raw instanceof Date) ? raw : parseYMD(raw);
        if (!d) return null;
        const y = d.getFullYear(), mo = d.getMonth() + 1, day = d.getDate();
        const idx = months.findIndex(m => m.year === y && m.month === mo);
        if (idx < 0) return null;
        const daysInMo = new Date(y, mo, 0).getDate();
        return TX + ((idx + (day - 1) / daysInMo) / totalM) * TW;
    }

    function rowH(n) {
        // Height tall enough to fit n stacked bars with padding top & bottom
        return Math.max(0.56, ROW_PAD * 2 + n * (BAR_H + BAR_GAP) - BAR_GAP);
    }

    // Split groups across slides so nothing overflows vertically
    function paginate() {
        const pages = [];
        let page = [], used = 0;
        for (const g of activeGroups) {
            const rh = rowH(tasks.filter(t => t.group === g).length);
            if (used + rh > AVAIL_H && page.length > 0) {
                pages.push(page); page = [g]; used = rh;
            } else {
                page.push(g); used += rh;
            }
        }
        if (page.length) pages.push(page);
        return pages;
    }

    // ── Slide chrome (title bar + quarter row + pill + footer) ───────────
    function chrome(slide, pageLabel) {
        // White background
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: 0, w: W, h: H,
            fill: { color: C.bg }, line: { color: C.bg }
        });

        // ── VOIS red title bar ──
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: 0, w: W, h: TITLE_H,
            fill: { color: C.titleBg }, line: { color: C.titleBg }
        });
        slide.addText(projName, {
            x: 0.15, y: 0, w: 8.5, h: TITLE_H,
            fontSize: 16, bold: true, color: C.txtW,
            fontFace: 'Segoe UI', valign: 'middle'
        });
        if (pageLabel) {
            slide.addText(pageLabel, {
                x: 8.7, y: 0, w: W - 8.7 - 0.1, h: TITLE_H,
                fontSize: 10, color: 'FFBBBB',
                fontFace: 'Segoe UI', valign: 'middle', align: 'right'
            });
        }

        // ── Quarter marker shapes + labels (above the pill) ──
        qGroups.forEach((q, qi) => {
            const si = months.findIndex(m => m.year === q.months[0].year && m.month === q.months[0].month);
            if (si < 0) return;
            const qx     = TX + (si / totalM) * TW;
            const colour  = C.qShapes[qi % C.qShapes.length];
            const shapeH  = 0.16;
            const shapeY  = QTR_Y + (QTR_H - shapeH) / 2;

            // Alternate circle → rounded-square → diamond
            if (qi % 3 === 0) {
                slide.addShape(pptx.shapes.OVAL, {
                    x: qx - shapeH / 2, y: shapeY, w: shapeH, h: shapeH,
                    fill: { color: colour }, line: { color: colour }
                });
            } else if (qi % 3 === 1) {
                slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                    x: qx - shapeH / 2, y: shapeY, w: shapeH, h: shapeH,
                    rectRadius: 0.2,
                    fill: { color: colour }, line: { color: colour }
                });
            } else {
                slide.addShape(pptx.shapes.DIAMOND, {
                    x: qx - shapeH / 2, y: shapeY, w: shapeH, h: shapeH,
                    fill: { color: colour }, line: { color: colour }
                });
            }

            // Quarter label to the right of the shape
            slide.addText(`${q.fyLabel} Q${q.quarter}`, {
                x: qx + shapeH / 2 + 0.04, y: QTR_Y, w: 1.6, h: QTR_H,
                fontSize: 6.5, bold: true, color: '555555',
                fontFace: 'Segoe UI', valign: 'middle'
            });
        });

        // ── Milestones in quarter row ──
        (milestones || []).forEach(ms => {
            const mx = dToX(ms.date);
            if (mx === null || mx < TX || mx > TX + TW) return;
            const shapeH = 0.14;
            slide.addShape(pptx.shapes.OVAL, {
                x: mx - shapeH / 2, y: QTR_Y + (QTR_H - shapeH) / 2,
                w: shapeH, h: shapeH,
                fill: { color: '555555' }, line: { color: C.txtW, width: 0.5 }
            });
            slide.addText(ms.name || '', {
                x: mx - 0.55, y: QTR_Y - 0.22, w: 1.1, h: 0.2,
                fontSize: 5.5, color: '555555', fontFace: 'Segoe UI',
                align: 'center', bold: true
            });
        });

        // ── TODAY label + downward arrow above pill ──
        const todayX = dToX(new Date());
        if (todayX !== null && todayX >= TX && todayX <= TX + TW) {
            slide.addText('Today', {
                x: todayX - 0.3, y: QTR_Y, w: 0.6, h: QTR_H * 0.65,
                fontSize: 7.5, bold: true, color: C.today,
                fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
            // Downward-pointing arrow ▼ just above the pill
            slide.addText('▼', {
                x: todayX - 0.14, y: QTR_Y + QTR_H * 0.62, w: 0.28, h: 0.16,
                fontSize: 9, bold: true, color: C.today,
                fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
            // Dashed vertical line through content
            slide.addShape(pptx.shapes.LINE, {
                x: todayX, y: PILL_Y, w: 0, h: H - PILL_Y - FOOTER_H,
                line: { color: C.today, width: 1.2, dashType: 'dash' }
            });
        }

        // ── Dark pill — month names ──
        // Outer pill shape
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: TX, y: PILL_Y, w: TW, h: PILL_H,
            rectRadius: 0.5,
            fill: { color: C.pill }, line: { color: C.pill }
        });

        // Month name cells inside the pill
        const today = new Date();
        months.forEach((m, i) => {
            const mx  = TX + (i / totalM) * TW;
            const mw  = (1 / totalM) * TW;
            const isNow = today.getMonth() + 1 === m.month && today.getFullYear() === m.year;

            // Subtle highlight for current month (slightly lighter inner pill)
            if (isNow) {
                slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                    x: mx + 0.03, y: PILL_Y + 0.05, w: mw - 0.06, h: PILL_H - 0.1,
                    rectRadius: 0.5,
                    fill: { color: 'E60000', transparency: 25 },
                    line: { color: 'E60000', transparency: 25 }
                });
            }

            slide.addText(getMonthShortName(m.month), {
                x: mx, y: PILL_Y, w: mw, h: PILL_H,
                fontSize: 10, bold: true, color: C.txtW,
                fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
        });

        // ── Footer ──
        const ts = new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        slide.addText(`VOIS  ·  Generated ${ts}`, {
            x: 0, y: H - FOOTER_H, w: W, h: FOOTER_H,
            fontSize: 6, color: C.txtMut,
            fontFace: 'Segoe UI', align: 'center', valign: 'middle'
        });
    }

    // ── Draw one swimlane row ─────────────────────────────────────────────
    function drawSwimlane(slide, gName, groupTasks, rowY, rh, laneIdx) {
        const gHex = (groups[gName]?.color || '#808080').replace('#', '');

        // Alternating row background
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: rowY, w: W, h: rh,
            fill: { color: laneIdx % 2 === 0 ? C.lane1 : C.lane2 },
            line: { color: laneIdx % 2 === 0 ? C.lane1 : C.lane2 }
        });

        // Bottom separator line
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: rowY + rh - 0.005, w: W, h: 0.005,
            fill: { color: C.border }, line: { color: C.border }
        });

        // ── Coloured label box on the left ──
        if (showLabels) {
            slide.addShape(pptx.shapes.RECTANGLE, {
                x: 0, y: rowY, w: LW, h: rh,
                fill: { color: gHex }, line: { color: gHex }
            });
            slide.addText(gName, {
                x: 0.07, y: rowY + 0.04, w: LW - 0.12, h: rh - 0.08,
                fontSize: rh > 0.75 ? 9.5 : 8.5,
                bold: true, color: C.txtW, fontFace: 'Segoe UI',
                valign: 'middle', align: 'center', wrap: true
            });
        }

        // ── Rounded pill bars for each task ──
        groupTasks.forEach((task, ti) => {
            const tHex = (task.color || '#808080').replace('#', '');
            const x1   = dToX(task.startDate);
            const x2   = dToX(task.endDate);
            if (x1 === null || x2 === null) return;

            const bx  = Math.min(x1, x2);
            const bw  = Math.max(Math.abs(x2 - x1), 0.1);
            const by  = rowY + ROW_PAD + ti * (BAR_H + BAR_GAP);

            // Pill bar
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: bx, y: by, w: bw, h: BAR_H,
                rectRadius: 0.5,
                fill: { color: tHex }, line: { color: tHex }
            });

            // Progress overlay (white fill on left portion)
            let pct = typeof task.progress === 'number' ? task.progress : null;
            if (pct === null) {
                if (task.status === 'Completed')        pct = 100;
                else if (task.status === 'Not Started') pct = 0;
            }
            if (pct !== null && pct > 0 && pct < 100) {
                slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                    x: bx, y: by, w: bw * (pct / 100), h: BAR_H,
                    rectRadius: 0.5,
                    fill: { color: 'FFFFFF', transparency: 60 },
                    line: { color: 'FFFFFF', transparency: 100 }
                });
            }

            // Task name inside bar if wide enough, else to the right as grey text
            const label = task.name || '';
            if (bw >= 0.45) {
                slide.addText(label, {
                    x: bx + 0.07, y: by, w: bw - 0.14, h: BAR_H,
                    fontSize: 7.5, bold: true, color: C.txtW,
                    fontFace: 'Segoe UI', valign: 'middle', wrap: false
                });
            } else {
                slide.addText(label, {
                    x: bx + bw + 0.06, y: by, w: 1.4, h: BAR_H,
                    fontSize: 7, color: C.txtD,
                    fontFace: 'Segoe UI', valign: 'middle', wrap: false
                });
            }
        });
    }

    // ════════════════════════════════════════════════════════════════════
    // BUILD SLIDES
    // ════════════════════════════════════════════════════════════════════
    const pages = paginate();

    pages.forEach((pageGroups, pi) => {
        const slide = pptx.addSlide();
        const lbl   = pages.length > 1 ? `Page ${pi + 1} / ${pages.length}` : '';
        chrome(slide, lbl);

        let curY = ROW_Y0;
        pageGroups.forEach((gName, li) => {
            const gt = tasks.filter(t => t.group === gName);
            const rh = rowH(gt.length);
            drawSwimlane(slide, gName, gt, curY, rh, li);
            curY += rh;
        });
    });

    // ── Save ──────────────────────────────────────────────────────────────
    const safeName = projName.replace(/[^\w\-]+/g, '_');
    pptx.writeFile({ fileName: `${safeName}_${new Date().toISOString().slice(0, 10)}.pptx` })
        .then(() => showNotification('PPTX exported', 'success'))
        .catch(err => { console.error(err); showNotification('PPTX export failed', 'error'); });
}
