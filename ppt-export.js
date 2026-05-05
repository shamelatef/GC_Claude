/**
 * Export Gantt Chart to PowerPoint (.pptx) — Swimlane layout
 *
 * Structure:
 *   - Each GROUP becomes a horizontal swimlane row (label on left, task bars inside)
 *   - TASKS within a group stack vertically as coloured bars in the timeline
 *   - task.subtasks checklist items are ignored (task-level only)
 *   - If there is only one group OR all tasks share the same group name, the left
 *     label column is hidden and bars span the full width
 *   - Slides are auto-paginated: groups are split across slides when they won't fit
 *
 * Requires PptxGenJS (pptxgen.bundle.js) loaded on the page.
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
    const BG       = 'FFFFFF';
    const TITLE_BG = 'E60000';
    const HDR_BG   = '1F2937';   // dark navy month-header bar
    const LANE_BG2 = 'F5F5F5';   // alternating swimlane stripe
    const LANE_BG1 = 'FFFFFF';
    const BORDER   = 'DEDEDE';
    const GRID_M   = 'E8E8E8';
    const GRID_Q   = 'CCCCCC';
    const TODAY_C  = 'E60000';
    const TXT_W    = 'FFFFFF';
    const TXT_D    = '1A1A1A';
    const TXT_MUT  = '999999';

    // ── Layout (inches) ──────────────────────────────────────────────────
    const W        = 13.33;
    const H        = 7.5;
    const TITLE_H  = 0.42;
    const QTR_Y    = TITLE_H + 0.05;   // quarter-label row
    const QTR_H    = 0.24;
    const HDR_Y    = QTR_Y + QTR_H;    // month-bar top
    const HDR_H    = 0.34;
    const ROW_Y0   = HDR_Y + HDR_H;    // first swimlane top
    const BAR_H    = 0.21;
    const BAR_GAP  = 0.05;
    const ROW_PAD  = 0.08;             // vertical padding inside each swimlane
    const FOOTER_H = 0.22;
    const AVAIL_H  = H - ROW_Y0 - FOOTER_H - 0.02;

    // ── Determine whether to show the left label column ──────────────────
    const activeGroupNames = (groupOrder.length ? groupOrder : Object.keys(groups))
        .filter(g => tasks.some(t => t.group === g));
    const showLabels = activeGroupNames.length > 1;
    const LW = showLabels ? 1.75 : 0;
    const TX = LW;
    const TW = W - TX - 0.04;

    // ── Timeline bounds ──────────────────────────────────────────────────
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

    // ── Date → x position in the timeline ───────────────────────────────
    function dToX(raw) {
        const d = (raw instanceof Date) ? raw : parseYMD(raw);
        if (!d) return null;
        const y = d.getFullYear(), mo = d.getMonth() + 1, day = d.getDate();
        const idx = months.findIndex(m => m.year === y && m.month === mo);
        if (idx < 0) return null;
        const daysInMo = new Date(y, mo, 0).getDate();
        return TX + ((idx + (day - 1) / daysInMo) / totalM) * TW;
    }

    // ── Row height for a swimlane based on its task count ────────────────
    function rowH(n) {
        return Math.max(0.52, ROW_PAD * 2 + n * (BAR_H + BAR_GAP) - BAR_GAP);
    }

    // ── Paginate groups so each slide fits within available height ────────
    function paginate() {
        const pages = [];
        let page = [], used = 0;
        for (const g of activeGroupNames) {
            const n  = tasks.filter(t => t.group === g).length;
            const rh = rowH(n);
            if (used + rh > AVAIL_H && page.length > 0) {
                pages.push(page);
                page = [g];
                used = rh;
            } else {
                page.push(g);
                used += rh;
            }
        }
        if (page.length) pages.push(page);
        return pages;
    }

    // ── Draw slide chrome ─────────────────────────────────────────────────
    function chrome(slide, pageLabel) {
        const contentBottom = H - FOOTER_H;

        // White background
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: 0, w: W, h: H,
            fill: { color: BG }, line: { color: BG }
        });

        // VOIS Red title bar
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: 0, w: W, h: TITLE_H,
            fill: { color: TITLE_BG }, line: { color: TITLE_BG }
        });
        slide.addText(projName, {
            x: 0.15, y: 0, w: 8.5, h: TITLE_H,
            fontSize: 16, bold: true, color: TXT_W, fontFace: 'Segoe UI', valign: 'middle'
        });
        if (pageLabel) {
            slide.addText(pageLabel, {
                x: 8.7, y: 0, w: W - 8.7 - 0.1, h: TITLE_H,
                fontSize: 10, color: 'FFBBBB', fontFace: 'Segoe UI', valign: 'middle', align: 'right'
            });
        }

        // Quarter labels row (light grey text above the month bar)
        qGroups.forEach((q, qi) => {
            const startIdx = months.findIndex(m =>
                m.year === q.months[0].year && m.month === q.months[0].month
            );
            if (startIdx < 0) return;
            const qx = TX + (startIdx / totalM) * TW;
            const qw = (q.months.length / totalM) * TW;

            // Quarter label text
            slide.addText(`${q.fyLabel}  Q${q.quarter}`, {
                x: qx + 0.03, y: QTR_Y, w: qw - 0.06, h: QTR_H,
                fontSize: 7, bold: true, color: '888888',
                fontFace: 'Segoe UI', valign: 'middle'
            });

            // Dashed quarter boundary line down through the content
            if (qi > 0) {
                slide.addShape(pptx.shapes.LINE, {
                    x: qx, y: QTR_Y, w: 0, h: contentBottom - QTR_Y,
                    line: { color: GRID_Q, width: 0.8, dashType: 'lgDash' }
                });
            }
        });

        // Label column header (dark box, same height as month bar)
        if (showLabels) {
            slide.addShape(pptx.shapes.RECTANGLE, {
                x: 0, y: HDR_Y, w: LW, h: HDR_H,
                fill: { color: '2D2D2D' }, line: { color: '2D2D2D' }
            });
            slide.addText('Sprint / Group', {
                x: 0.1, y: HDR_Y, w: LW - 0.15, h: HDR_H,
                fontSize: 8, bold: true, color: TXT_W, fontFace: 'Segoe UI', valign: 'middle'
            });
        }

        // Month header bar (dark navy)
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: TX, y: HDR_Y, w: TW, h: HDR_H,
            fill: { color: HDR_BG }, line: { color: HDR_BG }
        });

        // Month name labels + highlight today's month
        months.forEach((m, i) => {
            const mx = TX + (i / totalM) * TW;
            const mw = (1 / totalM) * TW;
            const isNow = new Date().getMonth() + 1 === m.month && new Date().getFullYear() === m.year;
            if (isNow) {
                slide.addShape(pptx.shapes.RECTANGLE, {
                    x: mx, y: HDR_Y, w: mw, h: HDR_H,
                    fill: { color: TODAY_C }, line: { color: TODAY_C }
                });
            }
            slide.addText(getMonthShortName(m.month), {
                x: mx, y: HDR_Y, w: mw, h: HDR_H,
                fontSize: 9, bold: true, color: TXT_W,
                fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
        });

        // Vertical month grid lines through content area
        months.slice(1).forEach((m, i) => {
            const gx = TX + ((i + 1) / totalM) * TW;
            slide.addShape(pptx.shapes.LINE, {
                x: gx, y: ROW_Y0, w: 0, h: contentBottom - ROW_Y0,
                line: { color: GRID_M, width: 0.5 }
            });
        });

        // Today vertical marker
        const todayX = dToX(new Date());
        if (todayX !== null && todayX >= TX && todayX <= TX + TW) {
            slide.addShape(pptx.shapes.LINE, {
                x: todayX, y: QTR_Y, w: 0, h: contentBottom - QTR_Y,
                line: { color: TODAY_C, width: 1.5, dashType: 'dash' }
            });
            slide.addText('Today', {
                x: todayX - 0.28, y: QTR_Y, w: 0.56, h: QTR_H,
                fontSize: 7, bold: true, color: TODAY_C, fontFace: 'Segoe UI', align: 'center', valign: 'middle'
            });
        }

        // Milestones (shown in the quarter row as diamonds + label)
        (milestones || []).forEach(ms => {
            const mx = dToX(ms.date);
            if (mx === null || mx < TX || mx > TX + TW) return;
            slide.addShape(pptx.shapes.DIAMOND, {
                x: mx - 0.085, y: QTR_Y + (QTR_H - 0.15) / 2, w: 0.15, h: 0.15,
                fill: { color: '333333' }, line: { color: TXT_W, width: 0.5 }
            });
            slide.addText(ms.name || '', {
                x: mx - 0.55, y: QTR_Y - 0.18, w: 1.1, h: 0.2,
                fontSize: 5.5, bold: true, color: '555555', fontFace: 'Segoe UI', align: 'center'
            });
        });

        // Footer
        const ts = new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        slide.addText(`VOIS  ·  Generated ${ts}`, {
            x: 0, y: H - FOOTER_H, w: W, h: FOOTER_H,
            fontSize: 6.5, color: TXT_MUT, fontFace: 'Segoe UI', align: 'center', valign: 'middle'
        });
    }

    // ── Draw one swimlane row ─────────────────────────────────────────────
    function drawSwimlane(slide, gName, groupTasks, rowY, rh, laneIdx) {
        const gColor = (groups[gName]?.color || '#808080').replace('#', '');

        // Alternating row background
        const laneBg = laneIdx % 2 === 0 ? LANE_BG1 : LANE_BG2;
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: rowY, w: W, h: rh,
            fill: { color: laneBg }, line: { color: laneBg }
        });

        // Bottom border
        slide.addShape(pptx.shapes.RECTANGLE, {
            x: 0, y: rowY + rh - 0.006, w: W, h: 0.006,
            fill: { color: BORDER }, line: { color: BORDER }
        });

        // Left label cell (only if showLabels)
        if (showLabels) {
            slide.addShape(pptx.shapes.RECTANGLE, {
                x: 0, y: rowY, w: LW, h: rh,
                fill: { color: gColor }, line: { color: gColor }
            });
            slide.addText(gName, {
                x: 0.08, y: rowY + 0.02, w: LW - 0.14, h: rh - 0.04,
                fontSize: rh > 0.7 ? 9.5 : 8.5,
                bold: true, color: TXT_W, fontFace: 'Segoe UI',
                valign: 'middle', align: 'center', wrap: true
            });
        }

        // Task bars
        groupTasks.forEach((task, ti) => {
            const tc  = (task.color || '#808080').replace('#', '');
            const x1  = dToX(task.startDate);
            const x2  = dToX(task.endDate);
            if (x1 === null || x2 === null) return;

            const bx  = Math.min(x1, x2);
            const bw  = Math.max(Math.abs(x2 - x1), 0.1);
            const by  = rowY + ROW_PAD + ti * (BAR_H + BAR_GAP);

            // Bar background
            slide.addShape(pptx.shapes.RECTANGLE, {
                x: bx, y: by, w: bw, h: BAR_H,
                fill: { color: tc }, line: { color: tc }
            });

            // Progress fill (white overlay)
            let pct = typeof task.progress === 'number' ? task.progress : null;
            if (pct === null) {
                if (task.status === 'Completed')        pct = 100;
                else if (task.status === 'Not Started') pct = 0;
            }
            if (pct !== null && pct > 0) {
                slide.addShape(pptx.shapes.RECTANGLE, {
                    x: bx, y: by, w: bw * (pct / 100), h: BAR_H,
                    fill: { color: 'FFFFFF', transparency: 65 },
                    line: { color: 'FFFFFF', transparency: 100 }
                });
            }

            // Task name inside bar if wide enough, else to the right
            const label = task.name || '';
            if (bw >= 0.5) {
                slide.addText(label, {
                    x: bx + 0.05, y: by, w: bw - 0.1, h: BAR_H,
                    fontSize: 7.5, bold: true, color: TXT_W,
                    fontFace: 'Segoe UI', valign: 'middle', wrap: false
                });
            } else if (bx + bw + 0.6 < W) {
                // Label to the right of a narrow bar
                slide.addText(label, {
                    x: bx + bw + 0.04, y: by, w: 1.0, h: BAR_H,
                    fontSize: 7, color: TXT_D, fontFace: 'Segoe UI', valign: 'middle', wrap: false
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
        const pageLabel = pages.length > 1 ? `Page ${pi + 1} / ${pages.length}` : '';
        chrome(slide, pageLabel);

        let curY = ROW_Y0;
        pageGroups.forEach((gName, laneIdx) => {
            const gt = tasks.filter(t => t.group === gName);
            const rh = rowH(gt.length);
            drawSwimlane(slide, gName, gt, curY, rh, laneIdx);
            curY += rh;
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
