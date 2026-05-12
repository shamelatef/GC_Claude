/**
 * Export Gantt Chart to PowerPoint
 *
 * Timeline pill matches the web-app header exactly:
 *   Row 1 (top)    → Fiscal Year  "FY 26/27"  — dark navy bg, white text
 *   Row 2 (middle) → Quarter      "Q1"         — medium bg, red text
 *   Row 3 (bottom) → Month+Year   "May 26"     — slate bg, white text
 *
 * For week-unit axes a single-row pill is used (weeks don't map to FY).
 * For quarter-unit axes a 2-row pill is used (FY + Quarter).
 *
 * Fiscal year (matches web app exactly):
 *   FY starts April 1.
 *   Q1 = Apr–Jun  |  Q2 = Jul–Sep  |  Q3 = Oct–Dec  |  Q4 = Jan–Mar
 *
 * Bottom-left year label:
 *   Single calendar year  →  "20" stacked above "26"  (large orange)
 *   Multiple years        →  start YY stacked above end YY  (e.g. "26" / "27")
 *
 * mode === 'groups-and-tasks'  →  compact group headers + task chevrons
 * mode === 'groups-only'       →  one summary chevron per group
 * mode === 'mixed'             →  expanded groups show tasks; others show summary bar
 */
function exportToPPT(mode, expandedGroups) {
    const VALID = ['groups-and-tasks', 'groups-only', 'mixed'];
    if (!VALID.includes(mode)) mode = 'groups-and-tasks';

    if (!tasks || tasks.length === 0) {
        showNotification('No tasks to export', 'error');
        return;
    }
    showNotification('Building PPTX…', 'info');

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 × 7.5"

    // ── Fixed slide geometry (inches) ────────────────────────────────────
    const W        = 13.33, H = 7.5;
    const MARGIN_L = 0.40;
    const MARGIN_R = 0.30;
    const MARGIN_T = 0.42;
    const MARGIN_B = 0.35;
    const GX       = MARGIN_L;   // left edge of every row

    // Name column + timeline zone
    const NAME_COL_W = 1.80;   // fixed width of the task-label column (inches)
    const NAME_GAP   = 0.13;   // gap between name column right-edge and first bar
    const TX     = GX + NAME_COL_W + NAME_GAP;  // timeline always starts here
    const SIDE_W = 1.15;       // right column for date labels
    const TR     = W - MARGIN_R - SIDE_W;
    const TW     = TR - TX;

    // ── Multi-row pill heights ────────────────────────────────────────────
    const ROW_FY_H  = 0.22;   // Fiscal Year label row
    const ROW_QT_H  = 0.22;   // Quarter label row
    const ROW_MO_H  = 0.26;   // Month+Year label row (also used for single-row week pill)

    // Total pill area height depends on unit (computed after axis is built)
    const MS_AREA   = 0.80;
    // PILL_TOP and CONTENT_H are computed once we know the unit

    // ── Design tokens ─────────────────────────────────────────────────────
    const C = {
        bg:          'FFFFFF',
        ink:         '1A1A1A',
        // Pill rows
        pillFY:      '2B3C4E',   // darkest  — FY row bg
        pillQT:      '3D4F63',   // medium   — Quarter row bg
        pillMO:      '475569',   // lightest — Month row bg
        pillWeek:    '3D4A5C',   // single-row week pill
        pillDivider: 'FFFFFF',   // cell divider lines
        qtText:      'E53E3E',   // red quarter labels (matches web-app)
        // Group header
        groupHdr:    'EEF2FF',
        groupBorder: 'C4D0E8',
        groupAccent: '3D4A5C',
        groupText:   '1E3A5F',
        // Misc
        guide:       'E2E8F0',
        year:        'E58A3A',   // orange year label
        dateText:    '8A9299',   // muted grey date labels
        msLine:      'CBD5E1',
    };

    // ── Helpers ───────────────────────────────────────────────────────────
    const ACCENT_W   = 0.11;
    // Group header text column: starts after accent bar + inner padding
    const GRP_TEXT_X = GX + ACCENT_W + 0.10;
    const GRP_TEXT_W = NAME_COL_W - ACCENT_W - 0.10;   // ~1.59" — never touches timeline
    const GRP_FS     = 11;
    const TASK_FS    = 10;
    const GRP_H      = 0.32;   // fixed group header height — always single-line
    const TASK_H_1   = 0.42;   // task row: name fits in one line
    const TASK_H_2   = 0.60;   // task row: name needs two lines

    function estW(text, pt) {
        const up = text === text.toUpperCase() && /[A-Z]/.test(text);
        return text.length * pt * (up ? 0.0061 : 0.0052);
    }

    // Returns 0-100 group progress (matches web-app logic exactly)
    function groupProgress(g) {
        const gt = tasks.filter(t => t.group === g);
        if (!gt.length) return null;
        const withPct = gt.filter(t => typeof t.progress === 'number');
        if (withPct.length) {
            return Math.round(withPct.reduce((s, t) => s + t.progress, 0) / withPct.length);
        }
        const done = gt.filter(t => t.status === 'Completed').length;
        return Math.round((done / gt.length) * 100);
    }

    // Blend a hex colour toward white by `factor` (0 = original, 1 = white)
    function lightenHex(hex, factor) {
        const r = parseInt(hex.slice(0,2),16), g = parseInt(hex.slice(2,4),16), b = parseInt(hex.slice(4,6),16);
        const b2 = v => Math.round(v + (255-v)*factor).toString(16).padStart(2,'0');
        return b2(r)+b2(g)+b2(b);
    }

    // Draws a CHEVRON with progress fill on the timeline bar.
    // When pct is known: light background chevron + solid fill rectangle overlay.
    // When pct is null:  single solid chevron (original behaviour).
    function drawChevronWithProgress(slide, x, y, w, h, color, pct) {
        const minW = Math.max(0.28, w);
        if (pct === null || pct >= 100) {
            // Solid bar — no split needed
            const col = pct >= 100 ? color : color;
            slide.addShape(pptx.shapes.CHEVRON, {
                x, y, w:minW, h, fill:{color}, line:{color}
            });
        } else {
            // 1 — Light background chevron (empty portion)
            const bgColor = lightenHex(color, 0.55);
            slide.addShape(pptx.shapes.CHEVRON, {
                x, y, w:minW, h, fill:{color:bgColor}, line:{color:bgColor}
            });
            // 2 — Solid fill rectangle (filled portion), inset slightly so chevron edge shows
            if (pct > 0) {
                const fillW = Math.max(0.02, minW * pct / 100);
                slide.addShape(pptx.shapes.RECTANGLE, {
                    x, y:y + 0.02, w:fillW, h:h - 0.04,
                    fill:{color}, line:{color}
                });
            }
        }
    }

    // ── Date helpers ──────────────────────────────────────────────────────
    function ymd(s)  { return (s instanceof Date) ? s : parseYMD(s); }
    function days(a, b) { return Math.round((b - a) / 86400000); }
    function weekMon(d) {
        const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
        x.setDate(x.getDate() - (x.getDay() + 6) % 7);
        return x;
    }
    const MN  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const DOW = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    function fmtMD(d)    { return `${MN[d.getMonth()]} ${d.getDate()}`; }
    function fmtMDY(d)   { return `${MN[d.getMonth()]} ${d.getDate()} ${d.getFullYear()}`; }
    function fmtDOWMD(d) { return `${DOW[d.getDay()]} ${MN[d.getMonth()]} ${d.getDate()}`; }
    // Date range: omit year when both dates share the same year
    function fmtRange(d1, d2) {
        return d1.getFullYear() === d2.getFullYear()
            ? `${fmtMD(d1)} – ${fmtMD(d2)}`
            : `${fmtMDY(d1)} – ${fmtMDY(d2)}`;
    }

    // ── Fiscal-year helpers (identical to web-app groupMonthsByQuarters) ──
    // FY starts April. Q1=Apr–Jun, Q2=Jul–Sep, Q3=Oct–Dec, Q4=Jan–Mar
    function toFY(y) { return String(y).slice(-2).padStart(2, '0'); }
    function fyInfoFor(calYear, calMonth) {  // calMonth: 1..12
        const isAfterMarch = calMonth >= 4;
        const fyStartYear  = isAfterMarch ? calYear : calYear - 1;
        const fyEndYear    = fyStartYear + 1;
        const fyLabel      = `FY ${toFY(fyStartYear)}/${toFY(fyEndYear)}`;
        let quarter;
        if      (calMonth >= 4  && calMonth <= 6)  quarter = 1;
        else if (calMonth >= 7  && calMonth <= 9)  quarter = 2;
        else if (calMonth >= 10 && calMonth <= 12) quarter = 3;
        else                                        quarter = 4;
        return { fyStartYear, fyEndYear, fyLabel, quarter,
                 fyKey: `${fyStartYear}-Q${quarter}` };
    }

    // ── Global date range ──────────────────────────────────────────────────
    const allDates = [
        ...tasks.flatMap(t => [ymd(t.startDate), ymd(t.endDate)]),
        ...(milestones || []).map(m => ymd(m.date))
    ].filter(Boolean);
    if (!allDates.length) { showNotification('No date data', 'error'); return; }
    const minDate = new Date(Math.min(...allDates));
    const maxDate = new Date(Math.max(...allDates));

    // ── Axis ──────────────────────────────────────────────────────────────
    const totalDays = Math.max(1, days(minDate, maxDate));
    const unit = totalDays <= 400 ? 'month' : 'quarter';

    function buildAxis() {
        const cells = [];

        if (unit === 'month') {
            const start = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
            let cur = new Date(start);
            while (cur <= maxDate) {
                const next = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
                const m = cur.getMonth() + 1, y = cur.getFullYear();
                const fy = fyInfoFor(y, m);
                cells.push({
                    label: `${MN[m-1]} ${toFY(y)}`,   // e.g. "May 26"
                    month: m, year: y,
                    start: new Date(cur), end: new Date(next),
                    ...fy
                });
                cur = next;
            }
            return { axisStart:start, axisEnd:cells[cells.length-1].end, cells };
        }

        // quarter unit — still enrich with FY info
        const qm    = Math.floor(minDate.getMonth() / 3) * 3;
        const start = new Date(minDate.getFullYear(), qm, 1);
        let cur = new Date(start);
        while (cur <= maxDate) {
            const next = new Date(cur.getFullYear(), cur.getMonth() + 3, 1);
            const m = cur.getMonth() + 1, y = cur.getFullYear();
            const fy = fyInfoFor(y, m);
            cells.push({
                label: `Q${fy.quarter}`,
                month: m, year: y,
                start: new Date(cur), end: new Date(next),
                ...fy
            });
            cur = next;
        }
        return { axisStart:start, axisEnd:cells[cells.length-1].end, cells };
    }

    const axis   = buildAxis();
    const axisMs = axis.axisEnd - axis.axisStart;

    // Pill total height depends on unit
    const PILL_AREA_H = unit === 'month'
        ? ROW_FY_H + ROW_QT_H + ROW_MO_H   // 3 rows: FY / Quarter / Month
        : ROW_FY_H + ROW_QT_H;             // 2 rows: FY / Quarter

    const PILL_TOP    = H - MARGIN_B - MS_AREA - PILL_AREA_H;
    const CONTENT_TOP = MARGIN_T + 0.04;
    const CONTENT_BOT = PILL_TOP - 0.14;
    const CONTENT_H   = CONTENT_BOT - CONTENT_TOP;

    function dToX(raw) {
        const d = ymd(raw); if (!d) return null;
        const t = Math.max(axis.axisStart.getTime(), Math.min(axis.axisEnd.getTime(), d.getTime()));
        return TX + ((t - axis.axisStart.getTime()) / axisMs) * TW;
    }

    // ── Ordered groups ─────────────────────────────────────────────────────
    const ordered = (groupOrder.length ? groupOrder : Object.keys(groups))
        .filter(g => tasks.some(t => t.group === g));

    // ══════════════════════════════════════════════════════════════════════
    // CHROME
    // ══════════════════════════════════════════════════════════════════════
    function chrome(slide, pageLabel) {
        // White background
        slide.addShape(pptx.shapes.RECTANGLE, {
            x:0, y:0, w:W, h:H, fill:{color:C.bg}, line:{color:C.bg}
        });

        // Title
        const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';
        slide.addText(projName, {
            x:GX, y:0.06, w:8, h:0.28,
            fontSize:13, bold:true, color:C.ink, fontFace:'Segoe UI', valign:'middle'
        });
        if (pageLabel) {
            slide.addText(pageLabel, {
                x:W - 1.50, y:0.06, w:1.30, h:0.28,
                fontSize:9, color:C.dateText, fontFace:'Segoe UI', valign:'middle', align:'right'
            });
        }

        // ── Multi-row timeline pill ───────────────────────────────────────
        if (unit === 'month') {
            // ── Row 1: Fiscal Year ──────────────────────────────────────
            const fyRowY = PILL_TOP;
            // Build unique FY groups in order
            const fyGroups = [];
            const fyMap = new Map();
            axis.cells.forEach(c => {
                if (!fyMap.has(c.fyLabel)) {
                    fyMap.set(c.fyLabel, { fyLabel:c.fyLabel, cells:[] });
                    fyGroups.push(fyMap.get(c.fyLabel));
                }
                fyMap.get(c.fyLabel).cells.push(c);
            });
            // Rounded rect spanning full pill (to get rounded corners only on top)
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x:TX, y:fyRowY, w:TW, h:PILL_AREA_H, rectRadius:0.05,
                fill:{color:C.pillFY}, line:{color:C.pillFY}
            });
            // FY group labels + vertical dividers
            fyGroups.forEach((fg, gi) => {
                const x1 = dToX(fg.cells[0].start);
                const x2 = dToX(fg.cells[fg.cells.length - 1].end);
                const fw  = Math.max(0, x2 - x1);
                if (gi > 0) {
                    slide.addShape(pptx.shapes.LINE, {
                        x:x1, y:fyRowY + 0.03, w:0, h:ROW_FY_H - 0.06,
                        line:{color:C.pillDivider, width:0.6, transparency:50}
                    });
                }
                slide.addText(fg.fyLabel, {
                    x:x1, y:fyRowY, w:Math.max(0.10, fw), h:ROW_FY_H,
                    fontSize:9, bold:true, color:'FFFFFF',
                    fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
                });
            });

            // ── Row 2: Quarter ──────────────────────────────────────────
            const qtRowY = PILL_TOP + ROW_FY_H;
            slide.addShape(pptx.shapes.RECTANGLE, {
                x:TX, y:qtRowY, w:TW, h:ROW_QT_H,
                fill:{color:C.pillQT}, line:{color:C.pillQT}
            });
            // Build unique quarter groups
            const qtGroups = [];
            const qtMap = new Map();
            axis.cells.forEach(c => {
                if (!qtMap.has(c.fyKey)) {
                    qtMap.set(c.fyKey, { key:c.fyKey, quarter:c.quarter, cells:[] });
                    qtGroups.push(qtMap.get(c.fyKey));
                }
                qtMap.get(c.fyKey).cells.push(c);
            });
            qtGroups.forEach((qg, qi) => {
                const x1 = dToX(qg.cells[0].start);
                const x2 = dToX(qg.cells[qg.cells.length - 1].end);
                const qw  = Math.max(0, x2 - x1);
                if (qi > 0) {
                    slide.addShape(pptx.shapes.LINE, {
                        x:x1, y:qtRowY + 0.03, w:0, h:ROW_QT_H - 0.06,
                        line:{color:C.pillDivider, width:0.6, transparency:50}
                    });
                }
                slide.addText(`Q${qg.quarter}`, {
                    x:x1, y:qtRowY, w:qw, h:ROW_QT_H,
                    fontSize:10, bold:true, color:C.qtText,
                    fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
                });
            });

            // ── Row 3: Month + Year ─────────────────────────────────────
            const moRowY = PILL_TOP + ROW_FY_H + ROW_QT_H;
            slide.addShape(pptx.shapes.RECTANGLE, {
                x:TX, y:moRowY, w:TW, h:ROW_MO_H,
                fill:{color:C.pillMO}, line:{color:C.pillMO}
            });
            axis.cells.forEach((c, i) => {
                const cx  = dToX(c.start);
                const cx2 = dToX(c.end);
                const cw  = Math.max(0, cx2 - cx);
                if (i > 0) {
                    slide.addShape(pptx.shapes.LINE, {
                        x:cx, y:moRowY + 0.03, w:0, h:ROW_MO_H - 0.06,
                        line:{color:C.pillDivider, width:0.5, transparency:55}
                    });
                }
                slide.addText(c.label, {
                    x:cx, y:moRowY, w:cw, h:ROW_MO_H,
                    fontSize:9, bold:true, color:'FFFFFF',
                    fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
                });
            });

        } else {
            // ── quarter unit: 2-row pill (FY + Quarter) ─────────────────
            const fyRowY = PILL_TOP;
            const qtRowY = PILL_TOP + ROW_FY_H;
            // Rounded rect for full pill
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x:TX, y:fyRowY, w:TW, h:PILL_AREA_H, rectRadius:0.05,
                fill:{color:C.pillFY}, line:{color:C.pillFY}
            });
            // FY groups
            const fyGroupsQ = [];
            const fyMapQ = new Map();
            axis.cells.forEach(c => {
                if (!fyMapQ.has(c.fyLabel)) {
                    fyMapQ.set(c.fyLabel, { fyLabel:c.fyLabel, cells:[] });
                    fyGroupsQ.push(fyMapQ.get(c.fyLabel));
                }
                fyMapQ.get(c.fyLabel).cells.push(c);
            });
            fyGroupsQ.forEach((fg, gi) => {
                const x1 = dToX(fg.cells[0].start);
                const x2 = dToX(fg.cells[fg.cells.length-1].end);
                const fw  = Math.max(0, x2 - x1);
                if (gi > 0) slide.addShape(pptx.shapes.LINE, {
                    x:x1, y:fyRowY + 0.03, w:0, h:ROW_FY_H - 0.06,
                    line:{color:C.pillDivider, width:0.6, transparency:50}
                });
                slide.addText(fg.fyLabel, {
                    x:x1, y:fyRowY, w:Math.max(0.10, fw), h:ROW_FY_H,
                    fontSize:9, bold:true, color:'FFFFFF',
                    fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
                });
            });
            // Quarter row
            slide.addShape(pptx.shapes.RECTANGLE, {
                x:TX, y:qtRowY, w:TW, h:ROW_QT_H,
                fill:{color:C.pillQT}, line:{color:C.pillQT}
            });
            axis.cells.forEach((c, i) => {
                const cx  = dToX(c.start);
                const cx2 = dToX(c.end);
                const cw  = Math.max(0, cx2 - cx);
                if (i > 0) slide.addShape(pptx.shapes.LINE, {
                    x:cx, y:qtRowY + 0.03, w:0, h:ROW_QT_H - 0.06,
                    line:{color:C.pillDivider, width:0.6, transparency:50}
                });
                slide.addText(c.label, {
                    x:cx, y:qtRowY, w:cw, h:ROW_QT_H,
                    fontSize:10, bold:true, color:C.qtText,
                    fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
                });
            });
        }

        // ── Vertical grid lines (align with axis cell boundaries) ────────
        axis.cells.forEach((c, i) => {
            if (i === 0) return;
            const gx = dToX(c.start);
            slide.addShape(pptx.shapes.LINE, {
                x:gx, y:CONTENT_TOP - 0.03, w:0, h:PILL_TOP - (CONTENT_TOP - 0.03),
                line:{color:C.guide, width:0.6}
            });
        });

        // (Orange year label removed)

        // ── Milestones ────────────────────────────────────────────────────
        const msTop = PILL_TOP + PILL_AREA_H + 0.10;
        const placed = [];
        (milestones || []).forEach(m => {
            const d = ymd(m.date);
            if (!d || d < axis.axisStart || d > axis.axisEnd) return;
            const mx = dToX(d);
            let lane = 0;
            while (placed.some(p => p.lane === lane && Math.abs(p.x - mx) < 1.20)) lane++;
            placed.push({ x:mx, lane });
            const baseY = msTop + lane * 0.42;
            slide.addShape(pptx.shapes.LINE, {
                x:mx, y:PILL_TOP + PILL_AREA_H, w:0, h:baseY - (PILL_TOP + PILL_AREA_H),
                line:{color:C.msLine, width:0.75, dashType:'dash'}
            });
            const sz = 0.14, col = (m.color || '#3B8FD9').replace('#','');
            const sh  = m.shape === 'star' ? pptx.shapes.STAR_5 : pptx.shapes.DIAMOND;
            slide.addShape(sh, { x:mx-sz/2, y:baseY-sz/2, w:sz, h:sz, fill:{color:col}, line:{color:col} });
            slide.addText(fmtDOWMD(d), {
                x:mx-0.65, y:baseY+sz/2+0.02, w:1.30, h:0.16,
                fontSize:7, color:C.dateText, fontFace:'Segoe UI', align:'center', valign:'middle'
            });
            slide.addText(m.name || '', {
                x:mx-0.65, y:baseY+sz/2+0.16, w:1.30, h:0.18,
                fontSize:9, bold:true, color:C.ink, fontFace:'Segoe UI',
                align:'center', valign:'middle'
            });
        });

        // ── Today marker ──────────────────────────────────────────────────
        const todayD = new Date(); todayD.setHours(0,0,0,0);
        if (todayD >= axis.axisStart && todayD <= axis.axisEnd) {
            const todayX = dToX(todayD);
            slide.addShape(pptx.shapes.LINE, {
                x:todayX, y:CONTENT_TOP - 0.04, w:0, h:PILL_TOP - CONTENT_TOP + 0.04,
                line:{color:'EF4444', width:1.5}
            });
            slide.addText('Today', {
                x:todayX - 0.24, y:CONTENT_TOP - 0.22, w:0.48, h:0.18,
                fontSize:7, bold:true, color:'EF4444',
                fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
            });
        }

        // ── Status legend (centred, matches web-app STATUS_COLORS) ───────
        const LEGEND = [
            { label:'Not Started',  color:'808080' },
            { label:'In Progress',  color:'4C9141' },
            { label:'Delayed',      color:'FFA500' },
            { label:'Blocked',      color:'CC0000' },
            { label:'Action Needed',color:'4A0072' },
            { label:'Completed',    color:'00B4D8' },
        ];
        const dotSz   = 0.10;
        const dotGap  = 0.05;
        const itemGap = 0.18;
        // Pre-measure each label at 7.5pt to compute total legend width
        const labelWidths = LEGEND.map(s => Math.max(0.55, estW(s.label, 7.5) + 0.06));
        const itemWidths  = labelWidths.map(lw => dotSz + dotGap + lw);
        const totalW = itemWidths.reduce((a, b) => a + b, 0) + itemGap * (LEGEND.length - 1);
        let lx = (W - totalW) / 2;
        const legendH = 0.18;
        const legendY = H - 0.34;

        LEGEND.forEach((s, i) => {
            slide.addShape(pptx.shapes.OVAL, {
                x:lx, y:legendY + (legendH - dotSz) / 2, w:dotSz, h:dotSz,
                fill:{color:s.color}, line:{color:s.color}
            });
            slide.addText(s.label, {
                x:lx + dotSz + dotGap, y:legendY, w:labelWidths[i], h:legendH,
                fontSize:7.5, color:C.ink, fontFace:'Segoe UI', valign:'middle', wrap:false
            });
            lx += itemWidths[i] + (i < LEGEND.length - 1 ? itemGap : 0);
        });

        // ── Footer ────────────────────────────────────────────────────────
        const ts = new Date().toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'});
        slide.addText(`Generated ${ts}`, {
            x:0, y:H - 0.16, w:W, h:0.14,
            fontSize:6.5, color:C.dateText, fontFace:'Segoe UI', align:'center', valign:'middle'
        });
    }

    // ══════════════════════════════════════════════════════════════════════
    // GROUP HEADER — compact, no timeline bar
    // ══════════════════════════════════════════════════════════════════════
    function drawGroupHeader(slide, row, y, h) {
        const label  = (row.group || '') + (row.isContinuation ? ' (cont.)' : (row.count > 0 ? ` (${row.count})` : ''));
        const gColor = (groups[row.group]?.color || '#3D4A5C').replace('#', '');

        // 1 — Full-width lavender background (always rendered first so it sits behind text)
        slide.addShape(pptx.shapes.RECTANGLE, {
            x:GX, y:y + 0.02, w:W - GX - MARGIN_R, h:h - 0.04,
            fill:{color:C.groupHdr}, line:{color:C.groupBorder, width:0.75}
        });

        // 2 — Accent bar — matches the group's assigned colour
        slide.addShape(pptx.shapes.RECTANGLE, {
            x:GX, y:y + 0.02, w:ACCENT_W, h:h - 0.04,
            fill:{color:gColor}, line:{color:gColor}
        });

        // 3 — Label text: spans the full lavender bar width (GRP_TEXT_X → right margin).
        slide.addText(label, {
            x:GRP_TEXT_X, y:y + 0.01, w:W - GRP_TEXT_X - MARGIN_R, h:h - 0.02,
            fontSize:GRP_FS, bold:true, color:C.groupText,
            fontFace:'Segoe UI', align:'left', valign:'middle',
            wrap:false, autoFitType:'shrink'
        });

    }

    // ══════════════════════════════════════════════════════════════════════
    // TASK ROW — name to the left of chevron, date to the right
    // ══════════════════════════════════════════════════════════════════════
    function drawTaskRow(slide, row, y, h) {
        const t      = row.task;
        const tColor = (t.color || '#808080').replace('#', '');
        const ts = ymd(t.startDate), te = ymd(t.endDate);
        const x1 = dToX(ts), x2 = dToX(te);
        if (x1 === null || x2 === null) return;
        const bw     = Math.max(0, x2 - x1);
        const isZero = bw < 0.04;
        // Content zone = row height minus the inter-task gap at the bottom
        const ch     = Math.max(0.20, h - TASK_GAP);
        const barH   = Math.min(0.26, ch * 0.66);
        const barY   = y + (ch - barH) / 2;
        const display = taskDisplayName(t.name);

        slide.addText(display, {
            x:GRP_TEXT_X, y, w:TX - GRP_TEXT_X, h:ch,
            fontSize:TASK_FS, bold:true, color:C.ink,
            fontFace:'Segoe UI', align:'left', valign:'middle',
            wrap:true
        });

        // ── Chevron / diamond with progress fill ─────────────────────────
        if (isZero) {
            const sz = Math.min(0.16, barH);
            slide.addShape(pptx.shapes.DIAMOND, {
                x:x1 - sz/2, y:barY + (barH-sz)/2, w:sz, h:sz,
                fill:{color:tColor}, line:{color:tColor}
            });
        } else {
            const pct = typeof t.progress === 'number' ? t.progress : null;
            drawChevronWithProgress(slide, x1, barY, bw, barH, tColor, pct);
        }

        // ── Date range label — right of bar ─────────────────────────────
        const dateX = x2 + (isZero ? 0.14 : 0.12);
        const dateW = (W - MARGIN_R) - dateX;
        if (dateW > 0.25) {
            slide.addText(fmtRange(ts, te), {
                x:dateX, y:barY, w:dateW, h:barH,
                fontSize:Math.min(9, barH * 30), color:C.dateText,
                fontFace:'Segoe UI', align:'left', valign:'middle', wrap:false
            });
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    // GROUP-ONLY ROW
    // Layout:  [accent bar] [group name — wraps in name column] [chevron in timeline] [date]
    // The row height is pre-computed by groupOnlyRowH() so the chevron always
    // stays vertically centred relative to however many lines the name occupies.
    // ══════════════════════════════════════════════════════════════════════
    function drawGroupOnlyRow(slide, row, y, h) {
        const gColor = (groups[row.group]?.color || '#808080').replace('#', '');
        const x1 = dToX(row.gStart), x2 = dToX(row.gEnd);
        if (x1 === null || x2 === null) return;
        const bw  = Math.max(0, x2 - x1);
        const pct = groupProgress(row.group);

        // Layout: chevron + optional progress track below it
        const barH     = 0.26;
        const trackH   = 0.06;
        const trackGap = 0.04;
        const hasTrack = pct !== null && bw >= 0.04;
        const zoneH    = hasTrack ? barH + trackGap + trackH : barH;
        const barY     = y + (h - zoneH) / 2;
        const trackY   = barY + barH + trackGap;
        const minW     = Math.max(0.28, bw);

        const label = (row.group || '') + (row.count > 0 ? ` (${row.count})` : '');

        // ── Layer 1: Background — name column ────────────────────────────
        slide.addShape(pptx.shapes.RECTANGLE, {
            x:GX, y:y + 0.02, w:TX - GX, h:h - 0.04,
            fill:{color:C.groupHdr}, line:{color:C.groupBorder, width:0.75}
        });

        // ── Layer 2: Accent bar ───────────────────────────────────────────
        slide.addShape(pptx.shapes.RECTANGLE, {
            x:GX, y:y + 0.02, w:ACCENT_W, h:h - 0.04,
            fill:{color:gColor}, line:{color:gColor}
        });

        // ── Layer 3: Group name ───────────────────────────────────────────
        slide.addText(label, {
            x:GRP_TEXT_X, y, w:TX - GRP_TEXT_X, h,
            fontSize:GRP_FS, bold:true, color:C.groupText,
            fontFace:'Segoe UI', align:'left', valign:'middle',
            wrap:true
        });

        // ── Layer 4: Chevron / diamond ────────────────────────────────────
        if (bw < 0.04) {
            const sz = 0.16;
            slide.addShape(pptx.shapes.DIAMOND, {
                x:x1 - sz/2, y:barY + (barH - sz)/2, w:sz, h:sz,
                fill:{color:gColor}, line:{color:gColor}
            });
        } else {
            drawChevronWithProgress(slide, x1, barY, bw, barH, gColor, pct);

            // ── Layer 5a: % label overlaid on chevron (white, centred) ────
            if (pct !== null) {
                slide.addText(`${pct}%`, {
                    x:x1, y:barY, w:minW, h:barH,
                    fontSize:8, bold:true, color:'FFFFFF',
                    fontFace:'Segoe UI', align:'center', valign:'middle',
                    wrap:false
                });
            }

            // ── Layer 5b: Thin progress track below chevron ───────────────
            if (hasTrack && trackY + trackH <= y + h - 0.01) {
                // Background pill
                slide.addShape(pptx.shapes.RECTANGLE, {
                    x:x1, y:trackY, w:minW, h:trackH,
                    fill:{color:'E8E8E8'}, line:{color:'CCCCCC', width:0.5}
                });
                // Filled portion
                if (pct > 0) {
                    const fillW = Math.max(0.04, minW * pct / 100);
                    slide.addShape(pptx.shapes.RECTANGLE, {
                        x:x1, y:trackY, w:fillW, h:trackH,
                        fill:{color:gColor}, line:{color:gColor}
                    });
                }
            }

            // ── Layer 6: Date range right of bar ──────────────────────────
            const dateX = x1 + minW + 0.10;
            const dateW = (W - MARGIN_R) - dateX;
            if (dateW > 0.25) {
                slide.addText(fmtRange(row.gStart, row.gEnd), {
                    x:dateX, y:barY, w:dateW, h:barH,
                    fontSize:9, color:C.dateText,
                    fontFace:'Segoe UI', align:'left', valign:'middle', wrap:false
                });
            }
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    // PAGINATION + RENDER
    // ══════════════════════════════════════════════════════════════════════
    const TASK_GAP    = 0.14;   // whitespace below each task row
    const TASK_MAX_LINES = 2;   // never let a task name exceed 2 lines in the PPT

    // Normalize a task name: collapse embedded newlines/tabs to spaces.
    // Full name is preserved — no truncation.
    function taskDisplayName(name) {
        return (name || '').replace(/[\r\n\t]+/g, ' ').replace(/\s{2,}/g, ' ').trim().toUpperCase();
    }

    function taskRowH(t) {
        const colW   = TX - GRP_TEXT_X;
        const display = taskDisplayName(t.name);
        const lines  = Math.max(1, Math.ceil(estW(display, TASK_FS) / colW));
        return lines * 0.20 + 0.20 + TASK_GAP;   // content + gap
    }

    function groupOnlyRowH(g, count) {
        const label  = (g || '') + (count > 0 ? ` (${count})` : '');
        const colW   = TX - GRP_TEXT_X;
        const lines  = Math.max(1, Math.ceil(estW(label, GRP_FS) / colW));
        const lineH  = 0.19;
        const padV   = 0.18;
        const textH  = lines * lineH + padV;
        // Must fit: chevron (0.26) + gap (0.04) + progress track (0.06) + padding (0.18)
        const minH   = 0.54;
        return Math.min(Math.max(textH, minH), 1.40);
    }

    function paginate(rows) {
        // Only scale UP (fill empty space). Never shrink — font size doesn't scale
        // with rowH, so shrinking causes text to overflow its allocated box.
        const scaleToFit = (pg) => {
            const pgH = pg.reduce((a, r) => a + r.rowH, 0);
            const sc  = Math.min(1.30, CONTENT_H / Math.max(pgH, 0.01));
            if (sc <= 1.0) return pg;
            return pg.map(r => ({ ...r, rowH: r.rowH * sc }));
        };

        const totalH = rows.reduce((a, r) => a + r.rowH, 0);
        if (totalH <= CONTENT_H) return [scaleToFit(rows)];

        // Step 1: raw split into pages
        const rawPages = []; let page = [], used = 0;
        for (const r of rows) {
            if (used + r.rowH > CONTENT_H && page.length) {
                rawPages.push(page); page = []; used = 0;
            }
            page.push(r); used += r.rowH;
        }
        if (page.length) rawPages.push(page);

        // Step 2: inject cont. headers — make room first so the page stays within CONTENT_H.
        // If adding GRP_H would push the page over the limit, pop rows off the end
        // and prepend them to the next page before injecting the header.
        const groupRowFor = {};
        rows.forEach(r => { if (r.type === 'group') groupRowFor[r.group] = r; });

        const processed = [];
        for (let pi = 0; pi < rawPages.length; pi++) {
            const pg = rawPages[pi];
            if (pi === 0) { processed.push(pg); continue; }
            const first = pg[0];
            if (!first || first.type !== 'task') { processed.push(pg); continue; }
            const src = groupRowFor[first.group];
            if (!src) { processed.push(pg); continue; }

            const contRow = { ...src, isContinuation: true, rowH: GRP_H };
            // Pop rows off this page until cont. header fits
            const working = [...pg];
            let pgH = working.reduce((a, r) => a + r.rowH, 0) + contRow.rowH;
            const spill = [];
            while (pgH > CONTENT_H && working.length > 1) {
                spill.unshift(working.pop());
                pgH -= spill[0].rowH;
            }
            // Prepend spilled rows to the next raw page (or create a new one)
            if (spill.length) {
                if (pi + 1 < rawPages.length) {
                    rawPages[pi + 1] = [...spill, ...rawPages[pi + 1]];
                } else {
                    rawPages.push(spill);
                }
            }
            processed.push([contRow, ...working]);
        }

        // Step 3: scale each page upward to fill CONTENT_H (never down)
        return processed.map(scaleToFit);
    }

    const safeName = (projects[activeProjectIndex]?.name || 'Gantt').replace(/[^\w\-]+/g, '_');
    const dateStamp = new Date().toISOString().slice(0, 10);

    // ── groups-only ───────────────────────────────────────────────────────
    if (mode === 'groups-only') {
        const groupRows = ordered.map(g => {
            const gTasks = tasks.filter(t => t.group === g);
            const dates  = gTasks.flatMap(t => [ymd(t.startDate), ymd(t.endDate)]).filter(Boolean);
            const gStart = dates.length ? new Date(Math.min(...dates)) : null;
            const gEnd   = dates.length ? new Date(Math.max(...dates)) : null;
            return { group:g, count:gTasks.length, gStart, gEnd,
                     rowH:groupOnlyRowH(g, gTasks.length) };
        }).filter(r => r.gStart && r.gEnd);

        if (!groupRows.length) { showNotification('No group data', 'error'); return; }

        paginate(groupRows).forEach((rows, pi, pages) => {
            const slide = pptx.addSlide();
            chrome(slide, pages.length > 1 ? `Page ${pi+1} / ${pages.length}` : '');
            let y = CONTENT_TOP;
            rows.forEach(r => { drawGroupOnlyRow(slide, r, y, r.rowH); y += r.rowH; });
        });

        pptx.writeFile({ fileName:`${safeName}_groups-only_${dateStamp}.pptx` })
            .then(() => showNotification('PPTX exported', 'success'))
            .catch(err => { console.error(err); showNotification('PPTX export failed', 'error'); });
        return;
    }

    // ── mixed / groups-and-tasks ──────────────────────────────────────────
    // For 'mixed': groups in expandedGroups show header + tasks; others show summary bar.
    // For 'groups-and-tasks': all groups are expanded (expandedGroups = all).
    const effectiveExpanded = mode === 'mixed'
        ? (expandedGroups instanceof Set ? expandedGroups : new Set())
        : new Set(ordered);   // all expanded for groups-and-tasks

    const allRows = [];
    ordered.forEach(g => {
        const gTasks = tasks.filter(t => t.group === g);
        if (!gTasks.length) return;
        const dates  = gTasks.flatMap(t => [ymd(t.startDate), ymd(t.endDate)]).filter(Boolean);
        const gStart = dates.length ? new Date(Math.min(...dates)) : null;
        const gEnd   = dates.length ? new Date(Math.max(...dates)) : null;
        if (!gStart || !gEnd) return;

        if (effectiveExpanded.has(g)) {
            allRows.push({ type:'group', group:g, count:gTasks.length, gStart, gEnd, rowH:GRP_H });
            gTasks.forEach(t => allRows.push({ type:'task', task:t, group:g, rowH:taskRowH(t) }));
        } else {
            allRows.push({ type:'group-summary', group:g, count:gTasks.length, gStart, gEnd,
                           rowH:groupOnlyRowH(g, gTasks.length) });
        }
    });

    paginate(allRows).forEach((rows, pi, pages) => {
        const slide = pptx.addSlide();
        chrome(slide, pages.length > 1 ? `Page ${pi+1} / ${pages.length}` : '');
        let y = CONTENT_TOP;
        rows.forEach(r => {
            if      (r.type === 'group')         drawGroupHeader(slide, r, y, r.rowH);
            else if (r.type === 'group-summary')  drawGroupOnlyRow(slide, r, y, r.rowH);
            else                                  drawTaskRow(slide, r, y, r.rowH);
            y += r.rowH;
        });
    });

    const fileSuffix = mode === 'mixed' ? 'mixed' : 'groups-and-tasks';
    pptx.writeFile({ fileName:`${safeName}_${fileSuffix}_${dateStamp}.pptx` })
        .then(() => showNotification('PPTX exported', 'success'))
        .catch(err => { console.error(err); showNotification('PPTX export failed', 'error'); });
}
