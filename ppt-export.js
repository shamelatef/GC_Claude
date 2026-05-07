/**
 * Export Gantt Chart to PowerPoint — Office Timeline style.
 *
 * mode === 'tasks'                → task rows only (tiny group tag on far left)
 * mode === 'tasks-with-subtasks'  → task rows + subtask arrows
 * mode === 'groups-and-tasks'     → group-name header row, then task rows under it
 * mode === 'groups-only'          → one summary bar per group (no individual tasks)
 *
 * Requires PptxGenJS and host-page globals:
 *   tasks, milestones, groups, groupOrder,
 *   projects, activeProjectIndex, parseYMD, showNotification
 */
function exportToPPT(mode) {
    const VALID = ['tasks','tasks-with-subtasks','groups-and-tasks','groups-only'];
    if (!VALID.includes(mode)) mode = 'tasks';

    if (!tasks || tasks.length === 0) {
        showNotification('No tasks to export', 'error');
        return;
    }
    showNotification('Building PPTX…', 'info');

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33" × 7.5"

    // ── Slide constants (inches) ─────────────────────────────────────────
    const W = 13.33, H = 7.5;
    const MARGIN_L = 0.55, MARGIN_R = 0.3, MARGIN_T = 0.45, MARGIN_B = 0.35;
    const SIDE_COL_W = 1.30;

    // "tasks" / "tasks-with-subtasks" use a narrow group-tag column + label column
    // "groups-and-tasks" / "groups-only" use a single wider label column (no group tag)
    const useGroupHeaders = (mode === 'groups-and-tasks' || mode === 'groups-only');
    const GROUP_TAG_W = useGroupHeaders ? 0 : 0.55;   // slim accent stripe for task modes
    const LABEL_COL_W = useGroupHeaders ? 2.70 : 2.50; // wide enough for long task names

    const GX  = MARGIN_L;
    const LX  = GX + GROUP_TAG_W;
    const TX  = LX + LABEL_COL_W;
    const TR  = W - MARGIN_R - SIDE_COL_W;
    const TW  = TR - TX;

    const PILL_H    = 0.32;
    const MS_AREA   = 1.05;
    const PILL_TOP  = H - MARGIN_B - MS_AREA - PILL_H;
    const CONTENT_TOP = MARGIN_T + 0.05;
    const CONTENT_BOT = PILL_TOP - 0.18;
    const CONTENT_H   = CONTENT_BOT - CONTENT_TOP;

    const C = {
        bg:'FFFFFF', pill:'475569', guide:'D8DDE3', ink:'1A1A1A',
        muted:'9AA0A6', msLine:'CFD4DA', year:'E58A3A', sideText:'374151',
        groupTag:'6B7280', groupHdr:'F0F4FF', groupHdrText:'1E3A5F'
    };

    // ── Label fit helper ─────────────────────────────────────────────────
    // Pre-computes font size so text never overflows its box.
    // Segoe UI Bold uppercase: ~0.0072" per character per point.
    const CHAR_W_PER_PT = 0.0072;
    function fitLabel(text, availW) {
        for (let fs = 11; fs >= 6; fs--) {
            if (text.length * fs * CHAR_W_PER_PT <= availW) return { text, fontSize: fs };
        }
        // Still too long at 6pt — truncate with ellipsis
        const maxChars = Math.floor(availW / (6 * CHAR_W_PER_PT));
        return { text: text.slice(0, Math.max(3, maxChars - 1)) + '…', fontSize: 6 };
    }

    // ── Date helpers ─────────────────────────────────────────────────────
    function ymd(s){ return (s instanceof Date) ? s : parseYMD(s); }
    function days(a,b){ return Math.round((b-a)/86400000); }
    function startOfWeekMon(d){
        const x=new Date(d.getFullYear(),d.getMonth(),d.getDate());
        const wd=(x.getDay()+6)%7; x.setDate(x.getDate()-wd); return x;
    }
    const MN=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const DOW=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    function fmtMD(d){ return `${MN[d.getMonth()]} ${d.getDate()}`; }
    function fmtDOWMD(d){ return `${DOW[d.getDay()]} ${MN[d.getMonth()]} ${d.getDate()}`; }

    // ── Global date range ────────────────────────────────────────────────
    const allDates=[
        ...tasks.flatMap(t=>[ymd(t.startDate),ymd(t.endDate)]),
        ...(milestones||[]).map(m=>ymd(m.date))
    ].filter(Boolean);
    if(!allDates.length){ showNotification('No date data to export','error'); return; }
    const minDate=new Date(Math.min(...allDates));
    const maxDate=new Date(Math.max(...allDates));

    // ── Axis ─────────────────────────────────────────────────────────────
    const totalDays=Math.max(1,days(minDate,maxDate));
    const unit = totalDays<=80 ? 'week' : totalDays<=400 ? 'month' : 'quarter';

    function buildAxis(){
        const cells=[];
        if(unit==='week'){
            const start=startOfWeekMon(minDate);
            let cur=new Date(start), i=1;
            while(cur<=maxDate){
                const next=new Date(cur); next.setDate(next.getDate()+7);
                cells.push({label:i===1?'Week 1':String(i), start:new Date(cur), end:new Date(next)});
                cur=next; i++;
            }
            return {axisStart:start, axisEnd:cells[cells.length-1].end, cells};
        }
        if(unit==='month'){
            const start=new Date(minDate.getFullYear(),minDate.getMonth(),1);
            let cur=new Date(start);
            while(cur<=maxDate){
                const next=new Date(cur.getFullYear(), cur.getMonth()+1, 1);
                cells.push({label:MN[cur.getMonth()], start:new Date(cur), end:new Date(next)});
                cur=next;
            }
            return {axisStart:start, axisEnd:cells[cells.length-1].end, cells};
        }
        const qStartMo=Math.floor(minDate.getMonth()/3)*3;
        const start=new Date(minDate.getFullYear(),qStartMo,1);
        let cur=new Date(start);
        while(cur<=maxDate){
            const next=new Date(cur.getFullYear(), cur.getMonth()+3, 1);
            const q=Math.floor(cur.getMonth()/3)+1;
            cells.push({label:`Q${q} ${cur.getFullYear()}`, start:new Date(cur), end:new Date(next)});
            cur=next;
        }
        return {axisStart:start, axisEnd:cells[cells.length-1].end, cells};
    }
    const axis=buildAxis();
    const axisMs=axis.axisEnd-axis.axisStart;
    function dToX(raw){
        const d=ymd(raw); if(!d) return null;
        const t=Math.max(axis.axisStart.getTime(), Math.min(axis.axisEnd.getTime(), d.getTime()));
        return TX + ((t-axis.axisStart.getTime())/axisMs)*TW;
    }

    // ── Ordered groups ───────────────────────────────────────────────────
    const ordered = (groupOrder.length ? groupOrder : Object.keys(groups))
        .filter(g => tasks.some(t => t.group === g));

    // ════════════════════════════════════════════════════════════════════
    // SHARED CHROME (title, grid lines, pill, milestones, footer)
    // ════════════════════════════════════════════════════════════════════
    function chrome(slide, pageLabel){
        slide.addShape(pptx.shapes.RECTANGLE,{
            x:0, y:0, w:W, h:H, fill:{color:C.bg}, line:{color:C.bg}
        });
        const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';
        slide.addText(projName,{
            x:MARGIN_L, y:0.05, w:7, h:0.30, fontSize:13, bold:true,
            color:C.ink, fontFace:'Segoe UI', valign:'middle'
        });
        if(pageLabel){
            slide.addText(pageLabel,{
                x:W-1.6, y:0.05, w:1.4, h:0.30, fontSize:10, color:C.muted,
                fontFace:'Segoe UI', valign:'middle', align:'right'
            });
        }
        axis.cells.forEach((c,i)=>{
            if(i===0) return;
            const gx=dToX(c.start);
            slide.addShape(pptx.shapes.LINE,{
                x:gx, y:CONTENT_TOP-0.05, w:0, h:PILL_TOP-(CONTENT_TOP-0.05),
                line:{color:C.guide, width:0.75}
            });
        });
        // Timeline pill
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE,{
            x:TX, y:PILL_TOP, w:TW, h:PILL_H, rectRadius:0.06,
            fill:{color:C.pill}, line:{color:C.pill}
        });
        const cellW=TW/axis.cells.length;
        axis.cells.forEach((c,i)=>{
            const cx=TX+i*cellW;
            if(i>0){
                slide.addShape(pptx.shapes.LINE,{
                    x:cx, y:PILL_TOP+0.04, w:0, h:PILL_H-0.08,
                    line:{color:'FFFFFF', width:0.6, transparency:60}
                });
            }
            slide.addText(c.label,{
                x:cx, y:PILL_TOP, w:cellW, h:PILL_H,
                fontSize:i===0?11:12, bold:true, color:'FFFFFF',
                fontFace:'Segoe UI', align:'center', valign:'middle'
            });
        });
        slide.addText(String(axis.axisStart.getFullYear()),{
            x:MARGIN_L-0.15, y:PILL_TOP-0.02, w:GROUP_TAG_W+LABEL_COL_W, h:PILL_H+0.04,
            fontSize:22, bold:true, color:C.year, fontFace:'Segoe UI',
            align:'left', valign:'middle'
        });
        // Milestones
        const msTop=PILL_TOP+PILL_H+0.14;
        const placed=[];
        const minSpacing=1.25;
        (milestones||[]).forEach(m=>{
            const d=ymd(m.date);
            if(!d || d<axis.axisStart || d>axis.axisEnd) return;
            const mx=dToX(d);
            let lane=0;
            while(placed.some(p=>p.lane===lane && Math.abs(p.x-mx)<minSpacing)) lane++;
            placed.push({x:mx, lane});
            const baseY=msTop+lane*0.46;
            slide.addShape(pptx.shapes.LINE,{
                x:mx, y:PILL_TOP+PILL_H, w:0, h:baseY-(PILL_TOP+PILL_H),
                line:{color:C.msLine, width:0.75, dashType:'dash'}
            });
            const sz=0.16;
            const color=(m.color||'#3B8FD9').replace('#','');
            const shape=m.shape==='star'?pptx.shapes.STAR_5:pptx.shapes.DIAMOND;
            slide.addShape(shape,{x:mx-sz/2, y:baseY-sz/2, w:sz, h:sz, fill:{color}, line:{color}});
            slide.addText(fmtDOWMD(d),{
                x:mx-0.7, y:baseY+sz/2+0.02, w:1.4, h:0.18,
                fontSize:8, color:C.muted, fontFace:'Segoe UI', align:'center', valign:'middle'
            });
            slide.addText(m.name||'',{
                x:mx-0.7, y:baseY+sz/2+0.18, w:1.4, h:0.20,
                fontSize:9.5, bold:true, color:C.ink, fontFace:'Segoe UI',
                align:'center', valign:'middle'
            });
        });
        const ts=new Date().toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
        slide.addText(`Generated ${ts}`,{
            x:0, y:H-0.22, w:W, h:0.2, fontSize:7, color:C.muted,
            fontFace:'Segoe UI', align:'center', valign:'middle'
        });
    }

    // ════════════════════════════════════════════════════════════════════
    // MODE: tasks / tasks-with-subtasks
    // ════════════════════════════════════════════════════════════════════
    if (mode === 'tasks' || mode === 'tasks-with-subtasks') {
        const includeSubtasks = (mode === 'tasks-with-subtasks');
        const rowsList = [];
        ordered.forEach(g=>{
            tasks.filter(t=>t.group===g).forEach(t=>{
                const subs = (includeSubtasks && Array.isArray(t.subtasks)) ? t.subtasks : [];
                rowsList.push({task:t, group:g, subs});
            });
        });

        function rowWeight(r){ return 1 + (r.subs.length ? 0.55*r.subs.length + 0.18 : 0); }
        const totalWeight = rowsList.reduce((a,r)=>a+rowWeight(r),0);
        const ROW_UNIT_MIN=0.34, ROW_UNIT_MAX=0.55;
        let rowUnit=Math.min(ROW_UNIT_MAX, CONTENT_H/Math.max(1,totalWeight));
        let pages;
        if(rowUnit>=ROW_UNIT_MIN){
            pages=[rowsList];
        } else {
            rowUnit=ROW_UNIT_MIN;
            pages=[]; let page=[],used=0;
            for(const r of rowsList){
                const h=rowWeight(r)*rowUnit;
                if(used+h>CONTENT_H && page.length>0){ pages.push(page); page=[]; used=0; }
                page.push(r); used+=h;
            }
            if(page.length>0) pages.push(page);
        }

        function drawTaskRow(slide, row, y, h){
            const t=row.task;
            const tColor=(t.color||'#475569').replace('#','');
            const groupColor=(groups[row.group]?.color||'#9AA0A6').replace('#','');
            const ts=ymd(t.startDate), te=ymd(t.endDate);
            const x1=dToX(ts), x2=dToX(te);
            if(x1===null||x2===null) return;
            const w=Math.max(0,x2-x1);
            const isZero=w<0.04;
            // Group tag — colored accent stripe only
            slide.addShape(pptx.shapes.RECTANGLE,{
                x:GX, y:y+h*0.15, w:0.10, h:h*0.70,
                fill:{color:groupColor}, line:{color:groupColor}
            });
            // Task name — pre-fitted so it never clips or overflows into the timeline
            const labelMaxW = TX - LX - 0.20;
            const lbl = fitLabel((t.name||'').toUpperCase(), labelMaxW);
            slide.addText(lbl.text,{
                x:LX+0.08, y:y, w:labelMaxW, h:h,
                fontSize:lbl.fontSize, bold:true, color:C.ink,
                fontFace:'Segoe UI', align:'left', valign:'middle',
                wrap:false
            });
            const subs=row.subs, hasSubs=subs.length>0;
            const parentH=hasSubs?Math.min(0.34,h*0.42):Math.min(0.34,h*0.78);
            const parentY=y+(hasSubs?0.04:(h-parentH)/2);
            if(isZero){
                const sz=Math.min(0.18,parentH);
                slide.addShape(pptx.shapes.DIAMOND,{
                    x:x1-sz/2, y:parentY+(parentH-sz)/2, w:sz, h:sz,
                    fill:{color:tColor}, line:{color:tColor}
                });
            } else {
                // Bar only — name is already in label column, no duplicate inside bar
                slide.addShape(pptx.shapes.PENTAGON,{
                    x:x1, y:parentY, w:Math.max(0.32,w), h:parentH,
                    fill:{color:tColor}, line:{color:tColor}
                });
            }
            slide.addText(`${fmtMD(ts)} – ${fmtMD(te)}`,{
                x:x2+0.08, y:parentY, w:(W-MARGIN_R)-(x2+0.08), h:parentH,
                fontSize:Math.min(10,parentH*26), bold:true, color:tColor,
                fontFace:'Segoe UI', align:'left', valign:'middle', wrap:false
            });
            if(hasSubs&&!isZero){
                const subsTop=parentY+parentH+0.04;
                const subsAvailH=(y+h)-subsTop-0.02;
                const subH=Math.max(0.14,Math.min(0.22,subsAvailH/Math.max(1,subs.length)));
                const tsMs=ts.getTime(),teMs=te.getTime();
                const step=(teMs-tsMs)/subs.length;
                subs.forEach((sub,si)=>{
                    const sStart=new Date(tsMs+si*step), sEnd=new Date(tsMs+(si+1)*step);
                    const sx1=dToX(sStart), sx2=dToX(sEnd);
                    const sw=Math.max(0,sx2-sx1);
                    const sy=subsTop+si*subH+(subH-Math.min(subH*0.78,0.18))/2;
                    const sah=Math.min(subH*0.78,0.18);
                    const subColor=tColor;
                    const shape=si===0?pptx.shapes.PENTAGON:pptx.shapes.CHEVRON;
                    slide.addShape(shape,{
                        x:sx1, y:sy, w:Math.max(0.22,sw), h:sah,
                        fill:{color:subColor, transparency:sub.done?0:35},
                        line:{color:subColor, transparency:sub.done?0:20}
                    });
                    const fs=Math.min(9,sah*32);
                    const charW=fs*0.0075;
                    const labelOpts={fontSize:fs,bold:true,fontFace:'Segoe UI',
                        align:'center',valign:'middle',wrap:false,
                        italic:!!sub.done,strike:!!sub.done};
                    if(sw>=(sub.name||'').length*charW+0.18&&sw>=0.5){
                        slide.addText(sub.name||'',{...labelOpts,
                            x:sx1+0.06,y:sy,w:Math.max(0.16,sw-0.18),h:sah,
                            color:'FFFFFF',align:'center'});
                    } else {
                        slide.addText(sub.name||'',{...labelOpts,
                            x:sx2+0.06,y:sy-0.02,w:(W-MARGIN_R)-(sx2+0.06),h:sah+0.04,
                            color:sub.done?C.muted:C.sideText,align:'left',bold:false});
                    }
                });
            }
        }

        pages.forEach((rows,pi)=>{
            const slide=pptx.addSlide();
            chrome(slide, pages.length>1?`Page ${pi+1} / ${pages.length}`:'');
            let y=CONTENT_TOP;
            rows.forEach(r=>{ const h=rowWeight(r)*rowUnit; drawTaskRow(slide,r,y,h); y+=h; });
        });

        const projName=projects[activeProjectIndex]?.name||'Gantt Chart';
        const safeName=projName.replace(/[^\w\-]+/g,'_');
        const suffix=includeSubtasks?'_with-subtasks':'_tasks';
        pptx.writeFile({fileName:`${safeName}${suffix}_${new Date().toISOString().slice(0,10)}.pptx`})
            .then(()=>showNotification('PPTX exported','success'))
            .catch(err=>{ console.error(err); showNotification('PPTX export failed','error'); });
        return;
    }

    // ════════════════════════════════════════════════════════════════════
    // MODE: groups-only
    //   One row per group — group name on left, summary bar on timeline
    // ════════════════════════════════════════════════════════════════════
    if (mode === 'groups-only') {
        // Build one entry per group with date range
        const groupRows = ordered.map(g => {
            const gTasks = tasks.filter(t => t.group === g);
            const dates = gTasks.flatMap(t=>[ymd(t.startDate),ymd(t.endDate)]).filter(Boolean);
            const gStart = dates.length ? new Date(Math.min(...dates)) : null;
            const gEnd   = dates.length ? new Date(Math.max(...dates)) : null;
            return { group:g, count:gTasks.length, gStart, gEnd };
        }).filter(r => r.gStart && r.gEnd);

        if (!groupRows.length) { showNotification('No group data to export','error'); return; }

        const ROW_UNIT_MIN=0.36, ROW_UNIT_MAX=0.65;
        const totalWeight = groupRows.length;
        let rowUnit = Math.min(ROW_UNIT_MAX, CONTENT_H / Math.max(1, totalWeight));
        let pages;
        if (rowUnit >= ROW_UNIT_MIN) {
            pages = [groupRows];
        } else {
            rowUnit = ROW_UNIT_MIN;
            pages = []; let page=[], used=0;
            for (const r of groupRows) {
                if (used + rowUnit > CONTENT_H && page.length > 0) { pages.push(page); page=[]; used=0; }
                page.push(r); used += rowUnit;
            }
            if (page.length > 0) pages.push(page);
        }

        function drawGroupRow(slide, row, y, h) {
            const gColor = (groups[row.group]?.color || '#475569').replace('#','');
            const x1 = dToX(row.gStart), x2 = dToX(row.gEnd);
            if (x1 === null || x2 === null) return;
            const w = Math.max(0, x2 - x1);
            const isZero = w < 0.04;
            const barH = Math.min(0.36, h * 0.72);
            const barY = y + (h - barH) / 2;

            // Left accent stripe
            slide.addShape(pptx.shapes.RECTANGLE,{
                x:GX, y:y+h*0.10, w:0.12, h:h*0.80,
                fill:{color:gColor}, line:{color:gColor}
            });
            // Group name
            const taskCount = row.count > 0 ? ` (${row.count})` : '';
            slide.addText((row.group || '') + taskCount, {
                x: GX + 0.20, y, w: LABEL_COL_W - 0.22, h,
                fontSize: Math.min(13, h * 28), bold: true, color: C.ink,
                fontFace: 'Segoe UI', align: 'left', valign: 'middle',
                wrap: false, shrinkText: true
            });

            if (isZero) {
                const sz = Math.min(0.20, barH);
                slide.addShape(pptx.shapes.DIAMOND,{
                    x:x1-sz/2, y:barY+(barH-sz)/2, w:sz, h:sz,
                    fill:{color:gColor}, line:{color:gColor}
                });
            } else {
                slide.addShape(pptx.shapes.PENTAGON,{
                    x:x1, y:barY, w:Math.max(0.35, w), h:barH,
                    fill:{color:gColor}, line:{color:gColor}
                });
                const fs = Math.min(12, barH * 30);
                const charW = fs * 0.0075;
                if (w >= (row.group||'').length * charW + 0.2 && w >= 0.5) {
                    slide.addText(row.group || '', {
                        x:x1+0.12, y:barY, w:Math.max(0.16, w-0.3), h:barH,
                        fontSize:fs, bold:true, color:'FFFFFF',
                        fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
                    });
                }
            }
            // Date range
            slide.addText(`${fmtMD(row.gStart)} – ${fmtMD(row.gEnd)}`,{
                x:x2+0.08, y:barY, w:(W-MARGIN_R)-(x2+0.08), h:barH,
                fontSize:Math.min(10,barH*26), bold:true, color:gColor,
                fontFace:'Segoe UI', align:'left', valign:'middle', wrap:false
            });
        }

        pages.forEach((rows, pi) => {
            const slide = pptx.addSlide();
            chrome(slide, pages.length>1?`Page ${pi+1} / ${pages.length}`:'');
            let y = CONTENT_TOP;
            rows.forEach(r => { drawGroupRow(slide, r, y, rowUnit); y += rowUnit; });
        });

        const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';
        const safeName = projName.replace(/[^\w\-]+/g,'_');
        pptx.writeFile({fileName:`${safeName}_groups-only_${new Date().toISOString().slice(0,10)}.pptx`})
            .then(()=>showNotification('PPTX exported','success'))
            .catch(err=>{ console.error(err); showNotification('PPTX export failed','error'); });
        return;
    }

    // ════════════════════════════════════════════════════════════════════
    // MODE: groups-and-tasks
    //   Group-name header row, then task rows underneath — no group tag column
    // ════════════════════════════════════════════════════════════════════
    // (mode === 'groups-and-tasks')

    const GROUP_HDR_WEIGHT = 0.70; // group header rows are a bit shorter than task rows
    const TASK_ROW_WEIGHT  = 1.00;

    // Build flat list: alternating {type:'group',...} and {type:'task',...}
    const allRows = [];
    ordered.forEach(g => {
        const gTasks = tasks.filter(t => t.group === g);
        if (!gTasks.length) return;
        const dates = gTasks.flatMap(t=>[ymd(t.startDate),ymd(t.endDate)]).filter(Boolean);
        const gStart = dates.length ? new Date(Math.min(...dates)) : null;
        const gEnd   = dates.length ? new Date(Math.max(...dates)) : null;
        allRows.push({ type:'group', group:g, count:gTasks.length, gStart, gEnd });
        gTasks.forEach(t => allRows.push({ type:'task', task:t, group:g }));
    });

    function rowW(r){ return r.type==='group' ? GROUP_HDR_WEIGHT : TASK_ROW_WEIGHT; }
    const totalWeight2 = allRows.reduce((a,r)=>a+rowW(r),0);

    const ROW_UNIT_MIN2=0.30, ROW_UNIT_MAX2=0.52;
    let rowUnit2 = Math.min(ROW_UNIT_MAX2, CONTENT_H/Math.max(1,totalWeight2));
    let pages2;
    if (rowUnit2 >= ROW_UNIT_MIN2) {
        pages2 = [allRows];
    } else {
        rowUnit2 = ROW_UNIT_MIN2;
        pages2 = []; let page=[], used=0;
        for (const r of allRows) {
            const h = rowW(r) * rowUnit2;
            if (used + h > CONTENT_H && page.length > 0) { pages2.push(page); page=[]; used=0; }
            page.push(r); used += h;
        }
        if (page.length > 0) pages2.push(page);
    }

    function drawGroupHeader(slide, row, y, h) {
        const gColor = (groups[row.group]?.color || '#475569').replace('#','');

        // Tinted background strip across entire label column
        slide.addShape(pptx.shapes.RECTANGLE,{
            x:GX, y:y+0.02, w:LABEL_COL_W, h:h-0.04,
            fill:{color:C.groupHdr}, line:{color:'D0DAF0', width:0.5}
        });
        // Colored left accent bar
        slide.addShape(pptx.shapes.RECTANGLE,{
            x:GX, y:y+0.02, w:0.12, h:h-0.04,
            fill:{color:gColor}, line:{color:gColor}
        });
        // Group name + count
        const taskCount = row.count > 0 ? ` (${row.count})` : '';
        slide.addText((row.group || '') + taskCount, {
            x:GX+0.18, y, w:LABEL_COL_W-0.20, h,
            fontSize:Math.min(12.5, h*30), bold:true, color:C.groupHdrText,
            fontFace:'Segoe UI', align:'left', valign:'middle',
            wrap:false, shrinkText:true
        });

        // Faint summary bar spanning the group's full date range
        if (row.gStart && row.gEnd) {
            const x1=dToX(row.gStart), x2=dToX(row.gEnd);
            if (x1!==null && x2!==null) {
                const bw=Math.max(0,x2-x1);
                const barH=Math.min(0.14, h*0.45);
                const barY=y+(h-barH)/2;
                if (bw >= 0.04) {
                    slide.addShape(pptx.shapes.RECTANGLE,{
                        x:x1, y:barY, w:bw, h:barH,
                        fill:{color:gColor, transparency:65},
                        line:{color:gColor, transparency:40, width:0.5}
                    });
                }
            }
        }
    }

    function drawTaskRowGT(slide, row, y, h) {
        const t = row.task;
        const tColor = (t.color||'#475569').replace('#','');
        const ts = ymd(t.startDate), te = ymd(t.endDate);
        const x1 = dToX(ts), x2 = dToX(te);
        if (x1===null||x2===null) return;
        const w = Math.max(0, x2-x1);
        const isZero = w < 0.04;
        const barH = Math.min(0.30, h*0.74);
        const barY = y + (h-barH)/2;

        // Task name — pre-fitted so it never clips or overflows into the timeline
        const labelMaxW = TX - GX - 0.22 - 0.20;
        const lbl = fitLabel((t.name||'').toUpperCase(), labelMaxW);
        slide.addText(lbl.text, {
            x:GX+0.22, y, w:labelMaxW, h,
            fontSize:lbl.fontSize, bold:true, color:C.ink,
            fontFace:'Segoe UI', align:'left', valign:'middle',
            wrap:false
        });

        if (isZero) {
            const sz=Math.min(0.18,barH);
            slide.addShape(pptx.shapes.DIAMOND,{
                x:x1-sz/2, y:barY+(barH-sz)/2, w:sz, h:sz,
                fill:{color:tColor}, line:{color:tColor}
            });
        } else {
            // Bar shape only — name lives in label column, no duplicate inside bar
            slide.addShape(pptx.shapes.PENTAGON,{
                x:x1, y:barY, w:Math.max(0.30,w), h:barH,
                fill:{color:tColor}, line:{color:tColor}
            });
        }
        // Date range
        slide.addText(`${fmtMD(ts)} – ${fmtMD(te)}`,{
            x:x2+0.08, y:barY, w:(W-MARGIN_R)-(x2+0.08), h:barH,
            fontSize:Math.min(9,barH*26), bold:true, color:tColor,
            fontFace:'Segoe UI', align:'left', valign:'middle', wrap:false
        });
    }

    pages2.forEach((rows, pi) => {
        const slide = pptx.addSlide();
        chrome(slide, pages2.length>1?`Page ${pi+1} / ${pages2.length}`:'');
        let y = CONTENT_TOP;
        rows.forEach(r => {
            const h = rowW(r) * rowUnit2;
            if (r.type === 'group') drawGroupHeader(slide, r, y, h);
            else drawTaskRowGT(slide, r, y, h);
            y += h;
        });
    });

    const projName2 = projects[activeProjectIndex]?.name || 'Gantt Chart';
    const safeName2 = projName2.replace(/[^\w\-]+/g,'_');
    pptx.writeFile({fileName:`${safeName2}_groups-and-tasks_${new Date().toISOString().slice(0,10)}.pptx`})
        .then(()=>showNotification('PPTX exported','success'))
        .catch(err=>{ console.error(err); showNotification('PPTX export failed','error'); });
}
