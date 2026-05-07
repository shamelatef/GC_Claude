/**
 * Export Gantt Chart to PowerPoint — Office Timeline style.
 *
 *   Section bars  =  tasks       (one per task; spans task start→end)
 *   Sub-arrows    =  subtasks    (split evenly across the parent task's span)
 *   Group         =  small left-margin category tag, no swimlane bar
 *
 * mode === 'tasks'                → only task bars (no subtask arrows below)
 * mode === 'tasks-with-subtasks'  → task bars + their split-out subtask arrows
 *
 * Requires PptxGenJS and the host page globals: tasks, milestones, groups,
 * groupOrder, projects, activeProjectIndex, parseYMD, showNotification.
 */
function exportToPPT(mode) {
    mode = mode === 'tasks-with-subtasks' ? 'tasks-with-subtasks' : 'tasks';
    const includeSubtasks = mode === 'tasks-with-subtasks';

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
    const GROUP_TAG_W = 0.85;   // small group-tag column on the far left
    const LABEL_COL_W = 1.55;   // task-name column
    const SIDE_COL_W  = 1.30;   // right column for date range / overflow

    const GX = MARGIN_L;
    const LX = GX + GROUP_TAG_W;
    const TX = LX + LABEL_COL_W;
    const TR = W - MARGIN_R - SIDE_COL_W;
    const TW = TR - TX;

    const PILL_H   = 0.32;
    const MS_AREA  = 1.05;
    const PILL_TOP = H - MARGIN_B - MS_AREA - PILL_H;
    const CONTENT_TOP = MARGIN_T + 0.05;
    const CONTENT_BOT = PILL_TOP - 0.18;
    const CONTENT_H   = CONTENT_BOT - CONTENT_TOP;

    const C = {
        bg:'FFFFFF', pill:'475569', guide:'D8DDE3', ink:'1A1A1A',
        muted:'9AA0A6', msLine:'CFD4DA', year:'E58A3A', sideText:'374151',
        groupTag:'6B7280'
    };

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

    // ── Date range ───────────────────────────────────────────────────────
    const allDates=[
        ...tasks.flatMap(t=>[ymd(t.startDate),ymd(t.endDate)]),
        ...(milestones||[]).map(m=>ymd(m.date))
    ].filter(Boolean);
    if(!allDates.length){ showNotification('No date data to export','error'); return; }
    const minDate=new Date(Math.min(...allDates));
    const maxDate=new Date(Math.max(...allDates));

    // ── Pick axis unit ───────────────────────────────────────────────────
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

    // ── Build flat row list (one entry per task), then paginate by row count
    //    A task contributes (1 + N subtasks if includeSubtasks) rows; the
    //    extra rows are NOT separate rows — they sit ON THE SAME row as the
    //    parent (sub-arrows are stacked below the task arrow). So each task
    //    is exactly ONE row regardless of subtask count, just with variable
    //    height. We just grow the row.
    // ────────────────────────────────────────────────────────────────────
    const ordered = (groupOrder.length ? groupOrder : Object.keys(groups))
        .filter(g => tasks.some(t => t.group === g));
    const rowsList = [];
    ordered.forEach(g=>{
        tasks.filter(t=>t.group===g).forEach(t=>{
            const subs = (includeSubtasks && Array.isArray(t.subtasks)) ? t.subtasks : [];
            rowsList.push({task:t, group:g, subs});
        });
    });

    // Each row has weight = 1 + 0.55 * subCount (subtasks are slimmer than parent)
    function rowWeight(r){
        return 1 + (r.subs.length ? 0.55*r.subs.length + 0.18 : 0);
    }
    const totalWeight = rowsList.reduce((a,r)=>a+rowWeight(r),0);

    // Try to fit everything on one slide by computing a row-unit height,
    // floor it; if too small, paginate.
    const ROW_UNIT_MIN = 0.34;   // base row height
    const ROW_UNIT_MAX = 0.55;
    let rowUnit = Math.min(ROW_UNIT_MAX, CONTENT_H/Math.max(1,totalWeight));
    let pages;
    if(rowUnit >= ROW_UNIT_MIN){
        pages = [rowsList];
    } else {
        rowUnit = ROW_UNIT_MIN;
        pages = [];
        let page=[], used=0;
        for(const r of rowsList){
            const h = rowWeight(r)*rowUnit;
            if(used + h > CONTENT_H && page.length>0){
                pages.push(page); page=[]; used=0;
            }
            page.push(r); used += h;
        }
        if(page.length>0) pages.push(page);
    }

    // ── Slide chrome ─────────────────────────────────────────────────────
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

        // Vertical guide lines per axis cell
        axis.cells.forEach((c,i)=>{
            if(i===0) return;
            const gx=dToX(c.start);
            slide.addShape(pptx.shapes.LINE,{
                x:gx, y:CONTENT_TOP-0.05, w:0, h:PILL_TOP-(CONTENT_TOP-0.05),
                line:{color:C.guide, width:0.75}
            });
        });

        // Pill
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE,{
            x:TX, y:PILL_TOP, w:TW, h:PILL_H, rectRadius:0.06,
            fill:{color:C.pill}, line:{color:C.pill}
        });
        const cellW = TW/axis.cells.length;
        axis.cells.forEach((c,i)=>{
            const cx=TX + i*cellW;
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
            const baseY=msTop + lane*0.46;
            slide.addShape(pptx.shapes.LINE,{
                x:mx, y:PILL_TOP+PILL_H, w:0, h:baseY-(PILL_TOP+PILL_H),
                line:{color:C.msLine, width:0.75, dashType:'dash'}
            });
            const sz=0.16;
            const color=(m.color||'#3B8FD9').replace('#','');
            const shape = m.shape==='star' ? pptx.shapes.STAR_5 : pptx.shapes.DIAMOND;
            slide.addShape(shape,{
                x:mx-sz/2, y:baseY-sz/2, w:sz, h:sz,
                fill:{color}, line:{color}
            });
            slide.addText(fmtDOWMD(d),{
                x:mx-0.7, y:baseY+sz/2+0.02, w:1.4, h:0.18,
                fontSize:8, color:C.muted, fontFace:'Segoe UI',
                align:'center', valign:'middle'
            });
            slide.addText(m.name||'',{
                x:mx-0.7, y:baseY+sz/2+0.18, w:1.4, h:0.20,
                fontSize:9.5, bold:true, color:C.ink, fontFace:'Segoe UI',
                align:'center', valign:'middle'
            });
        });

        const ts=new Date().toLocaleDateString('en-US',{month:'short', day:'numeric', year:'numeric'});
        slide.addText(`Generated ${ts}`,{
            x:0, y:H-0.22, w:W, h:0.2, fontSize:7, color:C.muted,
            fontFace:'Segoe UI', align:'center', valign:'middle'
        });
    }

    // ── Draw one task row (section bar + optional subtask arrows below) ─
    function drawRow(slide, row, y, h){
        const t=row.task;
        const tColor=(t.color||'#475569').replace('#','');
        const groupColor=(groups[row.group]?.color || '#9AA0A6').replace('#','');
        const ts=ymd(t.startDate), te=ymd(t.endDate);
        const x1=dToX(ts), x2=dToX(te);
        if(x1===null || x2===null) return;
        const w=Math.max(0, x2-x1);
        const isZero = w<0.04;

        // Group tag on far left (small, muted, in group color)
        slide.addShape(pptx.shapes.RECTANGLE,{
            x:GX, y:y+h*0.18, w:0.10, h:h*0.64,
            fill:{color:groupColor}, line:{color:groupColor}
        });
        slide.addText(row.group||'',{
            x:GX+0.14, y:y, w:GROUP_TAG_W-0.16, h:h,
            fontSize:Math.min(9, h*22), color:C.groupTag,
            fontFace:'Segoe UI', italic:true,
            align:'left', valign:'middle', wrap:false, shrinkText:true
        });

        // Task name on its label column (right-aligned)
        slide.addText((t.name||'').toUpperCase(),{
            x:LX, y:y, w:LABEL_COL_W-0.08, h:h,
            fontSize:Math.min(12, h*26), bold:true, color:C.ink,
            fontFace:'Segoe UI', align:'right', valign:'middle',
            wrap:false, shrinkText:true
        });

        // ── Section bar (the task itself) ──
        // Vertical area: parent occupies the TOP portion of the row,
        // subtasks stack underneath. If no subtasks, parent fills the row.
        const subs = row.subs;
        const hasSubs = subs.length>0;
        const parentH = hasSubs ? Math.min(0.34, h*0.42) : Math.min(0.34, h*0.78);
        const parentY = y + (hasSubs ? 0.04 : (h-parentH)/2);

        if(isZero){
            const sz=Math.min(0.18, parentH);
            slide.addShape(pptx.shapes.DIAMOND,{
                x:x1-sz/2, y:parentY+(parentH-sz)/2, w:sz, h:sz,
                fill:{color:tColor}, line:{color:tColor}
            });
        } else {
            // Big right-pointing pentagon for the task (section bar)
            slide.addShape(pptx.shapes.PENTAGON,{
                x:x1, y:parentY, w:Math.max(0.32,w), h:parentH,
                fill:{color:tColor}, line:{color:tColor}
            });
            const fs = Math.min(11, parentH*32);
            const charW = fs*0.0075;
            const minWForInside = (t.name||'').length*charW + 0.2;
            if(w >= minWForInside && w >= 0.6){
                slide.addText(t.name||'',{
                    x:x1+0.1, y:parentY, w:Math.max(0.16, w-0.28), h:parentH,
                    fontSize:fs, bold:true, color:'FFFFFF',
                    fontFace:'Segoe UI', align:'center', valign:'middle', wrap:false
                });
            }
        }

        // Date range to the right of the parent's tip
        const drText = `${fmtMD(ts)} – ${fmtMD(te)}`;
        slide.addText(drText,{
            x:x2+0.08, y:parentY, w:(W-MARGIN_R)-(x2+0.08), h:parentH,
            fontSize:Math.min(10, parentH*26), bold:true, color:tColor,
            fontFace:'Segoe UI', align:'left', valign:'middle', wrap:false
        });

        // ── Sub-arrows: split parent's span evenly across subtask count ──
        if(hasSubs && !isZero){
            const subsTop = parentY + parentH + 0.04;
            const subsAvailH = (y+h) - subsTop - 0.02;
            const subH = Math.max(0.14, Math.min(0.22, subsAvailH/Math.max(1,subs.length)));
            const tsMs = ts.getTime(), teMs = te.getTime();
            const totalMs = teMs - tsMs;
            const step = totalMs / subs.length;
            subs.forEach((sub, si)=>{
                const sStart = new Date(tsMs + si*step);
                const sEnd   = new Date(tsMs + (si+1)*step);
                const sx1 = dToX(sStart), sx2 = dToX(sEnd);
                const sw = Math.max(0, sx2-sx1);
                const sy = subsTop + si*subH + (subH - Math.min(subH*0.78, 0.18))/2;
                const sah = Math.min(subH*0.78, 0.18);

                // Subtle alternating color: parent color slightly desaturated
                const subColor = sub.done ? tColor : tColor;
                const shape = si===0 ? pptx.shapes.PENTAGON : pptx.shapes.CHEVRON;
                slide.addShape(shape,{
                    x:sx1, y:sy, w:Math.max(0.22, sw), h:sah,
                    fill:{color:subColor, transparency: sub.done ? 0 : 35},
                    line:{color:subColor, transparency: sub.done ? 0 : 20}
                });

                // Label inside if it fits
                const fs = Math.min(9, sah*32);
                const charW = fs*0.0075;
                const minWForInside = (sub.name||'').length*charW + 0.18;
                const labelOpts = {
                    fontSize:fs, bold:true, fontFace:'Segoe UI',
                    align:'center', valign:'middle', wrap:false,
                    italic: !!sub.done, strike: !!sub.done
                };
                if(sw >= minWForInside && sw >= 0.5){
                    slide.addText(sub.name||'',{
                        ...labelOpts,
                        x:sx1+0.06, y:sy, w:Math.max(0.16, sw-0.18), h:sah,
                        color:'FFFFFF',
                        align:'center'
                    });
                } else {
                    slide.addText(sub.name||'',{
                        ...labelOpts,
                        x:sx2+0.06, y:sy-0.02, w:(W-MARGIN_R)-(sx2+0.06), h:sah+0.04,
                        color: sub.done ? C.muted : C.sideText,
                        align:'left', bold:false
                    });
                }
            });
        }
    }

    // ── Build slides ─────────────────────────────────────────────────────
    pages.forEach((rows, pi)=>{
        const slide = pptx.addSlide();
        const lbl = pages.length>1 ? `Page ${pi+1} / ${pages.length}` : '';
        chrome(slide, lbl);

        let y = CONTENT_TOP;
        rows.forEach(r=>{
            const h = rowWeight(r) * rowUnit;
            drawRow(slide, r, y, h);
            y += h;
        });
    });

    // ── Save ─────────────────────────────────────────────────────────────
    const projName = projects[activeProjectIndex]?.name || 'Gantt Chart';
    const safeName = projName.replace(/[^\w\-]+/g,'_');
    const modeSuffix = includeSubtasks ? '_with-subtasks' : '_tasks';
    pptx.writeFile({fileName:`${safeName}${modeSuffix}_${new Date().toISOString().slice(0,10)}.pptx`})
        .then(()=>showNotification('PPTX exported','success'))
        .catch(err=>{ console.error(err); showNotification('PPTX export failed','error'); });
}
