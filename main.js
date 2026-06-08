// Global state for currently active project
let tasks = [];
let groups = {};
let groupStates = {};
let groupOrder = [];
let milestones = [];
let draggedElement = null;
let draggedType = null;
let hasUnsavedChanges = false;
let lastSavedData = null;
// Track previously typed group names (not only the ones created via tasks)
let previousGroupInputs = new Set();
// Drag/resize state for task bars
let taskDragState = null; // { type: 'move'|'resize', side?: 'left'|'right', taskId, startX, chartMin, chartMax, pxPerDay, origStart, origEnd, barEl, trackEl, trackRect }


// Export a 16:9 PNG (1920x1080) with ALL tasks visible by vertically fitting content
function saveChartForPPTFitAll() {
    const exportRoot = document.querySelector('.container');
    const target = document.querySelector('.chart-container') || document.getElementById('ganttChart');
    if (!target) {
        showNotification('Nothing to export yet', 'error');
        return;
    }

    // Snapshot current group expansion state and expand all to ensure all tasks are visible
    const prevStates = { ...groupStates };
    try {
        Object.keys(groups || {}).forEach(g => { groupStates[g] = true; });
        updateChart();
    } catch (_) {}

    // Next frame, measure and apply vertical fit transform, then export
    requestAnimationFrame(() => {
        try {
            exportRoot.classList.add('export-mode', 'ppt-export');
            // Measure full content height of the chart area
            const rect = target.getBoundingClientRect();
            const contentHeight = Math.max(rect.height, target.scrollHeight || 0);
            const OUT_W = 1920;
            const OUT_H = 1080;
            // Compute vertical scale to fit content height into 1080px
            const scaleY = contentHeight > 0 ? Math.min(1, OUT_H / contentHeight) : 1;
            // Apply only vertical scaling; keep width natural
            const prevTransform = target.style.transform || '';
            const prevTransformOrigin = target.style.transformOrigin || '';
            target.style.transformOrigin = 'top left';
            target.style.transform = `${prevTransform ? prevTransform + ' ' : ''}scaleY(${scaleY})`;

            html2canvas(target, {
                backgroundColor: null,
                scale: window.devicePixelRatio > 1 ? 2 : 2,
                useCORS: true,
                logging: false
            }).then(canvas => {
                // Draw onto an exact 1920x1080 canvas with letterboxing
                const OUT_CANVAS = document.createElement('canvas');
                OUT_CANVAS.width = OUT_W;
                OUT_CANVAS.height = OUT_H;
                const ctx = OUT_CANVAS.getContext('2d');
                ctx.fillStyle = '#ffffff';
                try { ctx.fillStyle = getComputedStyle(document.body).backgroundColor || '#ffffff'; } catch (_) {}
                ctx.fillRect(0, 0, OUT_W, OUT_H);
                const scale = Math.min(OUT_W / canvas.width, OUT_H / canvas.height);
                const drawW = Math.max(1, Math.round(canvas.width * scale));
                const drawH = Math.max(1, Math.round(canvas.height * scale));
                const dx = Math.floor((OUT_W - drawW) / 2);
                const dy = Math.floor((OUT_H - drawH) / 2);
                ctx.imageSmoothingEnabled = true;
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(canvas, dx, dy, drawW, drawH);

                const link = document.createElement('a');
                const ts = new Date().toISOString().replace(/[:.]/g, '-');
                link.download = `gantt-ppt-fit-all-${ts}.png`;
                link.href = OUT_CANVAS.toDataURL('image/png');
                link.click();
                showNotification('Exported for PPT (fit all tasks)', 'success');
            }).catch(err => {
                console.error('Export failed:', err);
                showNotification('Export for PPT failed', 'error');
            }).finally(() => {
                // Cleanup transform and classes
                try {
                    target.style.transform = prevTransform;
                    target.style.transformOrigin = prevTransformOrigin;
                    exportRoot.classList.remove('export-mode', 'ppt-export');
                } catch (_) {}
                // Restore prior group expansion and re-render
                try {
                    Object.keys(prevStates).forEach(g => { groupStates[g] = prevStates[g]; });
                    updateChart();
                } catch (_) {}
            });
        } catch (e) {
            console.error(e);
            showNotification('Export for PPT failed', 'error');
            try { exportRoot.classList.remove('export-mode', 'ppt-export'); } catch (_) {}
            try { Object.keys(prevStates).forEach(g => { groupStates[g] = prevStates[g]; }); updateChart(); } catch (_) {}
        }
    });
}

// Prevent single-click toggle from firing when a double-click rename is intended
const groupClickTimers = new Map();
function onGroupHeaderClick(event, groupName) {
    // Ignore clicks coming from kebab menu
    if (event.target && event.target.closest && event.target.closest('.kebab-menu')) return;
    if (event.detail === 1) {
        // Debounce: wait briefly; if dblclick happens, this will be canceled
        const existing = groupClickTimers.get(groupName);
        if (existing) clearTimeout(existing);
        const t = setTimeout(() => {
            toggleGroup(groupName);
            groupClickTimers.delete(groupName);
        }, 220);
        groupClickTimers.set(groupName, t);
    } else if (event.detail > 1) {
        const existing = groupClickTimers.get(groupName);
        if (existing) {
            clearTimeout(existing);
            groupClickTimers.delete(groupName);
        }
    }
}
function onGroupLabelDblClick(event, groupName) {
    if (event && event.stopPropagation) event.stopPropagation();
    const existing = groupClickTimers.get(groupName);
    if (existing) {
        clearTimeout(existing);
        groupClickTimers.delete(groupName);
    }
    openGroupEditModal(groupName);
}
function updateGroupStatus(groupName, status) {
    if (!groups[groupName]) return;
    groups[groupName].status = status;
    // Derive group color from status to keep visual consistent
    groups[groupName].color = STATUS_COLORS[status] || '#808080';
    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification(`Group "${groupName}" set to ${status}`, 'success');
}

// Add Milestone Modal Logic
let addMilestoneModalKeyHandlerBound = false;
// Edit Milestone Modal Logic
let editMilestoneId = null;
let editMilestoneModalKeyHandlerBound = false;
let milestoneClickBound = false;

// Status to color mapping (six fixed statuses)
const STATUS_COLORS = {
    'Not Started': '#808080',
    'In Progress': '#4C9141',
    'Delayed': '#FFA500',
    'Blocked': '#CC0000',
    'Action Needed': '#4A0072',
    'Completed': '#00B4D8'
};

// Auto-derive a group's status from its tasks by priority
// Priority (highest to lowest): Blocked > Action Needed > Delayed > In Progress > Completed > Not Started
function deriveGroupStatusFromTasks(groupName) {
    const groupTasks = tasks.filter(t => t.group === groupName);
    if (groupTasks.length === 0) return undefined;

    // A task with numeric progress > 0 is treated as at least In Progress
    const statuses = new Set(groupTasks.map(t => {
        if (typeof t.progress === 'number' && t.progress >= 100) return 'Completed';
        if (typeof t.progress === 'number' && t.progress > 0)    return 'In Progress';
        return t.status || 'Not Started';
    }));
    if (statuses.has('Blocked'))        return 'Blocked';
    if (statuses.has('Action Needed'))  return 'Action Needed';
    if (statuses.has('Delayed'))        return 'Delayed';
    if (statuses.has('In Progress'))    return 'In Progress';
    if (statuses.has('Completed') && statuses.has('Not Started')) return 'In Progress';
    if (statuses.has('Completed') && statuses.size === 1)         return 'Completed';
    return 'Not Started';
}

function applyAutoGroupStatus(groupName, { notify = false } = {}) {
    if (!groups[groupName]) return;
    const auto = deriveGroupStatusFromTasks(groupName);
    if (!auto) return;
    if (groups[groupName].status !== auto) {
        groups[groupName].status = auto;
        groups[groupName].color = STATUS_COLORS[auto] || '#808080';
        if (notify) showNotification(`Group "${groupName}" status updated to ${auto}`, 'info');
    }
}

function syncStatusSelection(selectId, chipId, colorInputId, previewId) {
    const sel = document.getElementById(selectId);
    const chip = document.getElementById(chipId);
    if (!sel || !chip) return;
    const hex = STATUS_COLORS[sel.value] || '#808080';
    chip.style.background = hex;
    const colorInput = colorInputId ? document.getElementById(colorInputId) : null;
    if (colorInput) colorInput.value = hex;
    try { if (colorInput && previewId) updateColorPreview(colorInputId, previewId); } catch (_) {}
}

function openAddMilestoneModal() {
    const backdrop = document.getElementById('addMilestoneModal');
    const nameInput = document.getElementById('addMilestoneName');
    const dateInput = document.getElementById('addMilestoneDate');
    if (!backdrop || !nameInput || !dateInput) return;
    // default date to today if empty
    if (!dateInput.value) {
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        dateInput.value = `${yyyy}-${mm}-${dd}`;
    }
    backdrop.style.display = 'flex';
    setTimeout(() => { nameInput.focus(); nameInput.select(); }, 0);
    if (!addMilestoneModalKeyHandlerBound) {
        addMilestoneModalKeyHandlerBound = true;
        backdrop.addEventListener('keydown', function(e){
            if (backdrop.style.display === 'none') return;
            if (e.key === 'Enter') { e.preventDefault(); confirmAddMilestoneModal(); }
            if (e.key === 'Escape') { e.preventDefault(); closeAddMilestoneModal(); }
        });
    }
}

function closeAddMilestoneModal() {
    const backdrop = document.getElementById('addMilestoneModal');
    const nameInput = document.getElementById('addMilestoneName');
    const dateInput = document.getElementById('addMilestoneDate');
    if (!backdrop) return;
    backdrop.style.display = 'none';
    if (nameInput) { nameInput.value = ''; nameInput.classList.remove('input-error'); }
    if (dateInput) dateInput.value = '';
}

function confirmAddMilestoneModal() {
    const nameInput = document.getElementById('addMilestoneName');
    const dateInput = document.getElementById('addMilestoneDate');
    if (!nameInput || !dateInput) return;
    const name = (nameInput.value || '').trim();
    const date = dateInput.value;
    if (!name) {
        nameInput.classList.add('input-error');
        nameInput.focus();
        return;
    }
    if (!date) {
        showNotification('Please choose a milestone date', 'error');
        dateInput.focus();
        return;
    }
    const id = milestones.length ? Math.max(...milestones.map(m => m.id || 0)) + 1 : 1;
    milestones.push({ id, name, date });
    markAsChanged();
    updateChart();
    closeAddMilestoneModal();
    showNotification('Milestone added', 'success');
}

// Render milestones as vertical lines with diamond labels across the chart body
function renderMilestonesOverlay() {
    if (!Array.isArray(milestones) || milestones.length === 0 || !chartMinDate || !chartMaxDate) return '';
    // Use the same month-based layout math as task/group bars so markers align perfectly
    const months = getMonthRange(chartMinDate, chartMaxDate);
    const totalMonths = Math.max(1, months.length);
    const monthWidth = 100 / totalMonths;

    const items = milestones
        .filter(m => m && m.date && parseYMD(m.date))
        .map(m => {
            const dObj = parseYMD(m.date);
            if (!dObj) return '';
            const y = dObj.getFullYear();
            const mIdx = dObj.getMonth() + 1; // 1..12
            const day = dObj.getDate();
            const monthIndex = months.findIndex(mm => mm.year === y && mm.month === mIdx);
            if (monthIndex === -1) return '';
            const daysInMonth = new Date(y, mIdx, 0).getDate();
            // Position at the start of the chosen day (same convention as task bar starts)
            const startOffset = ((day - 1) / daysInMonth) * monthWidth;
            const left = (monthIndex * monthWidth) + startOffset;
            const pct = Math.max(0, Math.min(100, left));
            const title = `${m.name || 'Milestone'}: ${formatDate(m.date)}`;
            const idAttr = m.id != null ? `data-milestone-id="${m.id}"` : '';
            return `
                <div class="milestone-marker" style="left:${pct}%" title="${title}" ${idAttr}>
                    <div class="milestone-line"></div>
                    <div class="milestone-badge">${escapeHtml(m.name || '')}</div>
                </div>
            `;
        }).join('');
    return items;
}

// Render a red dotted "today" vertical marker across the chart body
function renderTodayMarker() {
    if (!chartMinDate || !chartMaxDate) return '';
    // Normalize today to local Y-M-D (strip time)
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    // Only show if today falls within the current chart bounds
    if (today < chartMinDate || today > chartMaxDate) return '';

    const months = getMonthRange(chartMinDate, chartMaxDate);
    const totalMonths = Math.max(1, months.length);
    const monthWidth = 100 / totalMonths;

    const y = today.getFullYear();
    const m = today.getMonth() + 1; // 1-based month to match months[] model
    const d = today.getDate();
    const monthIndex = months.findIndex(mm => mm.year === y && mm.month === m);
    if (monthIndex === -1) return '';

    const daysInMonth = new Date(y, m, 0).getDate();
    const startOffset = ((d - 1) / daysInMonth) * monthWidth;
    const left = Math.max(0, Math.min(100, (monthIndex * monthWidth) + startOffset));

    return `
        <div class="today-marker" style="left:${left}%">
            <div class="today-line"></div>
            <div class="today-label" title="Today">Today</div>
        </div>
    `;
}

// Render a status legend strip below the chart body
function renderStatusLegend() {
    const entries = Object.entries(STATUS_COLORS).map(([label, color]) => `
        <span class="legend-item">
            <span class="legend-dot" style="background:${color}"></span>
            <span class="legend-text">${label}</span>
        </span>
    `).join('');
    return `<div id="statusLegend" class="status-legend">${entries}</div>`;
}

// Click handling for milestones (open edit modal)
function setupMilestoneInteractions() {
    const chartContainer = document.getElementById('ganttChart');
    if (!chartContainer) return;
    if (milestoneClickBound) return;
    chartContainer.addEventListener('click', (e) => {
        const marker = e.target.closest('.milestone-marker');
        if (!marker) return;
        const idStr = marker.getAttribute('data-milestone-id');
        const id = idStr ? parseInt(idStr, 10) : null;
        if (id) openEditMilestoneModal(id);
    });
    milestoneClickBound = true;
}

// Open/Edit/Delete Milestone Modal handlers
function openEditMilestoneModal(id) {
    const backdrop = document.getElementById('editMilestoneModal');
    const nameInput = document.getElementById('editMilestoneName');
    const dateInput = document.getElementById('editMilestoneDate');
    const m = milestones.find(mm => mm.id === id);
    if (!backdrop || !nameInput || !dateInput || !m) return;
    editMilestoneId = id;
    nameInput.value = m.name || '';
    // ensure date is yyyy-mm-dd
    try {
        const d = parseYMD(m.date) || new Date(m.date);
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        dateInput.value = `${yyyy}-${mm}-${dd}`;
    } catch(_) { dateInput.value = m.date || ''; }
    backdrop.style.display = 'flex';
    setTimeout(() => { nameInput.focus(); nameInput.select(); }, 0);

    // Backdrop click to close
    backdrop.addEventListener('click', function(e){ if (e.target === backdrop) closeEditMilestoneModal(); });

    // Keyboard shortcuts
    if (!editMilestoneModalKeyHandlerBound) {
        editMilestoneModalKeyHandlerBound = true;
        backdrop.addEventListener('keydown', function(e){
            if (backdrop.style.display === 'none') return;
            if (e.key === 'Enter') { e.preventDefault(); confirmEditMilestoneModal(); }
            if (e.key === 'Escape') { e.preventDefault(); closeEditMilestoneModal(); }
        });
    }
}

function closeEditMilestoneModal() {
    const backdrop = document.getElementById('editMilestoneModal');
    const nameInput = document.getElementById('editMilestoneName');
    const dateInput = document.getElementById('editMilestoneDate');
    if (!backdrop) return;
    backdrop.style.display = 'none';
    editMilestoneId = null;
    if (nameInput) nameInput.value = '';
    if (dateInput) dateInput.value = '';
}

function confirmEditMilestoneModal() {
    if (editMilestoneId == null) return;
    const nameInput = document.getElementById('editMilestoneName');
    const dateInput = document.getElementById('editMilestoneDate');
    if (!nameInput || !dateInput) return;
    const name = (nameInput.value || '').trim();
    const date = dateInput.value;
    if (!name) { nameInput.classList.add('input-error'); nameInput.focus(); return; }
    if (!date) { showNotification('Please choose a milestone date', 'error'); dateInput.focus(); return; }
    const idx = milestones.findIndex(m => m.id === editMilestoneId);
    if (idx === -1) return;
    milestones[idx].name = name;
    milestones[idx].date = date;
    markAsChanged();
    updateChart();
    closeEditMilestoneModal();
    showNotification('Milestone updated', 'success');
}

function confirmDeleteMilestone() {
    if (editMilestoneId == null) return;
    const m = milestones.find(mm => mm.id === editMilestoneId);
    if (!m) { closeEditMilestoneModal(); return; }
    if (!confirm(`Delete milestone "${m.name}"?`)) return;
    milestones = milestones.filter(mm => mm.id !== editMilestoneId);
    editMilestoneId = null;
    markAsChanged();
    updateChart();
    closeEditMilestoneModal();
    showNotification('Milestone deleted', 'success');
}

// Simple HTML escape for labels
function escapeHtml(s) {
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function updateGroupColor(groupName, color) {
    if (!groups[groupName]) return;
    groups[groupName].color = color;
    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification('Group color updated', 'success');
}
// Timeline bounds for pixel-to-day mapping
let chartMinDate = null; // Date
let chartMaxDate = null; // Date
let taskResizeWindowBound = false; // ensure window listeners are added only once
let taskResizeChartBound = false;  // ensure chart mousedown listener is added only once
// Throttle timestamp for timeline auto-extend during drag
let taskAutoExtendThrottle = 0;

// Multi-project management
let projects = [];
let activeProjectIndex = 0;

function makeEmptyProject(name = 'Project 1') {
    return {
        name,
        tasks: [],
        milestones: [],
        groups: {},
        groupStates: {},
        groupOrder: []
    };
}

function syncActiveFromGlobals() {
    if (!projects[activeProjectIndex]) return;
    projects[activeProjectIndex].tasks = JSON.parse(JSON.stringify(tasks));
    projects[activeProjectIndex].milestones = JSON.parse(JSON.stringify(milestones));
    projects[activeProjectIndex].groups = JSON.parse(JSON.stringify(groups));
    projects[activeProjectIndex].groupStates = JSON.parse(JSON.stringify(groupStates));
    projects[activeProjectIndex].groupOrder = JSON.parse(JSON.stringify(groupOrder));
}

function syncGlobalsFromActive() {
    const p = projects[activeProjectIndex];
    if (!p) return;
    tasks = JSON.parse(JSON.stringify(p.tasks || []));
    milestones = JSON.parse(JSON.stringify(p.milestones || []));
    groups = JSON.parse(JSON.stringify(p.groups || {}));
    groupStates = JSON.parse(JSON.stringify(p.groupStates || {}));
    groupOrder = JSON.parse(JSON.stringify(p.groupOrder || Object.keys(p.groups || {})));
    // Normalize: any task with progress > 0 is In Progress; progress >= 100 is Completed
    tasks.forEach(t => {
        if (typeof t.progress !== 'number') return;
        if (t.progress >= 100 && t.status !== 'Completed') {
            t.status = 'Completed';
            t.color = STATUS_COLORS['Completed'];
        } else if (t.progress > 0 && t.progress < 100) {
            t.status = 'In Progress';
            t.color = STATUS_COLORS['In Progress'];
        }
    });
    // Re-derive group statuses so groups with any task progress > 0 turn green too
    Object.keys(groups).forEach(g => applyAutoGroupStatus(g));
}

// Generic edit modal logic for task/group name edits
let editContext = null; // { type: 'task'|'group', taskId?: number, groupName?: string }
function openTaskEditModal(taskId) {
    const t = (tasks || []).find(x => x.id === taskId);
    if (!t) return;
    editContext = { type: 'task', taskId };
    const title = document.getElementById('editTitle');
    const input = document.getElementById('editInput');
    const modal = document.getElementById('editModal');
    if (!title || !input || !modal) return;
    title.textContent = 'Edit Task Name';
    input.value = t.name || '';
    modal.style.display = 'flex';
    input.focus();
    input.select();
}
function openGroupEditModal(groupName) {
    if (!groupName) return;
    editContext = { type: 'group', groupName };
    const title = document.getElementById('editTitle');
    const input = document.getElementById('editInput');
    const modal = document.getElementById('editModal');
    if (!title || !input || !modal) return;
    title.textContent = 'Edit Group Name';
    input.value = groupName || '';
    modal.style.display = 'flex';
    input.focus();
    input.select();
}
function closeEditModal() {
    const modal = document.getElementById('editModal');
    if (modal) modal.style.display = 'none';
    editContext = null;
}
function confirmEditModal() {
    const input = document.getElementById('editInput');
    const val = input ? input.value : '';
    if (!editContext) { closeEditModal(); return; }
    if (editContext.type === 'task' && typeof updateTaskField === 'function') {
        updateTaskField(editContext.taskId, 'name', val);
    } else if (editContext.type === 'group' && typeof updateGroupField === 'function') {
        updateGroupField(editContext.groupName, 'name', val);
    }
    closeEditModal();
}
// Close modal on Escape, save on Enter
document.addEventListener('keydown', function(e){
    const modal = document.getElementById('editModal');
    if (!modal || modal.style.display === 'none') return;
    if (e.key === 'Escape') { e.preventDefault(); closeEditModal(); }
    if (e.key === 'Enter') { e.preventDefault(); confirmEditModal(); }
});

// Kebab menu handling (portal dropdown to body to avoid overlap/clipping)
// Global utility to close and restore a single dropdown
function closeKebabDropdown(btn, dd) {
    if (!dd) return;
    try {
        dd.classList.remove('show');
        // Remove hover-away handlers
        if (dd._hoverBound) {
            const { startHoverClose, cancelHoverClose, btnRef } = dd._hoverHandlers || {};
            if (startHoverClose && cancelHoverClose) {
                dd.removeEventListener('mouseenter', cancelHoverClose);
                dd.removeEventListener('mouseleave', startHoverClose);
                if (btnRef) {
                    btnRef.removeEventListener('mouseenter', cancelHoverClose);
                    btnRef.removeEventListener('mouseleave', startHoverClose);
                }
            }
            if (dd._hoverTimeout) {
                clearTimeout(dd._hoverTimeout);
                dd._hoverTimeout = null;
            }
            dd._hoverBound = false;
            dd._hoverHandlers = null;
        }
        // Restore to original container (button's parent .kebab-menu)
        if (btn && btn.parentElement && dd.parentElement !== btn.parentElement) {
            btn.parentElement.appendChild(dd);
        }
        // Clear inline styles
        dd.style.left = '';
        dd.style.top = '';
        dd.style.right = '';
        dd.style.position = '';
        dd.style.display = '';
        dd.style.transform = '';
        dd.style.visibility = '';
        // Remove z-boost on owner
        const own = btn ? (btn.closest('.chart-row') || btn.closest('.chart-group')) : null;
        if (own) own.classList.remove('has-open-menu');
        // Remove bound reposition handlers
        if (dd._repositionHandler) {
            window.removeEventListener('scroll', dd._repositionHandler, true);
            window.removeEventListener('resize', dd._repositionHandler, true);
            dd._repositionHandler = null;
        }
        if (btn) btn.removeAttribute('aria-expanded');
    } catch(_) {}
}

function closeAllKebabDropdowns() {
    document.querySelectorAll('.kebab-dropdown.show').forEach(dd => {
        const btn = dd._ownerButton || dd.previousElementSibling || document.querySelector('.kebab-icon[aria-expanded="true"]');
        closeKebabDropdown(btn, dd);
    });
}

function toggleKebabMenu(button) {

    // Close all other open dropdowns
    closeAllKebabDropdowns();

    // Toggle the clicked dropdown
    const dropdown = button.nextElementSibling;
    if (!dropdown) return;
    const isOpen = dropdown.classList.contains('show');
    if (isOpen) {
        closeKebabDropdown(button, dropdown);
        return;
    }

    // Mark owner and raise z-index
    const owner = button.closest('.chart-row') || button.closest('.chart-group');
    if (owner) owner.classList.add('has-open-menu');

    // Make visible and portal to body
    dropdown.classList.add('show');
    dropdown.style.visibility = 'hidden';
    dropdown.style.display = 'block';
    // Store reference to the button for closing
    dropdown._ownerButton = button;
    document.body.appendChild(dropdown);

    const positionDropdown = () => {
        const rect = button.getBoundingClientRect();
        const ddWidth = Math.max(160, dropdown.offsetWidth || 160);
        const ddHeight = Math.max(80, dropdown.offsetHeight || 100);
        const pad = 8;
        const spaceBelow = window.innerHeight - rect.bottom;
        const openAbove = spaceBelow < ddHeight + 12;
        let top = openAbove ? (rect.top - ddHeight - 6) : (rect.bottom + 6);
        let left = rect.right - ddWidth;
        left = Math.max(pad, Math.min(left, window.innerWidth - ddWidth - pad));
        top = Math.max(pad, Math.min(top, window.innerHeight - ddHeight - pad));
        dropdown.style.position = 'fixed';
        dropdown.style.left = left + 'px';
        dropdown.style.top = top + 'px';
        dropdown.style.right = 'auto';
        dropdown.style.transform = 'none';
        dropdown.style.visibility = '';
    };

    positionDropdown();
    // Reposition on scroll/resize
    dropdown._repositionHandler = () => positionDropdown();
    window.addEventListener('scroll', dropdown._repositionHandler, true);
    window.addEventListener('resize', dropdown._repositionHandler, true);

    // Hover-away to close, with small grace period for moving cursor
    const startHoverClose = () => {
        if (dropdown._hoverTimeout) clearTimeout(dropdown._hoverTimeout);
        dropdown._hoverTimeout = setTimeout(() => closeKebabDropdown(button, dropdown), 250);
    };
    const cancelHoverClose = () => {
        if (dropdown._hoverTimeout) {
            clearTimeout(dropdown._hoverTimeout);
            dropdown._hoverTimeout = null;
        }
    };
    dropdown.addEventListener('mouseenter', cancelHoverClose);
    dropdown.addEventListener('mouseleave', startHoverClose);
    button.addEventListener('mouseenter', cancelHoverClose);
    button.addEventListener('mouseleave', startHoverClose);
    dropdown._hoverBound = true;
    dropdown._hoverHandlers = { startHoverClose, cancelHoverClose, btnRef: button };

    // Accessibility toggle
    button.setAttribute('aria-expanded', 'true');

    // Close dropdown when clicking outside
    const closeHandler = (e) => {
        if (!dropdown.contains(e.target) && e.target !== button) {
            document.removeEventListener('click', closeHandler);
            closeKebabDropdown(button, dropdown);
        }
    };
    setTimeout(() => document.addEventListener('click', closeHandler), 0);
}

// Close dropdowns when clicking outside
document.addEventListener('click', (e) => {
    if (!e.target.closest('.kebab-menu')) {
        closeAllKebabDropdowns();
    }
});

// Close ALL popups helper
function closeAllPopups() {
    try { closeAllKebabDropdowns(); } catch(_) {}
    try { closeEditModal(); } catch(_) {}
    try { closeColorModal(); } catch(_) {}
    try { closeAddTaskModal(); } catch(_) {}
    try { closeRenameModal(); } catch(_) {}
}

// Global Escape closes everything
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        closeAllPopups();
    }
});

// Click on modal backdrops closes the modal
(function setupModalBackdropClose(){
    const map = [
        ['editModal', () => { try { closeEditModal(); } catch(_) {} }],
        ['colorModal', () => { try { closeColorModal(); } catch(_) {} }],
        ['addTaskModal', () => { try { closeAddTaskModal(); } catch(_) {} }],
        ['renameModal', () => { try { closeRenameModal(); } catch(_) {} }],
        ['addMilestoneModal', () => { try { closeAddMilestoneModal(); } catch(_) {} }]
    ];
    map.forEach(([id, closer]) => {
        const el = document.getElementById(id);
        if (el && !el._backdropBound) {
            el.addEventListener('mousedown', (e) => { if (e.target === el) closer(); });
            el._backdropBound = true;
        }
    });
})();

// Utility: open a hidden native date picker and return the selected value
function openHiddenDatePicker(initialValue, onSelect) {
    const input = document.createElement('input');
    input.type = 'date';
    if (initialValue) input.value = initialValue;
    // Keep in viewport (some browsers ignore clicks on offscreen elements)
    input.style.position = 'fixed';
    input.style.left = '10px';
    input.style.top = '10px';
    input.style.width = '1px';
    input.style.height = '1px';
    input.style.opacity = '0';
    input.style.zIndex = '2147483647';
    document.body.appendChild(input);

    const cleanup = () => { try { document.body.removeChild(input); } catch (_) {} };
    const handleChange = (e) => { try { if (onSelect) onSelect(e.target.value); } finally { cleanup(); } };
    input.addEventListener('change', handleChange, { once: true });
    input.addEventListener('blur', cleanup, { once: true });
    // Ensure focus, then try showPicker if available; otherwise click
    input.focus({ preventScroll: true });
    setTimeout(() => {
        try {
            if (typeof input.showPicker === 'function') {
                input.showPicker();
            } else {
                input.click();
            }
        } catch (_) {
            // Fallback to click if showPicker fails
            try { input.click(); } catch (_) {}
        }
    }, 0);
}

// Close the specific kebab dropdown when mouse leaves the menu area
document.addEventListener('mouseout', (e) => {
    const menu = e.target.closest('.kebab-menu');
    if (!menu) return;
    const toEl = e.relatedTarget;
    // If the mouse moved to an element outside this menu, close it
    if (!toEl || !menu.contains(toEl)) {
        const dd = menu.querySelector('.kebab-dropdown');
        if (dd) dd.classList.remove('show');
    }
});

function renderProjectTabs() {
    const el = document.getElementById('projectTabs');
    if (!el) return;
    el.innerHTML = '';
    projects.forEach((p, idx) => {
        const tab = document.createElement('button');
        tab.className = 'project-tab' + (idx === activeProjectIndex ? ' active' : '');
        tab.innerHTML = `<span class="name">${p.name || 'Untitled'}</span>` +
                        ` <span class="rename">✎</span>`;
        tab.title = 'Click to switch. Double-click or click ✎ to rename.';
        tab.onclick = () => switchProject(idx);
        tab.ondblclick = () => openRenameModal(idx);
        tab.querySelector('.rename').onclick = (e) => { e.stopPropagation(); openRenameModal(idx); };
        const delBtn = document.createElement('button');
        delBtn.className = 'delete';
        delBtn.type = 'button';
        delBtn.title = 'Delete project';
        delBtn.setAttribute('aria-label', 'Delete project');
        delBtn.textContent = '×';
        delBtn.onclick = (e) => { e.stopPropagation(); deleteProject(idx); };
        const dupBtn = document.createElement('button');
        dupBtn.className = 'duplicate';
        dupBtn.type = 'button';
        dupBtn.title = 'Duplicate project';
        dupBtn.setAttribute('aria-label', 'Duplicate project');
        dupBtn.textContent = '⎘';
        dupBtn.onclick = (e) => { e.stopPropagation(); duplicateProject(idx); };
        tab.appendChild(dupBtn);
        tab.appendChild(delBtn);
        el.appendChild(tab);
    });
    const addBtn = document.createElement('button');
    addBtn.className = 'project-tab add';
    addBtn.textContent = '+ New Project';
    addBtn.onclick = addProject;
    el.appendChild(addBtn);
}

function addProject() {
    const base = 'Project ';
    let n = projects.length + 1;
    const defaultName = base + n;
    projects.push(makeEmptyProject(defaultName));
    activeProjectIndex = projects.length - 1;
    syncGlobalsFromActive();
    renderProjectTabs();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    markAsChanged();
}

function switchProject(index) {
    if (index === activeProjectIndex) return;
    // Save current into active project
    syncActiveFromGlobals();
    activeProjectIndex = index;
    syncGlobalsFromActive();
    renderProjectTabs();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    markAsSaved(); // switching shows saved state for the switched content
}

function deleteProject(index) {
    if (!projects[index]) return;
    const name = projects[index].name || 'Untitled';
    const ok = confirm(`Delete project "${name}"? This cannot be undone.`);
    if (!ok) return;

    if (projects.length <= 1) {
        // Keep at least one empty project so UI remains usable
        projects = [makeEmptyProject('Project 1')];
        activeProjectIndex = 0;
    } else {
        projects.splice(index, 1);
        // Adjust active index
        if (activeProjectIndex > index) {
            activeProjectIndex -= 1;
        } else if (activeProjectIndex === index) {
            activeProjectIndex = Math.max(0, activeProjectIndex - 1);
        }
    }

    // Load active project into globals and refresh UI
    syncGlobalsFromActive();
    renderProjectTabs();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    markAsChanged();
}

// Modal-based renaming
let renameModalIndex = null;

function openRenameModal(index) {
    renameModalIndex = index;
    const backdrop = document.getElementById('renameModal');
    const input = document.getElementById('renameProjectInput');
    input.value = projects[index]?.name || '';
    backdrop.style.display = 'flex';
    setTimeout(() => { input.focus(); input.select(); }, 0);
    input.onkeydown = (e) => {
        if (e.key === 'Enter') { e.preventDefault(); confirmRenameModal(); }
        if (e.key === 'Escape') { e.preventDefault(); closeRenameModal(); }
    };
}

function closeRenameModal() {
    const backdrop = document.getElementById('renameModal');
    const input = document.getElementById('renameProjectInput');
    backdrop.style.display = 'none';
    input.onkeydown = null;
    renameModalIndex = null;
}

function confirmRenameModal() {
    const input = document.getElementById('renameProjectInput');
    const val = (input.value || '').trim();
    const idx = renameModalIndex;
    if (idx != null && val) {
        performRename(idx, val);
    }
    closeRenameModal();
}

function performRename(index, newName) {
    if (!projects[index]) return;
    const cur = projects[index].name || 'Untitled';
    if (newName && newName !== cur) {
        projects[index].name = newName;
        renderProjectTabs();
        markAsChanged();
    }
}

// Keep API for any existing callers
function renameProject(index) { openRenameModal(index); }

// Auto-save functionality
function markAsChanged() {
    hasUnsavedChanges = true;
    // keep active project in sync when any change happens
    try { syncActiveFromGlobals(); } catch (e) {}
    updateSaveIndicator();
}

function markAsSaved() {
    hasUnsavedChanges = false;
    lastSavedData = getCurrentDataHash();
    updateSaveIndicator();
}

function getCurrentDataHash() {
    // Include projects for proper change detection
    return JSON.stringify({
        projects,
        activeProjectIndex
    });
}

function updateSaveIndicator() {
    const indicator = document.getElementById('saveIndicator');
    if (hasUnsavedChanges) {
        indicator.textContent = '● Unsaved changes';
        indicator.style.color = '#e74c3c';
    } else {
        indicator.textContent = '✓ All changes saved';
        indicator.style.color = '#27ae60';
    }
}

// Auto-save every 30 seconds
function autoSave() {
    if (hasUnsavedChanges && projects.length > 0) {
        try {
            // ensure active is synced
            syncActiveFromGlobals();
            const autoSaveData = {
                projects: projects,
                activeProjectIndex: activeProjectIndex,
                autoSaveTimestamp: new Date().toISOString(),
                version: '3.0'
            };
            
            // Use localStorage for auto-save only
            localStorage.setItem('gantt_autosave', JSON.stringify(autoSaveData));
            
            // Show brief notification
            showNotification('Auto-saved locally', 'success');
        } catch (error) {
            console.error('Auto-save failed:', error);
        }
    }
}

// Check for auto-saved data on load
function checkForAutoSave() {
    const autoSaveData = localStorage.getItem('gantt_autosave');
    if (autoSaveData && projects.length === 0) {
        try {
            const data = JSON.parse(autoSaveData);
            if ((data.version === '3.0' && data.projects && data.projects.length > 0) ||
                (data.tasks && data.tasks.length > 0)) {
                const saveTime = new Date(data.autoSaveTimestamp).toLocaleString();
                if (confirm(`Found auto-saved data from ${saveTime}.\n\nWould you like to restore it?`)) {
                    loadDataFromObject(data);
                    showNotification('Auto-saved data restored', 'success');
                } else {
                    localStorage.removeItem('gantt_autosave');
                }
            }
        } catch (error) {
            console.error('Error loading auto-save:', error);
            localStorage.removeItem('gantt_autosave');
        }
    }
}

function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 8px;
        color: white;
        font-weight: 600;
        z-index: 1000;
        transform: translateX(100%);
        transition: transform 0.3s ease;
        ${type === 'success' ? 'background: linear-gradient(45deg, #27ae60, #2ecc71);' : ''}
        ${type === 'error' ? 'background: linear-gradient(45deg, #e74c3c, #c0392b);' : ''}
        ${type === 'info' ? 'background: linear-gradient(45deg, #3498db, #2980b9);' : ''}
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.transform = 'translateX(0)';
    }, 100);
    
    setTimeout(() => {
        notification.style.transform = 'translateX(100%)';
        setTimeout(() => {
            document.body.removeChild(notification);
        }, 300);
    }, 3000);
}

// Prevent accidental page leaving
window.addEventListener('beforeunload', function(e) {
    if (hasUnsavedChanges) {
        const message = 'You have unsaved changes. Do you want to leave without saving?';
        e.returnValue = message;
        return message;
    }
});

// Data management functions
function loadDataFromObject(data, options = { mode: 'replaceAll' }) {
    const mode = options.mode || 'replaceAll';
    // Helper: normalize a project object
    const normalizeProject = (p) => ({
        name: p.name || 'Untitled',
        tasks: p.tasks || [],
        milestones: p.milestones || [],
        groups: p.groups || {},
        groupStates: p.groupStates || {},
        groupOrder: p.groupOrder || Object.keys(p.groups || {})
    });

    const isV3 = data && data.version === '3.0' && Array.isArray(data.projects);
    const isV2 = data && (Array.isArray(data.tasks) || (data.groups && typeof data.groups === 'object'));
    const isProjectsArray = Array.isArray(data) && data.length > 0 && data.every(p => typeof p === 'object');

    if (mode === 'replaceAll') {
        if (isV3) {
            projects = data.projects.map(normalizeProject);
            activeProjectIndex = Math.min(Math.max(0, data.activeProjectIndex || 0), projects.length - 1);
        } else if (isV2) {
            projects = [normalizeProject(data)];
            activeProjectIndex = 0;
        } else if (isProjectsArray) {
            projects = data.map(normalizeProject);
            activeProjectIndex = 0;
        }
    } else if (mode === 'append') {
        // Append projects to existing list
        if (!Array.isArray(projects)) projects = [];
        if (isV3) {
            projects.push(...data.projects.map(normalizeProject));
        } else if (isProjectsArray) {
            projects.push(...data.map(normalizeProject));
        } else if (isV2) {
            // single project: append as a new project using data.name if present
            projects.push(normalizeProject(data));
        }
        // Keep current active project
    } else if (mode === 'intoActive') {
        // Merge/replace into the current active project only (for single-project files)
        if (!projects || projects.length === 0) {
            projects = [normalizeProject({ name: 'Project 1' })];
            activeProjectIndex = 0;
        }
        const p = projects[activeProjectIndex] || normalizeProject({});
        if (isV2) {
            projects[activeProjectIndex] = normalizeProject({
                name: data.name || p.name,
                tasks: data.tasks || [],
                groups: data.groups || {},
                groupStates: data.groupStates || {},
                groupOrder: data.groupOrder || Object.keys(data.groups || {})
            });
        } else if (isV3 && Array.isArray(data.projects) && data.projects.length === 1) {
            // If exactly one project in v3, load it into active
            projects[activeProjectIndex] = normalizeProject(data.projects[0]);
        } else {
            // Fallback: treat as append if it's multi-project
            const toAppend = isV3 ? data.projects
                           : isProjectsArray ? data
                           : [];
            if (toAppend && toAppend.length) {
                projects.push(...toAppend.map(normalizeProject));
            }
        }
    }

    // Refresh UI from active
    syncGlobalsFromActive();
    renderProjectTabs();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    markAsSaved();
}

function saveChartAsPNG() {
    const chart = document.getElementById('ganttChart');
    
    // Create a clone of the chart for PNG generation without highlighted groups
    const chartClone = chart.cloneNode(true);
    
    // Remove any highlighting or hover effects from the clone
    const groupBars = chartClone.querySelectorAll('.group-bar');
    groupBars.forEach(bar => {
        bar.style.opacity = '0.6';
        bar.style.boxShadow = '0 2px 4px rgba(0, 0, 0, 0.1)';
    });

    // Temporarily add the clone to the document (invisible)
    chartClone.style.position = 'absolute';
    chartClone.style.left = '-9999px';
    chartClone.style.top = '0';
    document.body.appendChild(chartClone);
    
    html2canvas(chartClone, { 
        scale: 2,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        width: chartClone.scrollWidth,
        height: chartClone.scrollHeight,
        scrollX: 0,
        scrollY: 0
    }).then(canvas => {
        // Resize to fit PowerPoint-friendly width while preserving aspect ratio
        const MAX_W = 1920;
        let outCanvas = canvas;
        if (canvas.width > MAX_W) {
            const ratio = MAX_W / canvas.width;
            const oc = document.createElement('canvas');
            oc.width = Math.round(canvas.width * ratio);
            oc.height = Math.round(canvas.height * ratio);
            const ctx = oc.getContext('2d');
            ctx.imageSmoothingEnabled = true;
            ctx.imageSmoothingQuality = 'high';
            ctx.drawImage(canvas, 0, 0, oc.width, oc.height);
            outCanvas = oc;
        }
        const link = document.createElement('a');
        link.download = `gantt_chart_${new Date().toISOString().split('T')[0]}.png`;
        link.href = outCanvas.toDataURL('image/png');
        link.click();
        
        // Remove the clone
        document.body.removeChild(chartClone);
    }).catch(err => {
        console.error('Error generating PNG:', err);
        alert('Error generating PNG. Please try again.');
        document.body.removeChild(chartClone);
    });
}

function saveToFile() {
    // ensure latest globals synced back
    syncActiveFromGlobals();
    const data = {
        projects: projects,
        activeProjectIndex: activeProjectIndex,
        exportDate: new Date().toISOString(),
        version: '3.0'
    };
    
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.download = `gantt_data_${new Date().toISOString().split('T')[0]}.json`;
    link.href = URL.createObjectURL(blob);
    link.click();
    
    markAsSaved();
    localStorage.removeItem('gantt_autosave'); // Clear auto-save after manual save
    showNotification('File saved successfully', 'success');
}

// Export active project's tasks as CSV
function exportCSV() {
    // Use current globals for active project
    const rows = tasks.map(t => ({
        epic: t.group || '',
        activity: t.name || '',
        start: t.startDate || '',
        end: t.endDate || '',
        status: t.status || 'Not Started'
    }));
    // New standard header order
    const header = ['epic','activity','start','end','status'];
    const esc = v => '"' + String(v).replace(/"/g, '""') + '"';
    const lines = [header.join(',')].concat(rows.map(r => header.map(h => esc(r[h])).join(',')));
    const csv = lines.join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const ts = new Date().toISOString().split('T')[0];
    const projName = (projects[activeProjectIndex]?.name || 'project').replace(/[^\w\-]+/g,'_');
    link.download = `${projName}_tasks_${ts}.csv`;
    link.href = URL.createObjectURL(blob);
    link.click();
    if (typeof showNotification === 'function') showNotification('CSV exported', 'success');
}

function loadFromFile(input) {
    const file = input.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const text = e.target.result;
            const data = JSON.parse(text);

            const isV3 = data && data.version === '3.0' && Array.isArray(data.projects);
            const isV2 = data && (Array.isArray(data.tasks) || (data.groups && typeof data.groups === 'object'));
            const isProjectsArray = Array.isArray(data) && data.length > 0 && data.every(p => typeof p === 'object');

            if (isV3 || isV2 || isProjectsArray) {
                const beforeCount = Array.isArray(projects) ? projects.length : 0;
                if (isV3 || isProjectsArray) {
                    // Multi-project inputs: append new projects
                    const addCount = isV3 ? (data.projects?.length || 0) : data.length;
                    loadDataFromObject(data, { mode: 'append' });
                    localStorage.removeItem('gantt_autosave');
                    const afterCount = Array.isArray(projects) ? projects.length : 0;
                    const added = Math.max(0, afterCount - beforeCount);
                    showNotification(`Added ${added || addCount} project${(added || addCount) !== 1 ? 's' : ''} from file`, 'success');
                } else if (isV2) {
                    // Single-project inputs: load into active project only
                    loadDataFromObject(data, { mode: 'intoActive' });
                    localStorage.removeItem('gantt_autosave');
                    showNotification('Replaced current project with loaded file', 'success');
                }
            } else {
                showNotification('Invalid file format. Please select a JSON file exported from this app.', 'error');
            }
        } catch (error) {
            console.error('Error loading file:', error);
            showNotification('Invalid JSON file. Please check the contents.', 'error');
        }
    };
    reader.readAsText(file);
    
    // Reset the input
    input.value = '';
}

function importFromExcel(input) {
    const file = input.files[0];
    if (!file) return;

    const name = file.name.toLowerCase();
    if (name.endsWith('.csv')) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const text = e.target.result;
                const rows = parseCSV(text);
                processImportedRows(rows);
            } catch (err) {
                console.error('CSV import error:', err);
                showNotification('Error importing CSV', 'error');
            } finally {
                input.value = '';
            }
        };
        reader.readAsText(file);
    } else {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheet];
                const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
                // Normalize to array of plain objects
                const rows = json.map(r => ({ ...r }));
                processImportedRows(rows);
            } catch (err) {
                console.error('Excel import error:', err);
                showNotification('Error importing Excel file', 'error');
            } finally {
                input.value = '';
            }
        };
        reader.readAsArrayBuffer(file);
    }
}

function parseCSV(text) {
    // Simple CSV parser handling commas and quotes
    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length === 0) return [];
    const headers = splitCSVLine(lines[0]).map(h => h.trim());
    // Remove BOM if present on the first header
    if (headers.length > 0) {
        headers[0] = headers[0].replace(/^\uFEFF/, '');
    }
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        const cols = splitCSVLine(lines[i]);
        const obj = {};
        headers.forEach((h, idx) => {
            obj[h] = (cols[idx] ?? '').trim();
        });
        rows.push(obj);
    }
    return rows;
}

function splitCSVLine(line) {
    const result = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
            if (inQuotes && line[i + 1] === '"') { // Escaped quote
                cur += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === ',' && !inQuotes) {
            result.push(cur);
            cur = '';
        } else {
            cur += ch;
        }
    }
    result.push(cur);
    return result;
}

function processImportedRows(rows) {
    if (!Array.isArray(rows) || rows.length === 0) {
        showNotification('No rows found to import', 'error');
        return;
    }

    // Map headers case-insensitively
    const norm = s => String(s || '').toLowerCase().replace(/\s+/g, '').replace(/_/g, '');
    // Try to find expected headers
    const sample = rows[0];
    const keys = Object.keys(sample);
    const headerMap = {};
    for (const k of keys) {
        const nk = norm(k);
        // New preferred headers
        if (!headerMap.task && (nk === 'activity')) headerMap.task = k;
        if (!headerMap.group && (nk === 'epic')) headerMap.group = k;
        if (!headerMap.start && (nk === 'start')) headerMap.start = k;
        if (!headerMap.end && (nk === 'end')) headerMap.end = k;
        if (!headerMap.status && (nk === 'status')) headerMap.status = k;
        // Backward compatible headers
        if (!headerMap.task && (nk === 'taskname' || nk === 'task' || nk === 'task_name')) headerMap.task = k;
        if (!headerMap.group && (nk === 'groupname' || nk === 'group' || nk === 'group_name')) headerMap.group = k;
        if (!headerMap.start && (nk === 'startdate' || nk === 'start_date')) headerMap.start = k;
        if (!headerMap.end && (nk === 'enddate' || nk === 'end_date' || nk === 'finishdate' || nk === 'finish')) headerMap.end = k;
    }

    if (!headerMap.task || !headerMap.group || !headerMap.start || !headerMap.end) {
        showNotification('Missing required headers. Expected: epic, activity, start, end (status optional).', 'error');
        return;
    }

    let imported = 0;
    let idBase = Date.now();
    for (const row of rows) {
        const name = String(row[headerMap.task] || '').trim(); // activity
        const groupName = String(row[headerMap.group] || '').trim() || 'Ungrouped'; // epic
        const startUS = String(row[headerMap.start] || '').trim();
        const endUS = String(row[headerMap.end] || '').trim();
        const statusRaw = headerMap.status ? String(row[headerMap.status] || '').trim() : '';

        if (!name || !startUS || !endUS) continue;

        const startISO = usDateToISO(startUS);
        const endISO = usDateToISO(endUS);
        if (!startISO || !endISO) continue;
        if (new Date(startISO) > new Date(endISO)) continue;

        // Normalize status to one of the known keys, fallback to 'Not Started'
        const knownStatuses = Object.keys(STATUS_COLORS || {});
        const status = (knownStatuses.find(s => s.toLowerCase() === statusRaw.toLowerCase()) || 'Not Started');
        const color = STATUS_COLORS[status] || groups[groupName]?.color || '#667eea';
        const task = {
            id: idBase++,
            name,
            group: groupName,
            startDate: startISO,
            endDate: endISO,
            color,
            status
        };
        tasks.push(task);

    // Auto-create/refresh group status from tasks
    applyAutoGroupStatus(groupName);

        if (!groups[groupName]) {
            groups[groupName] = { name: groupName, color };
            groupStates[groupName] = true;
            groupOrder.push(groupName);
        }
        imported++;
    }

    if (imported > 0) {
        markAsChanged();
        updateGroupsList();
        updateChart();
        updateGroupSuggestions();
        showNotification(`Imported ${imported} task${imported !== 1 ? 's' : ''}`, 'success');
    } else {
        showNotification('No valid rows to import', 'error');
    }
}

function usDateToISO(us) {
    // Accept yyyy-mm-dd (ISO) or mm/dd/yyyy, trim spaces
    const s = String(us).trim();
    // ISO format
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
        const yyyy = m[1];
        const mm = m[2];
        const dd = m[3];
        const month = parseInt(mm, 10);
        const day = parseInt(dd, 10);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) return `${yyyy}-${mm}-${dd}`;
        return null;
    }
    // US format mm/dd/yyyy or m/d/yyyy
    m = s.match(/^([0-1]?\d)\/([0-3]?\d)\/(\d{4})$/);
    if (m) {
        const mm = m[1].padStart(2, '0');
        const dd = m[2].padStart(2, '0');
        const yyyy = m[3];
        const month = parseInt(mm, 10);
        const day = parseInt(dd, 10);
        if (month < 1 || month > 12 || day < 1 || day > 31) return null;
        return `${yyyy}-${mm}-${dd}`;
    }
    return null;
}

window.onload = function() {
    setTodayDate();
    checkForAutoSave(); // Check for auto-saved data first

    // If no projects loaded from auto-save, start with a default one
    if (!Array.isArray(projects) || projects.length === 0) {
        projects = [makeEmptyProject('Project 1')];
        activeProjectIndex = 0;
    }
    syncGlobalsFromActive();
    renderProjectTabs();

    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    setupDragAndDrop();
    setupGroupAutocomplete();
    
    // Autosave will now occur only on refresh/leave via beforeunload
    
    // Initialize save indicator
    updateSaveIndicator();
};

// Silent, synchronous autosave (no notifications/UI updates)
function quickAutoSave() {
    try {
        if (!Array.isArray(projects) || projects.length === 0) return;
        // ensure latest in-memory changes are reflected in active project
        syncActiveFromGlobals();
        const autoSaveData = {
            projects: projects,
            activeProjectIndex: activeProjectIndex,
            autoSaveTimestamp: new Date().toISOString(),
            version: '3.0'
        };
        localStorage.setItem('gantt_autosave', JSON.stringify(autoSaveData));
    } catch (e) {
        // ignore errors on unload
    }
}

// Save only when the page is actually navigating away
window.addEventListener('pagehide', quickAutoSave, { capture: true });

// Task Dates Modal Logic
let currentDateModalTaskId = null;
function openTaskDateModal(taskId) {
    currentDateModalTaskId = taskId;
    const t = tasks.find(x => x.id === taskId);
    if (!t) return;
    const m = document.getElementById('dateModal');
    document.getElementById('dateStartInput').value = t.startDate;
    document.getElementById('dateEndInput').value = t.endDate;
    m.style.display = 'flex';
}
function closeDateModal() {
    const m = document.getElementById('dateModal');
    if (m) m.style.display = 'none';
    currentDateModalTaskId = null;
}
function confirmDateModal() {
    if (currentDateModalTaskId == null) { closeDateModal(); return; }
    const start = document.getElementById('dateStartInput').value;
    const end = document.getElementById('dateEndInput').value;
    if (!start || !end) { showNotification('Please select both dates', 'error'); return; }
    if (new Date(start) > new Date(end)) { showNotification('Start date cannot be after end date', 'error'); return; }
    updateTaskDates(currentDateModalTaskId, start, end);
    closeDateModal();
}
function updateTaskDates(taskId, start, end) {
    const t = tasks.find(x => x.id === taskId);
    if (!t) return;
    t.startDate = start;
    t.endDate = end;
    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification('Task dates updated', 'success');
}

// Task/Group Color Modal Logic
let currentColorModalTaskId = null;
let currentColorModalGroupName = null;
let currentColorModalMode = null; // 'task' | 'group' | 'groupTasks'
let currentColorModalTargetInputId = null; // when mode==='input'

// Color palette (Word-like) support
let recentColors = [];
let colorPaletteBound = false;

function loadRecentColors() {
    try {
        const raw = localStorage.getItem('gantt_recent_colors');
        recentColors = Array.isArray(JSON.parse(raw)) ? JSON.parse(raw) : [];
    } catch (_) { recentColors = []; }
}

function saveRecentColors() {
    try { localStorage.setItem('gantt_recent_colors', JSON.stringify(recentColors.slice(0, 10))); } catch (_) {}
}

function renderRecentColors() {
    const wrap = document.getElementById('recentColors');
    const section = document.getElementById('recentSection');
    if (!wrap || !section) return;
    wrap.innerHTML = '';
    const colors = (recentColors || []).slice(0, 10);
    section.style.display = colors.length ? '' : 'none';
    colors.forEach(hex => {
        const btn = document.createElement('button');
        btn.className = 'swatch';
        btn.setAttribute('data-color', hex);
        btn.style.setProperty('--c', hex);
        btn.type = 'button';
        btn.title = hex;
        btn.addEventListener('click', () => selectPaletteColor(hex));
        wrap.appendChild(btn);
    });
}

function normalizeHex(hex) {
    if (!hex) return '';
    let v = String(hex).trim();
    // expand #abc -> #aabbcc
    if (/^#?[0-9a-fA-F]{3}$/.test(v)) {
        v = v.replace('#', '');
        v = '#' + v.split('').map(ch => ch + ch).join('').toUpperCase();
        return v;
    }
    if (!v.startsWith('#')) v = '#' + v;
    return v.toUpperCase();
}

function clearSwatchSelection() {
    document.querySelectorAll('.swatch.selected').forEach(el => el.classList.remove('selected'));
}

function setSelectedColorInPalette(hex) {
    const input = document.getElementById('colorInput');
    const target = normalizeHex(hex || '#667EEA');
    if (input) input.value = target;
    clearSwatchSelection();
    const match = document.querySelector(`.swatch[data-color="${target}"]`);
    if (match) match.classList.add('selected');
}

function selectPaletteColor(hex) {
    const value = normalizeHex(hex);
    setSelectedColorInPalette(value);
}

function addToRecentColors(hex) {
    const val = normalizeHex(hex);
    if (!val) return;
    recentColors = [val].concat((recentColors || []).filter(c => normalizeHex(c) !== val));
    if (recentColors.length > 10) recentColors = recentColors.slice(0, 10);
    saveRecentColors();
    renderRecentColors();
}

function bindColorPaletteOnce() {
    if (colorPaletteBound) return;
    colorPaletteBound = true;
    loadRecentColors();
    renderRecentColors();
    const palette = document.getElementById('colorPalette');
    const moreBtn = document.getElementById('moreColorsBtn');
    const input = document.getElementById('colorInput');
    if (palette) {
        palette.addEventListener('click', (e) => {
            const btn = e.target.closest('.swatch');
            if (!btn) return;
            const hex = btn.getAttribute('data-color');
            if (!hex) return;
            selectPaletteColor(hex);
        });
    }
    if (moreBtn && input) {
        moreBtn.addEventListener('click', (e) => {
            e.preventDefault();
            try {
                if (typeof input.showPicker === 'function') input.showPicker();
                else input.click();
            } catch (_) { try { input.click(); } catch (_) {} }
        });
        input.addEventListener('change', () => {
            setSelectedColorInPalette(input.value);
        });
    }
}

// Ensure bindings after DOM is ready
document.addEventListener('DOMContentLoaded', bindColorPaletteOnce);

// Preview chip sync for hidden color inputs
function updateColorPreview(inputId, previewId) {
    const inputEl = document.getElementById(inputId);
    const prevEl = document.getElementById(previewId);
    if (inputEl && prevEl) {
        const hex = normalizeHex(inputEl.value || '#808080');
        prevEl.style.background = hex;
    }
}

function attachPreviewSync(inputId, previewId) {
    const inputEl = document.getElementById(inputId);
    if (!inputEl) return;
    const handler = () => updateColorPreview(inputId, previewId);
    inputEl.addEventListener('input', handler);
    inputEl.addEventListener('change', handler);
    // initialize once
    handler();
}

document.addEventListener('DOMContentLoaded', () => {
    // Bind status selects to auto-sync colors and chips
    const taskStatus = document.getElementById('taskStatus');
    if (taskStatus) {
        taskStatus.addEventListener('change', () => syncStatusSelection('taskStatus','taskStatusChip','taskColor','taskColorPreview'));
        // initialize once
        syncStatusSelection('taskStatus','taskStatusChip','taskColor','taskColorPreview');
    }
    const addTaskStatus = document.getElementById('addTaskStatus');
    if (addTaskStatus) {
        addTaskStatus.addEventListener('change', () => syncStatusSelection('addTaskStatus','addTaskStatusChip','addTaskColor','addTaskColorPreview'));
        // initialize once
        syncStatusSelection('addTaskStatus','addTaskStatusChip','addTaskColor','addTaskColorPreview');
    }
});

// Add Task Modal Logic
let addTaskModalKeyHandlerBound = false;
function openAddTaskModal(groupName) {
    const backdrop = document.getElementById('addTaskModal');
    const nameInput = document.getElementById('addTaskNameInput');
    const groupInput = document.getElementById('addTaskGroupName');
    const startInput = document.getElementById('addTaskStartDate');
    const endInput = document.getElementById('addTaskEndDate');
    const statusSelect = document.getElementById('addTaskStatus');

    if (!backdrop || !nameInput || !groupInput || !startInput || !endInput) return;

    // Prefill readonly group
    groupInput.value = groupName || '';
    // Default status and sync color
    if (statusSelect) statusSelect.value = 'Not Started';
    syncStatusSelection('addTaskStatus','addTaskStatusChip');
    groupInput.readOnly = true;

    // Defaults: today/tomorrow and group color if exists
    const today = new Date();
    const isoToday = today.toISOString().split('T')[0];
    const tmr = new Date(); tmr.setDate(tmr.getDate() + 1);
    const isoTmr = tmr.toISOString().split('T')[0];
    if (!startInput.value) startInput.value = isoToday;
    if (!endInput.value) endInput.value = isoTmr;
    // Color is derived from status now; no color input to prime

    backdrop.style.display = 'flex';
    setTimeout(() => { nameInput.focus(); nameInput.select(); }, 0);

    // Keyboard: Enter = confirm, Escape = close (scoped while modal open)
    if (!addTaskModalKeyHandlerBound) {
        addTaskModalKeyHandlerBound = true;
        backdrop.addEventListener('keydown', function(e){
            if (backdrop.style.display === 'none') return;
            if (e.key === 'Enter') { e.preventDefault(); confirmAddTaskModal(); }
            if (e.key === 'Escape') { e.preventDefault(); closeAddTaskModal(); }
        });
    }
}

function closeAddTaskModal() {
    const backdrop = document.getElementById('addTaskModal');
    const nameInput = document.getElementById('addTaskNameInput');
    const groupInput = document.getElementById('addTaskGroupName');
    const startInput = document.getElementById('addTaskStartDate');
    const endInput = document.getElementById('addTaskEndDate');
    if (!backdrop) return;
    backdrop.style.display = 'none';
    if (nameInput) { nameInput.value = ''; nameInput.classList.remove('input-error'); }
    if (groupInput) groupInput.value = '';
    if (startInput) startInput.value = '';
    if (endInput) endInput.value = '';
}

function confirmAddTaskModal() {
    const nameInput = document.getElementById('addTaskNameInput');
    const groupInput = document.getElementById('addTaskGroupName');
    const startInput = document.getElementById('addTaskStartDate');
    const endInput = document.getElementById('addTaskEndDate');
    const statusSelect = document.getElementById('addTaskStatus');

    if (!nameInput || !groupInput || !startInput || !endInput) return;

    const taskName = (nameInput.value || '').trim();
    const groupName = (groupInput.value || '').trim() || 'Ungrouped';
    const startDate = startInput.value;
    const endDate = endInput.value;
    const taskStatus = statusSelect ? statusSelect.value : 'Not Started';
    const taskColor = STATUS_COLORS[taskStatus] || '#808080';

    // Clear previous error state
    nameInput.classList.remove('input-error');

    // Validation (reuse logic like addTask)
    if (!taskName || !startDate || !endDate) {
        if (!taskName) {
            nameInput.classList.add('input-error');
            nameInput.focus();
        }
        return;
    }
    if (new Date(startDate) > new Date(endDate)) {
        showNotification('Start date cannot be after end date', 'error');
        endInput.focus();
        return;
    }

    const task = {
        id: Date.now(),
        name: taskName,
        group: groupName,
        startDate: startDate,
        endDate: endDate,
        color: taskColor,
        status: taskStatus
    };

    tasks.push(task);

    if (groupName && groupName !== 'Ungrouped') {
        previousGroupInputs.add(groupName);
    }

    if (!groups[groupName]) {
        groups[groupName] = { name: groupName, color: taskColor };
        groupStates[groupName] = true;
        groupOrder.push(groupName);
    }
    // Auto-derive group status based on tasks
    applyAutoGroupStatus(groupName);

    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    closeAddTaskModal();
    showNotification('Task added successfully', 'success');
}
function updateTaskStatus(taskId, status) {
    const t = tasks.find(x => x.id === taskId);
    if (!t) return;
    t.status = status;
    t.color = STATUS_COLORS[status] || '#808080';
    if (status === 'Completed')   t.progress = 100;
    if (status === 'Not Started') t.progress = 0;
    // Auto-sync owning group status from tasks
    if (t.group) applyAutoGroupStatus(t.group);
    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification(`Status set to ${status}`, 'success');
}
function openGroupColorModal(groupName) {
    if (!groupName || !groups[groupName]) return;
    currentColorModalMode = 'group';
    currentColorModalTaskId = null;
    currentColorModalGroupName = groupName;
    const m = document.getElementById('colorModal');
    const modalEl = m ? m.querySelector('.modal') : null;
    const input = document.getElementById('colorInput');
    const title = document.getElementById('colorTitle');
    if (title) title.textContent = 'Change Group Color';
    bindColorPaletteOnce();
    setSelectedColorInPalette(groups[groupName].color || '#808080');
    renderRecentColors();
    if (modalEl) { modalEl.style.position=''; modalEl.style.left=''; modalEl.style.top=''; modalEl.style.transform=''; }
    m.style.display = 'flex';
}

// Open color modal to change ALL tasks in a group (bulk apply)
function openGroupTasksColorModal(groupName) {
    if (!groupName || !groups[groupName]) return;
    currentColorModalMode = 'groupTasks';
    currentColorModalTaskId = null;
    currentColorModalGroupName = groupName;
    const m = document.getElementById('colorModal');
    const modalEl = m ? m.querySelector('.modal') : null;
    const input = document.getElementById('colorInput');
    const title = document.getElementById('colorTitle');
    if (title) title.textContent = 'Change All Tasks Colors';
    // Prefer group color if set; fallback to first task color in group; else default
    const groupColor = groups[groupName].color;
    const firstTask = tasks.find(t => t.group === groupName);
    bindColorPaletteOnce();
    setSelectedColorInPalette(groupColor || (firstTask ? (firstTask.color || '#808080') : '#808080'));
    renderRecentColors();
    if (modalEl) { modalEl.style.position=''; modalEl.style.left=''; modalEl.style.top=''; modalEl.style.transform=''; }
    m.style.display = 'flex';
}

// Apply a color to all tasks within a group
function applyColorToGroupTasks(groupName, color) {
    if (!groupName) return;
    let count = 0;
    tasks.forEach(t => {
        if (t.group === groupName) {
            t.color = color;
            count++;
        }
    });
    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification(`Updated ${count} task${count !== 1 ? 's' : ''} color${count !== 1 ? 's' : ''}`, 'success');
}
function closeColorModal() {
    const m = document.getElementById('colorModal');
    if (m) m.style.display = 'none';
    currentColorModalTaskId = null;
    currentColorModalGroupName = null;
    currentColorModalMode = null;
    currentColorModalTargetInputId = null;
}
function confirmColorModal() {
    const color = document.getElementById('colorInput').value;
    if (currentColorModalMode === 'group' && currentColorModalGroupName) {
        updateGroupColor(currentColorModalGroupName, color);
    } else if (currentColorModalMode === 'groupTasks' && currentColorModalGroupName) {
        applyColorToGroupTasks(currentColorModalGroupName, color);
    } else if (currentColorModalMode === 'input' && currentColorModalTargetInputId) {
        const target = document.getElementById(currentColorModalTargetInputId);
        if (target) {
            target.value = color;
            // trigger input event for any listeners
            try { target.dispatchEvent(new Event('input', { bubbles: true })); } catch (_) {}
            try { target.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {}
        }
    }
    addToRecentColors(color);
    closeColorModal();
}

// Open color modal for an arbitrary color input field using the palette
function openColorPickerForInput(inputId, titleText = 'Choose Color', anchorEl = null) {
    const inputEl = document.getElementById(inputId);
    if (!inputEl) return;
    currentColorModalMode = 'input';
    currentColorModalTargetInputId = inputId;
    const m = document.getElementById('colorModal');
    const modalEl = m ? m.querySelector('.modal') : null;
    const title = document.getElementById('colorTitle');
    if (title) title.textContent = titleText;
    bindColorPaletteOnce();
    setSelectedColorInPalette(inputEl.value || '#808080');
    renderRecentColors();
    if (m) {
        m.style.display = 'flex';
        // If anchor provided, position as popover near the anchor within viewport
        if (anchorEl && modalEl) {
            // Prepare for measurement
            const prevPosition = modalEl.style.position;
            const prevLeft = modalEl.style.left;
            const prevTop = modalEl.style.top;
            const prevTransform = modalEl.style.transform;
            modalEl.style.position = 'fixed';
            modalEl.style.left = '-9999px';
            modalEl.style.top = '-9999px';
            modalEl.style.transform = 'none';
            // Next frame: measure and place
            setTimeout(() => {
                const rect = anchorEl.getBoundingClientRect();
                const mw = modalEl.offsetWidth || 320;
                const mh = modalEl.offsetHeight || 280;
                const pad = 8;
                let left = Math.min(window.innerWidth - mw - pad, Math.max(pad, rect.left));
                let top = Math.min(window.innerHeight - mh - pad, Math.max(pad, rect.bottom + 6));
                // If there's more space above than below, flip above the button
                const spaceBelow = window.innerHeight - rect.bottom;
                const spaceAbove = rect.top;
                if (spaceAbove > spaceBelow && rect.top - mh - 6 > pad) {
                    top = Math.max(pad, rect.top - mh - 6);
                }
                modalEl.style.left = left + 'px';
                modalEl.style.top = top + 'px';
                modalEl.style.position = 'fixed';
                modalEl.style.transform = 'none';
            }, 0);
        } else if (modalEl) {
            // Reset to centered modal when no anchor provided
            modalEl.style.position = '';
            modalEl.style.left = '';
            modalEl.style.top = '';
            modalEl.style.transform = '';
        }
    }
}
function updateTaskColor(taskId, color) {
    const t = tasks.find(x => x.id === taskId);
    if (!t) return;
    t.color = color;
    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification('Task color updated', 'success');
}

// Clear red highlight as user types a task name
(function(){
    const taskNameEl = document.getElementById('taskName');
    if (taskNameEl) {
        taskNameEl.addEventListener('input', function(){
            this.classList.remove('input-error');
        });
    }
})();

// Helper to draw attention to the task name field with a brief pulse
function triggerTaskNameAttention(duration = 1500) {
    const el = document.getElementById('taskName');
    if (!el) return;
    el.classList.add('attention-pulse');
    setTimeout(() => el.classList.remove('attention-pulse'), duration);
}

function setTodayDate() {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const today = `${yyyy}-${mm}-${dd}`;
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    
    if (!startDateInput.value) {
        startDateInput.value = today;
    }
    if (!endDateInput.value) {
        const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
        const ty = tomorrow.getFullYear();
        const tm = String(tomorrow.getMonth() + 1).padStart(2, '0');
        const td = String(tomorrow.getDate()).padStart(2, '0');
        endDateInput.value = `${ty}-${tm}-${td}`;
    }
}

function addTask() {
    const taskNameInput = document.getElementById('taskName');
    const groupNameInput = document.getElementById('groupName');
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    const statusSelect = document.getElementById('taskStatus');

    const taskName = taskNameInput.value.trim();
    const groupName = groupNameInput.value.trim() || 'Ungrouped';
    const startDate = startDateInput.value;
    const endDate = endDateInput.value;
    const taskStatus = statusSelect ? statusSelect.value : 'Not Started';
    const taskColor = STATUS_COLORS[taskStatus] || '#808080';

    // remove previous error highlight
    taskNameInput.classList.remove('input-error');

    // highlight task name box on missing fields
    if (!taskName || !startDate || !endDate) {
        if (!taskName) {
            taskNameInput.classList.add('input-error');
            taskNameInput.focus();
        }
        return;
    }

    if (new Date(startDate) > new Date(endDate)) {
        showNotification('Start date cannot be after end date', 'error');
        endDateInput.focus();
        return;
    }

    const task = {
        id: Date.now(),
        name: taskName,
        group: groupName,
        startDate: startDate,
        endDate: endDate,
        color: taskColor,
        status: taskStatus
    };

    tasks.push(task);

    // Remember this group name as a previous input
    if (groupName && groupName !== 'Ungrouped') {
        previousGroupInputs.add(groupName);
    }

    if (!groups[groupName]) {
        groups[groupName] = {
            name: groupName,
            color: taskColor
        };
        groupStates[groupName] = true;
        groupOrder.push(groupName);
    }

    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    clearForm();
    showNotification('Task added successfully', 'success');
}

function deleteTask(taskId) {
    const task = tasks.find(t => t.id === taskId);
    if (!task) return;

    if (!confirm(`Are you sure you want to delete "${task.name}"?`)) {
        return;
    }

    tasks = tasks.filter(t => t.id !== taskId);

    const groupTasks = tasks.filter(t => t.group === task.group);
    if (groupTasks.length === 0) {
        delete groups[task.group];
        delete groupStates[task.group];
        groupOrder = groupOrder.filter(g => g !== task.group);
    } else {
        // Recompute remaining group's status
        applyAutoGroupStatus(task.group);
    }

    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification('Task deleted', 'success');
}

function deleteGroup(groupName) {
    if (!groups[groupName]) return;
    const count = tasks.filter(t => t.group === groupName).length;
    const message = count > 0
        ? `Delete group "${groupName}" and its ${count} task${count !== 1 ? 's' : ''}?`
        : `Delete empty group "${groupName}"?`;
    if (!confirm(message)) return;

    // Remove all tasks in the group
    tasks = tasks.filter(t => t.group !== groupName);

    // Remove group state and order
    delete groups[groupName];
    delete groupStates[groupName];
    groupOrder = groupOrder.filter(g => g !== groupName);

    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    showNotification('Group deleted', 'success');
}

function updateTaskField(taskId, field, value) {
    const task = tasks.find(t => t.id === taskId);
    if (!task) return;

    if (field === 'startDate' && (parseYMD(value) > parseYMD(task.endDate))) {
        showNotification('Start date cannot be after end date', 'error');
        return;
    }
    if (field === 'endDate' && (parseYMD(value) < parseYMD(task.startDate))) {
        showNotification('End date cannot be before start date', 'error');
        return;
    }

    const oldGroup = task.group;
    task[field] = value;

    if (field === 'group') {
        const newGroup = value.trim() || 'Ungrouped';
        task.group = newGroup;

        if (!groups[newGroup]) {
            groups[newGroup] = {
                name: newGroup,
                color: task.color
            };
            groupStates[newGroup] = true;
            groupOrder.push(newGroup);
        }

        const oldGroupTasks = tasks.filter(t => t.group === oldGroup);
        if (oldGroupTasks.length === 0) {
            delete groups[oldGroup];
            delete groupStates[oldGroup];
            groupOrder = groupOrder.filter(g => g !== oldGroup);
        }
        // Recompute statuses for both groups
        if (oldGroup && groups[oldGroup]) applyAutoGroupStatus(oldGroup);
        if (newGroup && groups[newGroup]) applyAutoGroupStatus(newGroup);
    }

    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
}

function updateGroupField(groupName, field, value) {
    if (!groups[groupName]) return;
    
    if (field === 'name') {
        const newName = value.trim();
        if (newName && newName !== groupName && !groups[newName]) {
            // Update group name
            groups[newName] = { ...groups[groupName], name: newName };
            groupStates[newName] = groupStates[groupName];
            
            // Update all tasks in this group
            tasks.forEach(task => {
                if (task.group === groupName) {
                    task.group = newName;
                }
            });
            
            // Update group order
            const index = groupOrder.indexOf(groupName);
            if (index !== -1) {
                groupOrder[index] = newName;
            }
            
            // Clean up old group
            delete groups[groupName];
            delete groupStates[groupName];
        }
    } else if (field === 'color') {
        groups[groupName].color = value;
    }
    
    markAsChanged();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
}

function toggleGroup(groupName) {
    groupStates[groupName] = !groupStates[groupName];
    markAsChanged();
    updateGroupsList();
    updateChart();
}

function toggleAllGroups() {
    const allExpanded = Object.values(groupStates).every(state => state);
    const newState = !allExpanded;
    
    Object.keys(groupStates).forEach(groupName => {
        groupStates[groupName] = newState;
    });

    markAsChanged();
    updateGroupsList();
    updateChart();
}

function clearForm() {
    const tn = document.getElementById('taskName');
    tn.value = '';
    tn.classList.remove('input-error');
    document.getElementById('groupName').value = '';
    setGroupGhost('');
    setTodayDate();
    const statusSel = document.getElementById('taskStatus');
    if (statusSel) statusSel.value = 'Not Started';
    syncStatusSelection('taskStatus','taskStatusChip');
}

function getAllGroupNames() {
    const names = new Set(Object.keys(groups));
    previousGroupInputs.forEach(n => { if (n) names.add(n); });
    return Array.from(names).sort((a,b) => a.localeCompare(b));
}

function updateGroupSuggestions(filter = '') {
    const datalist = document.getElementById('groupSuggestions');
    const all = getAllGroupNames();
    const f = String(filter || '').toLowerCase();
    const list = f ? all.filter(n => n.toLowerCase().startsWith(f)) : all;
    datalist.innerHTML = list.map(groupName => 
        `<option value="${groupName}"></option>`
    ).join('');
    // Toggle hint if there is at least one completion candidate and user typed something
    const hint = document.getElementById('groupHint');
    if (hint) {
        hint.classList.toggle('show', !!f && list.length > 0);
    }
    // Update ghost suggestion to first match
    setGroupGhost(filter);
}

function setupGroupAutocomplete() {
    const input = document.getElementById('groupName');
    const hint = document.getElementById('groupHint');
    if (!input) return;
    syncGroupGhostMetrics();

    // Populate from existing known names initially
    updateGroupSuggestions('');

    input.addEventListener('input', () => {
        updateGroupSuggestions(input.value);
    });
    // Update ghost on caret moves as well
    input.addEventListener('keyup', () => setGroupGhost(input.value));
    input.addEventListener('click', () => setGroupGhost(input.value));
    window.addEventListener('resize', syncGroupGhostMetrics);

    // On blur, remember typed value as a previous input
    input.addEventListener('blur', () => {
        const val = (input.value || '').trim();
        if (val) {
            previousGroupInputs.add(val);
            updateGroupSuggestions('');
        }
        if (hint) hint.classList.remove('show');
        // Clear ghost on blur
        setGroupGhost('');
    });

    // Tab to autocomplete to first matching suggestion
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Tab') {
            const val = input.value.trim();
            if (!val) return; // let Tab behave normally
            const all = getAllGroupNames();
            const match = all.find(n => n.toLowerCase().startsWith(val.toLowerCase()));
            if (match && match !== val) {
                e.preventDefault();
                input.value = match;
                // Keep the caret at the end after completion
                requestAnimationFrame(() => {
                    input.setSelectionRange(match.length, match.length);
                });
                // Hide hint after autocompletion
                if (hint) hint.classList.remove('show');
                updateGroupSuggestions('');
                // Clear ghost after applying completion
                setGroupGhost('');
            }
        }
    });
}

function setGroupGhost(prefix) {
    const ghost = document.getElementById('groupGhost');
    const input = document.getElementById('groupName');
    if (!ghost || !input) return;
    const val = String(typeof prefix === 'string' ? prefix : input.value || '');
    if (!val) { ghost.textContent = ''; return; }
    const all = getAllGroupNames();
    const match = all.find(n => n.toLowerCase().startsWith(val.toLowerCase()));
    // Only render if caret at end (avoid misalignment for mid-text edits)
    const atEnd = input.selectionStart === input.selectionEnd && input.selectionEnd === val.length;
    if (!match || !atEnd) { ghost.textContent = ''; return; }
    const suffix = match.slice(val.length);
    if (!suffix) { ghost.textContent = ''; return; }
    ghost.innerHTML = `<span class="ghost-prefix">${val}</span><span class="ghost-suffix">${suffix}</span>`;
}

function syncGroupGhostMetrics() {
    const input = document.getElementById('groupName');
    const ghost = document.getElementById('groupGhost');
    if (!input || !ghost) return;
    const cs = window.getComputedStyle(input);
    // Mirror key metrics & spacing
    ghost.style.fontFamily = cs.fontFamily;
    ghost.style.fontSize = cs.fontSize;
    ghost.style.fontWeight = cs.fontWeight;
    ghost.style.lineHeight = cs.lineHeight;
    ghost.style.letterSpacing = cs.letterSpacing;
    ghost.style.textTransform = cs.textTransform;
    ghost.style.padding = cs.padding;
}

function updateGroupsList() {
    const groupsList = document.getElementById('groupsList');
    if (groupsList) { groupsList.style.display = 'none'; groupsList.innerHTML = ''; }
    return; // side list removed
}

function setupDragAndDrop() {
    document.addEventListener('dragstart', handleDragStart);
    document.addEventListener('dragover', handleDragOver);
    document.addEventListener('drop', handleDrop);
    document.addEventListener('dragend', handleDragEnd);
}

function handleDragStart(e) {
    // Handle task items in the list
    if (e.target.dataset.taskId && e.target.classList.contains('task-item')) {
        draggedElement = e.target;
        draggedType = 'task';
        e.dataTransfer.setData('text/plain', e.target.dataset.taskId);
        e.target.classList.add('dragging');
    } 
    // Handle group items in the list
    else if (e.target.dataset.group && e.target.classList.contains('group-item')) {
        draggedElement = e.target;
        draggedType = 'group';
        e.dataTransfer.setData('text/plain', e.target.dataset.group);
        e.target.classList.add('dragging');
    }
    // Handle chart rows (tasks in the chart)
    else if (e.target.closest('.chart-row')) {
        const chartRow = e.target.closest('.chart-row');
        if (chartRow.dataset.taskId) {
            draggedElement = chartRow;
            draggedType = 'chart-task';
            e.dataTransfer.setData('text/plain', chartRow.dataset.taskId);
            chartRow.classList.add('dragging');
        }
    }
    // Handle chart group headers - check for group header first, then group container
    else if (e.target.closest('.chart-group-header')) {
        const chartGroup = e.target.closest('.chart-group');
        if (chartGroup && chartGroup.dataset.group) {
            draggedElement = chartGroup;
            draggedType = 'chart-group';
            e.dataTransfer.setData('text/plain', chartGroup.dataset.group);
            chartGroup.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        }
    }
    // Handle chart groups (fallback)
    else if (e.target.closest('.chart-group')) {
        const chartGroup = e.target.closest('.chart-group');
        if (chartGroup.dataset.group) {
            draggedElement = chartGroup;
            draggedType = 'chart-group';
            e.dataTransfer.setData('text/plain', chartGroup.dataset.group);
            chartGroup.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        }
    }
}

function handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    
    // Handle drop zones in the list
    const dropZone = e.target.closest('.drop-zone');
    if (dropZone) {
        dropZone.classList.add('drag-over');
        return;
    }

    // Handle chart drop zones
    const chartGroup = e.target.closest('.chart-group');
    const chartRow = e.target.closest('.chart-row');
    
    if (draggedType === 'chart-task' || draggedType === 'task') {
        // Allow dropping on chart groups or between chart rows
        if (chartGroup && !chartRow) {
            chartGroup.classList.add('drag-over');
        } else if (chartRow && chartRow !== draggedElement) {
            chartRow.classList.add('drag-over');
        }
    } else if (draggedType === 'chart-group' || draggedType === 'group') {
        // Allow dropping between groups
        if (chartGroup && chartGroup !== draggedElement) {
            chartGroup.classList.add('drag-over');
        }
    }
}

function handleDrop(e) {
    e.preventDefault();
    
    // Handle list drop zones
    const dropZone = e.target.closest('.drop-zone');
    if (dropZone) {
        dropZone.classList.remove('drag-over');
        
        if (draggedType === 'task' || draggedType === 'chart-task') {
            const taskId = parseInt(e.dataTransfer.getData('text/plain'));
            const targetGroup = dropZone.dataset.group;
            
            if (taskId && targetGroup) {
                updateTaskField(taskId, 'group', targetGroup);
            }
        }
        return;
    }

    // Handle chart drops
    const chartGroup = e.target.closest('.chart-group');
    const chartRow = e.target.closest('.chart-row');
    
    if (draggedType === 'chart-task' || draggedType === 'task') {
        if (chartGroup && !chartRow) {
            // Dropped on a chart group - move task to that group
            const taskId = parseInt(e.dataTransfer.getData('text/plain'));
            const targetGroupName = chartGroup.dataset.group;
            
            if (taskId && targetGroupName) {
                updateTaskField(taskId, 'group', targetGroupName);
            }
        } else if (chartRow && chartRow !== draggedElement) {
            // Dropped on another task - reorder within group or move to target task's group
            const draggedTaskId = parseInt(e.dataTransfer.getData('text/plain'));
            const targetTaskId = parseInt(chartRow.dataset.taskId);
            
            if (draggedTaskId && targetTaskId) {
                reorderTasks(draggedTaskId, targetTaskId);
            }
        }
    } else if (draggedType === 'chart-group' || draggedType === 'group') {
        if (chartGroup && chartGroup !== draggedElement) {
            // Reorder groups
            const draggedGroupName = e.dataTransfer.getData('text/plain');
            const targetGroupName = chartGroup.dataset.group;
            
            if (draggedGroupName && targetGroupName) {
                reorderGroups(draggedGroupName, targetGroupName);
            }
        }
    }
}

function handleDragEnd(e) {
    if (draggedElement) {
        draggedElement.classList.remove('dragging');
        draggedElement = null;
        draggedType = null;
    }
    
    // Remove all drag-over classes
    document.querySelectorAll('.drag-over').forEach(el => {
        el.classList.remove('drag-over');
    });
}

function reorderTasks(draggedTaskId, targetTaskId) {
    const draggedTask = tasks.find(t => t.id === draggedTaskId);
    const targetTask = tasks.find(t => t.id === targetTaskId);
    
    if (!draggedTask || !targetTask) return;
    
    // If tasks are in different groups, move dragged task to target's group
    if (draggedTask.group !== targetTask.group) {
        updateTaskField(draggedTaskId, 'group', targetTask.group);
    } else {
        // Reorder within the same group
        const groupTasks = tasks.filter(t => t.group === draggedTask.group);
        const draggedIndex = groupTasks.findIndex(t => t.id === draggedTaskId);
        const targetIndex = groupTasks.findIndex(t => t.id === targetTaskId);
        
        if (draggedIndex !== -1 && targetIndex !== -1) {
            // Remove the dragged task from its current position
            const taskIndex = tasks.findIndex(t => t.id === draggedTaskId);
            const [movedTask] = tasks.splice(taskIndex, 1);
            
            // Find the new position in the main tasks array
            const targetTaskIndex = tasks.findIndex(t => t.id === targetTaskId);
            tasks.splice(targetTaskIndex, 0, movedTask);
            
            markAsChanged();
            updateGroupsList();
            updateChart();
        }
    }
}

function reorderGroups(draggedGroupName, targetGroupName) {
    const draggedIndex = groupOrder.indexOf(draggedGroupName);
    const targetIndex = groupOrder.indexOf(targetGroupName);
    
    if (draggedIndex !== -1 && targetIndex !== -1) {
        groupOrder.splice(draggedIndex, 1);
        groupOrder.splice(targetIndex, 0, draggedGroupName);
        markAsChanged();
        updateGroupsList();
        updateChart();
    }
}

// Parse a date string, treating YYYY-MM-DD as a local date (to avoid UTC timezone shifts)
function parseYMD(dateString) {
    if (!dateString) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
        const [y, m, d] = dateString.split('-').map(Number);
        const dt = new Date(y, (m || 1) - 1, d || 1);
        return isNaN(dt) ? null : dt;
    }
    const d = new Date(dateString);
    return isNaN(d) ? null : d;
}

function formatDate(dateString) {
    const date = (dateString instanceof Date) ? dateString : parseYMD(dateString);
    if (!date) return '';
    return date.toLocaleDateString('en-US', { 
        month: 'short', 
        day: 'numeric',
        year: 'numeric'
    });
}

function updateChart() {
    const chartContainer = document.getElementById('ganttChart');
    
    // Determine overall chart bounds from tasks and milestones
    const taskDates = tasks
        .flatMap(task => [parseYMD(task.startDate), parseYMD(task.endDate)])
        .filter(Boolean);
    const milestoneDates = (milestones || [])
        .filter(m => m && m.date && parseYMD(m.date))
        .map(m => parseYMD(m.date))
        .filter(Boolean);
    const combined = taskDates.concat(milestoneDates);
    if (combined.length === 0) {
        chartContainer.innerHTML = '<div class="no-tasks">Add tasks or milestones to see your Gantt chart</div>';
        return;
    }

    const minDate = new Date(Math.min(...combined.map(d => +d)));
    const maxDate = new Date(Math.max(...combined.map(d => +d)));
    // Expose for interactions
    chartMinDate = new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate());
    chartMaxDate = new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate());
    
    const months = getMonthRange(minDate, maxDate);
    const quarterGroups = groupMonthsByQuarters(months);

    const quartersHTML = quarterGroups.map(q => {
        return `
            <div class="timeline-quarter" style="flex: ${q.months.length};">
                <div class="quarter-year">${q.fyLabel}</div>
                <div class="quarter-name">Q${q.quarter}</div>
            </div>
        `;
    }).join('');

    const monthsHTML = months.map(month => `
        <div class="month-cell">
            ${getMonthShortName(month.month)} ${month.year.toString().slice(-2)}
        </div>
    `).join('');

    // Build header grid lines to align exactly with body grid
    const headerGridLines = months.slice(1).map((m, idx) => {
        const pos = ((idx + 1) * 100 / months.length);
        const isQuarter = (m.month === 4 || m.month === 7 || m.month === 10 || m.month === 1);
        return `<div class="${isQuarter ? 'qline' : 'vline'}" style="left: ${pos}%"></div>`;
    }).join('');

    const legendHTML = renderStatusLegend();

    const chartHTML = `
        <div class="chart-header">
            <div class="chart-title">Sprints &amp; Activities</div>
            <div class="timeline-container">
                <div class="timeline-quarters">${quartersHTML}</div>
                <div class="timeline-months">${monthsHTML}</div>
                <div class="timeline-grid header-grid" aria-hidden="true">${headerGridLines}</div>
            </div>
        </div>
        <div class="chart-body">
            <div class="milestones-overlay">${renderMilestonesOverlay()}${renderTodayMarker()}</div>
            ${createGroupChartHTML(months)}
        </div>
        ${legendHTML}
    `;

    chartContainer.innerHTML = chartHTML;
    // Re-bind interactions after re-render
    setupTaskResizeInteractions();
    setupMilestoneInteractions();
    setupStatusContextMenu();
}

// Utilities for date math and mapping
function dateToISO(d) {
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}
function addDays(d, n) {
    const nd = new Date(d);
    nd.setDate(nd.getDate() + n);
    return nd;
}

// Context menu for setting status on right-click over task/group bars
let statusMenuEl = null;
let statusMenuContext = null; // { type: 'task'|'group', taskId?, groupName? }
let statusMenuBound = false;
const STATUS_LIST = ['Not Started','In Progress','Delayed','Blocked','Action Needed','Completed'];
function setupStatusContextMenu() {
    const chart = document.getElementById('ganttChart');
    if (!chart) return;

    // Build menu element once
    if (!statusMenuEl) {
        statusMenuEl = document.createElement('div');
        statusMenuEl.id = 'statusContextMenu';
        statusMenuEl.className = 'status-context-menu';
        statusMenuEl.style.display = 'none';
        statusMenuEl.setAttribute('role', 'menu');
        // innerHTML populated per-show to highlight current
        document.body.appendChild(statusMenuEl);

        statusMenuEl.addEventListener('click', (e) => {
            const item = e.target.closest('.item');
            if (!item || !statusMenuContext) return;
            const val = item.getAttribute('data-status');
            if (statusMenuContext.type === 'task') {
                updateTaskStatus(statusMenuContext.taskId, val);
            } else if (statusMenuContext.type === 'group') {
                updateGroupStatus(statusMenuContext.groupName, val);
            }
            hideStatusMenu();
        });
        statusMenuEl.addEventListener('keydown', (e) => {
            const items = Array.from(statusMenuEl.querySelectorAll('.item'));
            if (items.length === 0) return;
            const active = document.activeElement && document.activeElement.classList.contains('item')
                ? document.activeElement : null;
            let idx = active ? items.indexOf(active) : -1;
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                idx = (idx + 1 + items.length) % items.length;
                items[idx].focus();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                idx = (idx - 1 + items.length) % items.length;
                items[idx].focus();
            } else if (e.key === 'Home') {
                e.preventDefault(); items[0].focus();
            } else if (e.key === 'End') {
                e.preventDefault(); items[items.length - 1].focus();
            } else if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                const focusEl = document.activeElement && document.activeElement.classList.contains('item') ? document.activeElement : null;
                if (focusEl) focusEl.click();
            } else if (e.key === 'Escape') {
                e.preventDefault(); hideStatusMenu();
            }
        });
    }

    function buildMenuHTML(current) {
        return STATUS_LIST.map(s => {
            const hex = STATUS_COLORS[s] || '#808080';
            const selected = (s === current);
            return `
                <div class="item${selected ? ' selected' : ''}" role="menuitem" tabindex="0" data-status="${s}">
                    <span class="chip" style="--chip:${hex}"></span>
                    <span class="label">${s}</span>
                    <span class="check" aria-hidden="true">✓</span>
                </div>
            `;
        }).join('');
    }

    function showStatusMenu(x, y, ctx) {
        statusMenuContext = ctx;
        // Determine current status for highlighting
        let current = null;
        if (ctx.type === 'task') {
            const t = tasks.find(tt => tt.id === ctx.taskId);
            current = t ? (t.status || 'Not Started') : null;
        } else if (ctx.type === 'group') {
            const g = groups[ctx.groupName];
            current = g ? (g.status || 'Not Started') : null;
        }
        statusMenuEl.innerHTML = buildMenuHTML(current);
        statusMenuEl.style.left = Math.max(0, x) + 'px';
        statusMenuEl.style.top = Math.max(0, y) + 'px';
        statusMenuEl.style.display = 'block';
        // Ensure within viewport
        requestAnimationFrame(() => {
            const rect = statusMenuEl.getBoundingClientRect();
            let nx = rect.left, ny = rect.top;
            if (rect.right > window.innerWidth) nx = Math.max(0, window.innerWidth - rect.width - 8);
            if (rect.bottom > window.innerHeight) ny = Math.max(0, window.innerHeight - rect.height - 8);
            statusMenuEl.style.left = nx + 'px';
            statusMenuEl.style.top = ny + 'px';
            // Focus selected or first item for keyboard navigation
            const items = Array.from(statusMenuEl.querySelectorAll('.item'));
            const sel = statusMenuEl.querySelector('.item.selected') || items[0];
            if (sel) sel.focus();
        });
        // Dismiss listeners
        window.addEventListener('click', onAnyClickDismiss, { once: true });
        window.addEventListener('contextmenu', onAnyClickDismiss, { once: true });
        window.addEventListener('resize', hideStatusMenu, { once: true });
        window.addEventListener('scroll', hideStatusMenu, { once: true });
        window.addEventListener('keydown', onEscDismiss, { once: true });
    }

    function hideStatusMenu() {
        if (statusMenuEl) statusMenuEl.style.display = 'none';
        statusMenuContext = null;
    }
    function onAnyClickDismiss() { hideStatusMenu(); }
    function onEscDismiss(e) { if (e.key === 'Escape') hideStatusMenu(); }

    // Delegate right-click on bars (listen on document to ensure capture even in headers)
    if (!statusMenuBound) {
        document.addEventListener('contextmenu', (e) => {
            const taskBar = e.target.closest('.task-bar');
            const groupBar = e.target.closest('.group-bar');
            const anyGroupEl = e.target.closest('[data-group]');
            if (!taskBar && !groupBar && !anyGroupEl) return; // let default context menu elsewhere
            e.preventDefault();
            e.stopPropagation();
            if (taskBar) {
                const id = parseInt(taskBar.getAttribute('data-task-id'), 10);
                showStatusMenu(e.clientX, e.clientY, { type: 'task', taskId: id });
            } else {
                const name = (groupBar && groupBar.getAttribute('data-group')) || (anyGroupEl && anyGroupEl.getAttribute('data-group'));
                if (name) showStatusMenu(e.clientX, e.clientY, { type: 'group', groupName: name });
            }
        });
        statusMenuBound = true;
    }
}

function clampDate(d, minD, maxD) {
    if (d < minD) return new Date(minD.getFullYear(), minD.getMonth(), minD.getDate());
    if (d > maxD) return new Date(maxD.getFullYear(), maxD.getMonth(), maxD.getDate());
    return d;
}
function daysBetween(a, b) {
    const msPerDay = 24 * 60 * 60 * 1000;
    const a0 = new Date(a.getFullYear(), a.getMonth(), a.getDate());
    const b0 = new Date(b.getFullYear(), b.getMonth(), b.getDate());
    return Math.round((b0 - a0) / msPerDay);
}

function setupTaskResizeInteractions() {
    const chart = document.getElementById('ganttChart');
    if (!chart) return;
    const bars = chart.querySelectorAll('.task-bar');
    if (!bars.length) return;

    // Build the same months model used by rendering for accurate mapping
    const months = getMonthRange(chartMinDate, chartMaxDate);
    const totalMonths = Math.max(1, months.length);
    const monthWidthPct = 100 / totalMonths;

    let rafPending = false;
    let lastClientX = 0;

    const daysInMonth = (y, m) => new Date(y, m, 0).getDate();

    // Position helpers (match createTaskRow math)
    const posFromDate = (d, inclusiveEnd = false) => {
        const y = d.getFullYear();
        const m = d.getMonth() + 1;
        const i = months.findIndex(mm => mm.year === y && mm.month === m);
        const dim = daysInMonth(y, m);
        const day = d.getDate();
        const frac = (inclusiveEnd ? day : (day - 1)) / dim; // start uses day-1, end uses day
        return i + frac; // fractional months from chartMin
    };

    const dateFromPos = (pos, inclusiveEnd = false) => {
        let idx = Math.floor(pos);
        let frac = pos - idx;
        idx = Math.max(0, Math.min(totalMonths - 1, idx));
        const ym = months[idx];
        const dim = daysInMonth(ym.year, ym.month);
        let day;
        if (inclusiveEnd) {
            day = Math.round(frac * dim);
            day = Math.max(1, Math.min(dim, day));
        } else {
            day = Math.round(frac * dim) + 1;
            day = Math.max(1, Math.min(dim, day));
        }
        return new Date(ym.year, ym.month - 1, day);
    };

    const computeLeftWidthPct = (startDate, endDate) => {
        // replicate: left = startIndex*mw + startOffset; width = ((endIndex-startIndex)*mw)+endOffset-startOffset
        const sy = startDate.getFullYear();
        const sm = startDate.getMonth() + 1;
        const ey = endDate.getFullYear();
        const em = endDate.getMonth() + 1;
        const sIdx = months.findIndex(m => m.year === sy && m.month === sm);
        const eIdx = months.findIndex(m => m.year === ey && m.month === em);
        const sDim = daysInMonth(sy, sm);
        const eDim = daysInMonth(ey, em);
        const sOff = ((startDate.getDate() - 1) / sDim) * monthWidthPct;
        const eOff = ((endDate.getDate()) / eDim) * monthWidthPct; // inclusive end
        const left = (sIdx * monthWidthPct) + sOff;
        const width = ((eIdx - sIdx) * monthWidthPct) + eOff - sOff;
        return { left, width };
    };

    const performUpdate = (clientX) => {
        if (!taskDragState) return;
        const { type, side, startX, pxPerMonth, startPosMonth, endPosMonth, barEl, taskName } = taskDragState;
        const dx = clientX - startX;
        const dMonths = dx / pxPerMonth; // fractional months delta

        let newStartPos = startPosMonth;
        let newEndPos = endPosMonth;

        if (type === 'move') {
            newStartPos = startPosMonth + dMonths;
            newEndPos = endPosMonth + dMonths;
        } else if (type === 'resize') {
            if (side === 'left') newStartPos = startPosMonth + dMonths;
            if (side === 'right') newEndPos = endPosMonth + dMonths;
        }

        // Clamp positions within chart
        newStartPos = Math.max(0, Math.min(totalMonths - 1 + 0.9999, newStartPos));
        newEndPos = Math.max(0, Math.min(totalMonths - 1 + 0.9999, newEndPos));

        // Convert back to dates (inclusive end for end date)
        let newStart = dateFromPos(newStartPos, false);
        let newEnd = dateFromPos(newEndPos, true);

        // Ensure order (no inversion)
        if (newStart > newEnd) {
            if (type === 'resize' && side === 'left') newStart = new Date(newEnd.getFullYear(), newEnd.getMonth(), newEnd.getDate());
            else if (type === 'resize' && side === 'right') newEnd = new Date(newStart.getFullYear(), newStart.getMonth(), newStart.getDate());
            else { // move case, keep duration
                const dur = daysBetween(taskDragState.origStart, taskDragState.origEnd);
                newEnd = addDays(newStart, dur);
            }
        }

        // Live update bar position/width using month math
        const { left, width } = computeLeftWidthPct(newStart, newEnd);
        const leftPct = left;
        const widthPct = width;
        barEl.style.left = leftPct + '%';
        barEl.style.width = widthPct + '%';
        barEl.classList.add('resizing');
        const durDays = Math.max(0, daysBetween(newStart, newEnd)) + 1;
        const label = taskName || barEl.title.split('\n')[0].split(':')[0];
        barEl.title = `${label}: ${formatDate(newStart)} - ${formatDate(newEnd)} (${durDays} day${durDays !== 1 ? 's' : ''})`;

        taskDragState.previewStart = newStart;
        taskDragState.previewEnd = newEnd;
    };

    const onMouseMove = (e) => {
        if (!taskDragState) return;
        lastClientX = e.clientX;
        if (rafPending) return;
        rafPending = true;
        requestAnimationFrame(() => {
            rafPending = false;
            performUpdate(lastClientX);
        });
    };

    const onMouseUp = () => {
        if (!taskDragState) return;
        const { taskId, barEl, previewStart, previewEnd, origStart, origEnd } = taskDragState;
        window.removeEventListener('mousemove', onMouseMove);
        window.removeEventListener('mouseup', onMouseUp, true);
        barEl.classList.remove('resizing');
        rafPending = false;

        // Commit only if changed
        const ps = previewStart || origStart;
        const pe = previewEnd || origEnd;
        if (dateToISO(ps) !== dateToISO(origStart) || dateToISO(pe) !== dateToISO(origEnd)) {
            updateTaskDates(taskId, dateToISO(ps), dateToISO(pe));
        }

        taskDragState = null;
    };

    bars.forEach(bar => {
        // Prevent native drag interfering
        bar.setAttribute('draggable', 'false');
        bar.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const handle = e.target.closest('.resize-handle');
            const isHandle = !!handle;
            const side = isHandle ? (handle.classList.contains('left') ? 'left' : 'right') : null;
            const taskId = parseInt(bar.getAttribute('data-task-id'), 10);
            if (!taskId) return;

            const trackEl = bar.closest('.timeline-track');
            if (!trackEl) return;
            const rect = trackEl.getBoundingClientRect();
            const pxPerMonth = rect.width / totalMonths;

            const t = tasks.find(x => x.id === taskId);
            if (!t) return;
            const origStart = new Date(t.startDate);
            const origEnd = new Date(t.endDate);
            const startPosMonth = posFromDate(origStart, false);
            const endPosMonth = posFromDate(origEnd, true);

            taskDragState = {
                type: isHandle ? 'resize' : 'move',
                side,
                taskId,
                startX: e.clientX,
                chartMin: chartMinDate,
                chartMax: chartMaxDate,
                pxPerMonth,
                origStart,
                origEnd,
                startPosMonth,
                endPosMonth,
                barEl: bar,
                trackEl,
                taskName: t.name || ''
            };

            // Bind window listeners for drag lifecycle
            window.addEventListener('mousemove', onMouseMove);
            window.addEventListener('mouseup', onMouseUp, true);
        });

        // Also disable handle native drag
        bar.querySelectorAll('.resize-handle').forEach(h => h.setAttribute('draggable','false'));
    });
}

function createGroupChartHTML(months) {
    return groupOrder.filter(groupName => groups[groupName]).map(groupName => {
        const groupTasks = tasks.filter(t => t.group === groupName);
        const isExpanded = groupStates[groupName];
        
        let groupStart = null;
        let groupEnd = null;
        
        if (groupTasks.length > 0) {
            const groupDates = groupTasks
                .flatMap(task => [parseYMD(task.startDate), parseYMD(task.endDate)])
                .filter(Boolean);
            groupStart = new Date(Math.min(...groupDates));
            groupEnd = new Date(Math.max(...groupDates));
        }

        const groupBarHTML = groupStart && groupEnd ? 
            createGroupBar(groupStart, groupEnd, months, groupName) : '';

        // Grid lines for group collapsed timeline (align with task rows)
        const groupGridLines = months.slice(1).map((m, idx) => {
            const pos = ((idx + 1) * 100 / months.length);
            const isQuarter = (m.month === 4 || m.month === 7 || m.month === 10 || m.month === 1);
            return `<div class="${isQuarter ? 'qline' : 'vline'}" style="left: ${pos}%"></div>`;
        }).join('');

        const tasksHTML = isExpanded ? groupTasks.map(task => 
            createTaskRow(task, months)
        ).join('') : '';

        const safeGroupName = groupName.replace(/'/g, "\\'");
        const groupColor = groups[groupName]?.color || '#808080';
        
        return `
            <div class="chart-group" draggable="true" data-group="${safeGroupName}">
                <div class="chart-group-header ${isExpanded ? 'expanded' : 'collapsed'}" onclick="onGroupHeaderClick(event, '${safeGroupName}')">
                    <div class="chart-group-label" ondblclick="onGroupLabelDblClick(event, '${safeGroupName}')" title="Double-click to rename group">
                        <span class="chart-group-toggle ${isExpanded ? '' : 'collapsed'}">▼</span>
                        <span class="group-label-text" title="${groupName}">${groupName} (${groupTasks.length})</span>
                        <!-- Kebab menu for group actions (inside label to align with task kebab) -->
                        <div class="kebab-menu" onclick="event.stopPropagation()">
                            <div class="kebab-icon" onclick="event.stopPropagation(); toggleKebabMenu(this)">⋮</div>
                            <div class="kebab-dropdown">
                                <div class="kebab-item" onclick="openGroupEditModal('${safeGroupName}')">
                                    <span class="kebab-item-icon">✏️</span>
                                    <span>Rename Group</span>
                                </div>
                                <div class="kebab-item" onclick="openAddTaskModal('${safeGroupName}')">
                                    <span class="kebab-item-icon">➕</span>
                                    <span>Add Task</span>
                                </div>
                                <div class="kebab-item" onclick="event.stopPropagation()" style="display:flex; align-items:center; gap:8px;">
                                    <span class="kebab-item-icon">✅</span>
                                    <span style="min-width:120px;">Set Group Status</span>
                                    <select onclick="event.stopPropagation()" onchange="updateGroupStatus('${safeGroupName}', this.value); this.closest('.kebab-dropdown')?.classList.remove('show');">
                                        <option value="Not Started" ${ (groups[safeGroupName] && (groups[safeGroupName].status || 'Not Started') === 'Not Started') ? 'selected' : ''}>Not Started</option>
                                        <option value="In Progress" ${ (groups[safeGroupName] && groups[safeGroupName].status === 'In Progress') ? 'selected' : ''}>In Progress</option>
                                        <option value="Delayed" ${ (groups[safeGroupName] && groups[safeGroupName].status === 'Delayed') ? 'selected' : ''}>Delayed</option>
                                        <option value="Blocked" ${ (groups[safeGroupName] && groups[safeGroupName].status === 'Blocked') ? 'selected' : ''}>Blocked</option>
                                        <option value="Action Needed" ${ (groups[safeGroupName] && groups[safeGroupName].status === 'Action Needed') ? 'selected' : ''}>Action Needed</option>
                                        <option value="Completed" ${ (groups[safeGroupName] && groups[safeGroupName].status === 'Completed') ? 'selected' : ''}>Completed</option>
                                    </select>
                                </div>
                                <div class="kebab-item danger" onclick="deleteGroup('${safeGroupName}')">
                                    <span class="kebab-item-icon">🗑️</span>
                                    <span>Delete Group</span>
                                </div>
                            </div>
                        </div>
                    </div>
                    ${!isExpanded ? `<div class=\"chart-group-timeline\">\n                        <div class=\"timeline-grid\" aria-hidden=\"true\">${groupGridLines}</div>\n                        ${groupBarHTML}\n                    </div>` : ``}
                </div>
                <div class="chart-group-tasks ${isExpanded ? '' : 'collapsed'}">
                    ${tasksHTML}
                </div>
            </div>
        `;
    }).join('');
}

function createGroupBar(startDate, endDate, months, groupName) {
    const startMonth = {
        year: startDate.getFullYear(),
        month: startDate.getMonth() + 1
    };
    const endMonth = {
        year: endDate.getFullYear(),
        month: endDate.getMonth() + 1
    };

    const startIndex = months.findIndex(m => 
        m.year === startMonth.year && m.month === startMonth.month
    );
    const endIndex = months.findIndex(m => 
        m.year === endMonth.year && m.month === endMonth.month
    );
    
    if (startIndex === -1 || endIndex === -1) return '';
    
    const totalMonths = months.length;
    const monthWidth = 100 / totalMonths;
    
    const startMonthDate = new Date(startMonth.year, startMonth.month - 1, 1);
    const startMonthEnd = new Date(startMonth.year, startMonth.month, 0);
    const startMonthDays = startMonthEnd.getDate();
    const daysFromMonthStart = startDate.getDate() - 1;
    const startOffset = (daysFromMonthStart / startMonthDays) * monthWidth;
    
    const endMonthDate = new Date(endMonth.year, endMonth.month - 1, 1);
    const endMonthEnd = new Date(endMonth.year, endMonth.month, 0);
    const endMonthDays = endMonthEnd.getDate();
    const daysToMonthEnd = endDate.getDate();
    const endOffset = (daysToMonthEnd / endMonthDays) * monthWidth;
    
    const left = (startIndex * monthWidth) + startOffset;
    const width = ((endIndex - startIndex) * monthWidth) + endOffset - startOffset;
    
    const duration = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
    const groupColor = groups[groupName]?.color || '#667eea';

    // Compute group-level progress from tasks
    const groupTasks = tasks.filter(t => t.group === groupName);
    let groupPct = null;
    if (groupTasks.length > 0) {
        const tasksWithPct = groupTasks.filter(t => typeof t.progress === 'number');
        if (tasksWithPct.length > 0) {
            groupPct = Math.round(tasksWithPct.reduce((sum, t) => sum + t.progress, 0) / tasksWithPct.length);
        } else {
            const completedCount = groupTasks.filter(t => t.status === 'Completed').length;
            groupPct = Math.round((completedCount / groupTasks.length) * 100);
        }
    }

    const progressHTML = groupPct !== null ? `
        <div class="bar-progress-fill" style="width:${groupPct}%; background: rgba(255,255,255,0.35);"></div>
    ` : '';
    const labelHTML = groupPct !== null ? `
        <span class="bar-pct-label">${groupPct}%</span>
    ` : '';

    return `
        <div class="group-bar" data-group="${groupName}"
             style="left: ${left}%; width: ${width}%; background: linear-gradient(45deg, ${groupColor}, ${adjustBrightness(groupColor, -20)});"
             title="${groupName}: ${formatDate(startDate)} - ${formatDate(endDate)} (${duration} day${duration !== 1 ? 's' : ''})${groupPct !== null ? ' — ' + groupPct + '% complete' : ''}">
            ${progressHTML}
            ${labelHTML}
        </div>
    `;
}

function createTaskRow(task, months) {
    const taskStart = parseYMD(task.startDate) || new Date(task.startDate);
    const taskEnd = parseYMD(task.endDate) || new Date(task.endDate);
    
    const startMonth = {
        year: taskStart.getFullYear(),
        month: taskStart.getMonth() + 1
    };
    const endMonth = {
        year: taskEnd.getFullYear(),
        month: taskEnd.getMonth() + 1
    };

    const startIndex = months.findIndex(m => 
        m.year === startMonth.year && m.month === startMonth.month
    );
    const endIndex = months.findIndex(m => 
        m.year === endMonth.year && m.month === endMonth.month
    );
    
    if (startIndex === -1 || endIndex === -1) return '';
    
    const totalMonths = months.length;
    const monthWidth = 100 / totalMonths;
    
    const startMonthDate = new Date(startMonth.year, startMonth.month - 1, 1);
    const startMonthEnd = new Date(startMonth.year, startMonth.month, 0);
    const startMonthDays = startMonthEnd.getDate();
    const daysFromMonthStart = taskStart.getDate() - 1;
    const startOffset = (daysFromMonthStart / startMonthDays) * monthWidth;
    
    const endMonthDate = new Date(endMonth.year, endMonth.month - 1, 1);
    const endMonthEnd = new Date(endMonth.year, endMonth.month, 0);
    const endMonthDays = endMonthEnd.getDate();
    const daysToMonthEnd = taskEnd.getDate();
    const endOffset = (daysToMonthEnd / endMonthDays) * monthWidth;
    
    const left = (startIndex * monthWidth) + startOffset;
    const width = ((endIndex - startIndex) * monthWidth) + endOffset - startOffset;
    
    const duration = Math.ceil((taskEnd - taskStart) / (1000 * 60 * 60 * 24)) + 1;

    const gridLines = months.slice(1).map((m, idx) => {
        const pos = ((idx + 1) * 100 / months.length);
        const isQuarter = (m.month === 4 || m.month === 7 || m.month === 10 || m.month === 1);
        return `<div class="${isQuarter ? 'qline' : 'vline'}" style="left: ${pos}%"></div>`;
    }).join('');

    // ── Progress ──────────────────────────────────────────────────────────────
    // Resolve progress percentage
    let pct = typeof task.progress === 'number' ? task.progress : null;
    if (pct === null) {
        if (task.status === 'Completed')       pct = 100;
        else if (task.status === 'Not Started') pct = 0;
    }

    const progressFillHTML = pct !== null ? `
        <div class="bar-progress-fill" style="width:${pct}%;"></div>
    ` : '';

    const pctLabelHTML = pct !== null ? `
        <span class="bar-pct-label">${pct}%</span>
    ` : '';

    const subtasks = Array.isArray(task.subtasks) ? task.subtasks : [];
    const doneCount = subtasks.filter(s => s.done).length;
    const subtaskBadge = subtasks.length > 0 ? `
        <span class="subtask-badge" title="${doneCount}/${subtasks.length} subtasks done">
            ☑ ${doneCount}/${subtasks.length}
        </span>
    ` : '';

    const labelProgressHTML = pct !== null ? `
        <div class="label-progress-track" title="${pct}% complete">
            <div class="label-progress-fill" style="width:${pct}%; background:${task.color};"></div>
        </div>
    ` : '';

    const barBg = `background: linear-gradient(90deg, ${task.color} 0%, ${adjustBrightness(task.color, -25)} 100%);`;

    const tooltipExtra = pct !== null ? ` — ${pct}% complete` : '';
    const subtaskTooltip = subtasks.length > 0 ? ` | ${doneCount}/${subtasks.length} subtasks` : '';

    return `
        <div class="chart-row" draggable="true" data-task-id="${task.id}">
            <div class="task-label" style="border-left-color: ${task.color}" onclick="event.stopPropagation()">
                <div class="task-label-inner">
                    <span class="task-label-text" ondblclick="openTaskEditModal(${task.id})" title="Double-click to rename">${escapeHtml(task.name)}</span>
                    ${subtaskBadge}
                    ${labelProgressHTML}
                </div>

                <!-- Kebab menu for task actions -->
                <div class="kebab-menu" onclick="event.stopPropagation()">
                    <div class="kebab-icon" onclick="event.stopPropagation(); toggleKebabMenu(this)">⋮</div>
                    <div class="kebab-dropdown">
                        <div class="kebab-item" onclick="openTaskEditModal(${task.id})">
                            <span class="kebab-item-icon">✏️</span>
                            <span>Rename</span>
                        </div>
                        <div class="kebab-item" onclick="openTaskDateModal(${task.id}); this.closest('.kebab-dropdown')?.classList.remove('show');">
                            <span class="kebab-item-icon">📅</span>
                            <span>Edit Dates…</span>
                        </div>
                        <div class="kebab-item" onclick="openProgressModal(${task.id}); this.closest('.kebab-dropdown')?.classList.remove('show');">
                            <span class="kebab-item-icon">📊</span>
                            <span>Set Progress…</span>
                        </div>
                        <div class="kebab-item" onclick="event.stopPropagation()" style="display:flex; align-items:center; gap:8px;">
                            <span class="kebab-item-icon">✅</span>
                            <span style="min-width:110px;">Change Status…</span>
                            <select onchange="updateTaskStatus(${task.id}, this.value); this.closest('.kebab-dropdown')?.classList.remove('show');" onclick="event.stopPropagation()">
                                <option value="Not Started" ${task.status === 'Not Started' ? 'selected' : ''}>Not Started</option>
                                <option value="In Progress" ${task.status === 'In Progress' ? 'selected' : ''}>In Progress</option>
                                <option value="Delayed" ${task.status === 'Delayed' ? 'selected' : ''}>Delayed</option>
                                <option value="Blocked" ${task.status === 'Blocked' ? 'selected' : ''}>Blocked</option>
                                <option value="Action Needed" ${task.status === 'Action Needed' ? 'selected' : ''}>Action Needed</option>
                                <option value="Completed" ${task.status === 'Completed' ? 'selected' : ''}>Completed</option>
                            </select>
                        </div>
                        <div class="kebab-item danger" onclick="deleteTask(${task.id})">
                            <span class="kebab-item-icon">🗑️</span>
                            <span>Delete Task</span>
                        </div>
                    </div>
                </div>
            </div>
            <div class="timeline-track">
                <div class="timeline-grid" aria-hidden="true">${gridLines}</div>
                <div class="task-bar" data-task-id="${task.id}"
                     style="left: ${left}%; width: ${width}%; ${barBg}"
                     title="${escapeHtml(task.name)}: ${formatDate(task.startDate)} – ${formatDate(task.endDate)} (${duration} day${duration !== 1 ? 's' : ''})${tooltipExtra}${subtaskTooltip}">
                    ${progressFillHTML}
                    ${pctLabelHTML}
                    <div class="resize-handle left" aria-hidden="true"></div>
                    <div class="resize-handle right" aria-hidden="true"></div>
                </div>
            </div>
        </div>
    `;
}

// ─── Progress modal ───────────────────────────────────────────────────────────
let progressModalTaskId = null;
function openProgressModal(taskId) {
    progressModalTaskId = taskId;
    const t = tasks.find(x => x.id === taskId);
    if (!t) return;
    const modal = document.getElementById('progressModal');
    const slider = document.getElementById('progressSlider');
    const valueEl = document.getElementById('progressValue');
    const title = document.getElementById('progressModalTitle');
    if (!modal || !slider || !valueEl) return;
    if (title) title.textContent = `Progress: ${escapeHtml(t.name)}`;
    const cur = typeof t.progress === 'number' ? t.progress : (t.status === 'Completed' ? 100 : 0);
    slider.value = cur;
    valueEl.textContent = cur + '%';
    modal.style.display = 'flex';
}
function closeProgressModal() {
    const modal = document.getElementById('progressModal');
    if (modal) modal.style.display = 'none';
    progressModalTaskId = null;
}
function confirmProgressModal() {
    if (progressModalTaskId == null) { closeProgressModal(); return; }
    const slider = document.getElementById('progressSlider');
    const pct = parseInt(slider?.value ?? 0, 10);
    const t = tasks.find(x => x.id === progressModalTaskId);
    if (!t) { closeProgressModal(); return; }
    t.progress = pct;
    // Auto-sync status: progress drives status unconditionally
    if (pct >= 100)  t.status = 'Completed';
    else if (pct > 0) t.status = 'In Progress';
    else if (pct === 0 && t.status === 'Completed') t.status = 'Not Started';
    t.color = STATUS_COLORS[t.status] || t.color;
    if (t.group) applyAutoGroupStatus(t.group);
    markAsChanged();
    updateGroupsList();
    updateChart();
    closeProgressModal();
    showNotification(`Progress set to ${pct}%`, 'success');
}

function getMonthRange(startDate, endDate) {
    const months = [];
    const current = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    const end = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

    while (current <= end) {
        months.push({
            year: current.getFullYear(),
            month: current.getMonth() + 1
        });
        current.setMonth(current.getMonth() + 1);
    }

    return months;
}

function groupMonthsByQuarters(months) {
    // Fiscal year starts in April.
    // Q1: Apr-Jun, Q2: Jul-Sep, Q3: Oct-Dec, Q4: Jan-Mar
    const buckets = new Map();

    const toFY = (y) => String(y).slice(-2).padStart(2, '0');

    months.forEach(m => {
        const calYear = m.year;
        const calMonth = m.month; // 1..12
        const isAfterMarch = calMonth >= 4; // Apr..Dec
        const fyStartYear = isAfterMarch ? calYear : (calYear - 1);
        const fyEndYear = fyStartYear + 1;
        const fyLabel = `FY ${toFY(fyStartYear)}/${toFY(fyEndYear)}`;

        let quarter;
        if (calMonth >= 4 && calMonth <= 6) quarter = 1;      // Apr-Jun
        else if (calMonth >= 7 && calMonth <= 9) quarter = 2; // Jul-Sep
        else if (calMonth >= 10 && calMonth <= 12) quarter = 3; // Oct-Dec
        else quarter = 4; // Jan-Mar

        const key = `${fyStartYear}-Q${quarter}`;
        if (!buckets.has(key)) {
            buckets.set(key, {
                fyStartYear,
                fyEndYear,
                fyLabel,
                quarter,
                months: []
            });
        }
        buckets.get(key).months.push(m);
    });

    // Sort buckets chronologically by fiscal start year, then quarter order 1..4
    const result = Array.from(buckets.values()).sort((a, b) => {
        if (a.fyStartYear !== b.fyStartYear) return a.fyStartYear - b.fyStartYear;
        return a.quarter - b.quarter;
    });
    return result;
}

function getMonthShortName(monthNum) {
    const names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return names[monthNum - 1];
}

function adjustBrightness(color, amount) {
    const usePound = color[0] === '#';
    const col = usePound ? color.slice(1) : color;
    const num = parseInt(col, 16);
    let r = (num >> 16) + amount;
    let g = (num >> 8 & 0x00FF) + amount;
    let b = (num & 0x0000FF) + amount;
    r = r > 255 ? 255 : r < 0 ? 0 : r;
    g = g > 255 ? 255 : g < 0 ? 0 : g;
    b = b > 255 ? 255 : b < 0 ? 0 : b;
    return (usePound ? '#' : '') + (r << 16 | g << 8 | b).toString(16).padStart(6, '0');
}

// Export chart as PNG using html2canvas with export-specific style overrides
function saveChartAsPNG() {
    try {
        const exportRoot = document.querySelector('.container');
        const target = document.querySelector('.chart-container') || document.getElementById('ganttChart');
        if (!target) {
            showNotification('Nothing to export yet', 'error');
            return;
        }
        // Toggle export mode to neutralize unsupported CSS during rasterization
        exportRoot.classList.add('export-mode');
        // Allow UI to apply styles
        requestAnimationFrame(() => {
            html2canvas(target, {
                backgroundColor: null,
                scale: window.devicePixelRatio > 1 ? 2 : 2,
                useCORS: true,
                logging: false
            }).then(canvas => {
                // Force exact 1920x1080 output with letterboxing (contain)
                const OUT_W = 1920;
                const OUT_H = 1080;
                const oc = document.createElement('canvas');
                oc.width = OUT_W;
                oc.height = OUT_H;
                const ctx = oc.getContext('2d');
                // Fill background; fall back to white if computed is transparent
                try {
                    const bg = getComputedStyle(document.body).backgroundColor || '#ffffff';
                    ctx.fillStyle = bg || '#ffffff';
                } catch (_) {
                    ctx.fillStyle = '#ffffff';
                }
                ctx.fillRect(0, 0, OUT_W, OUT_H);
                // Scale to fit while preserving aspect ratio
                const scale = Math.min(OUT_W / canvas.width, OUT_H / canvas.height);
                const drawW = Math.max(1, Math.round(canvas.width * scale));
                const drawH = Math.max(1, Math.round(canvas.height * scale));
                const dx = Math.floor((OUT_W - drawW) / 2);
                const dy = Math.floor((OUT_H - drawH) / 2);
                ctx.imageSmoothingEnabled = true;
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(canvas, dx, dy, drawW, drawH);
                const link = document.createElement('a');
                const ts = new Date().toISOString().replace(/[:.]/g, '-');
                link.download = `gantt-chart-${ts}.png`;
                link.href = oc.toDataURL('image/png');
                link.click();
                showNotification('PNG exported', 'success');
            }).catch(err => {
                console.error('Export failed:', err);
                showNotification('PNG export failed', 'error');
            }).finally(() => {
                exportRoot.classList.remove('export-mode');
            });
        });
    } catch (e) {
        console.error(e);
        showNotification('PNG export failed', 'error');
        const exportRoot = document.querySelector('.container');
        if (exportRoot) exportRoot.classList.remove('export-mode');
    }
}

// Export a PNG with all groups collapsed, then restore the original expansion state
function saveCollapsedChartAsPNG() {
    const exportRoot = document.querySelector('.container');
    const target = document.querySelector('.chart-container') || document.getElementById('ganttChart');
    if (!target) {
        showNotification('Nothing to export yet', 'error');
        return;
    }

    // Snapshot current expansion state
    const prevStates = { ...groupStates };

    try {
        // Collapse all groups
        Object.keys(groups || {}).forEach(g => { groupStates[g] = false; });
        updateChart();

        // Enable export mode and render next frame
        exportRoot.classList.add('export-mode');
        requestAnimationFrame(() => {
            html2canvas(target, {
                backgroundColor: null,
                scale: window.devicePixelRatio > 1 ? 2 : 2,
                useCORS: true,
                logging: false
            }).then(canvas => {
                // Force exact 1920x1080 output with letterboxing (contain)
                const OUT_W = 1920;
                const OUT_H = 1080;
                const oc = document.createElement('canvas');
                oc.width = OUT_W;
                oc.height = OUT_H;
                const ctx = oc.getContext('2d');
                try {
                    const bg = getComputedStyle(document.body).backgroundColor || '#ffffff';
                    ctx.fillStyle = bg || '#ffffff';
                } catch (_) {
                    ctx.fillStyle = '#ffffff';
                }
                ctx.fillRect(0, 0, OUT_W, OUT_H);
                const scale = Math.min(OUT_W / canvas.width, OUT_H / canvas.height);
                const drawW = Math.max(1, Math.round(canvas.width * scale));
                const drawH = Math.max(1, Math.round(canvas.height * scale));
                const dx = Math.floor((OUT_W - drawW) / 2);
                const dy = Math.floor((OUT_H - drawH) / 2);
                ctx.imageSmoothingEnabled = true;
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(canvas, dx, dy, drawW, drawH);
                const link = document.createElement('a');
                const ts = new Date().toISOString().replace(/[:.]/g, '-');
                link.download = `gantt-chart-collapsed-${ts}.png`;
                link.href = oc.toDataURL('image/png');
                link.click();
                showNotification('Collapsed PNG exported', 'success');
            }).catch(err => {
                console.error('Collapsed export failed:', err);
                showNotification('Collapsed PNG export failed', 'error');
            }).finally(() => {
                // Restore UI/state
                Object.keys(prevStates).forEach(g => { groupStates[g] = prevStates[g]; });
                exportRoot.classList.remove('export-mode');
                updateChart();
            });
        });
    } catch (e) {
        console.error(e);
        showNotification('Collapsed PNG export failed', 'error');
        // Ensure state is restored even on sync errors
        Object.keys(prevStates).forEach(g => { groupStates[g] = prevStates[g]; });
        if (exportRoot) exportRoot.classList.remove('export-mode');
        updateChart();
    }
}

// Add keyboard shortcuts
document.addEventListener('keydown', function(event) {
    if (event.ctrlKey || event.metaKey) {
        if (event.key === 'Enter') {
            event.preventDefault();
            addTask();
        } else if (event.key === 's') {
            event.preventDefault();
            saveToFile();
        }
    }
    if (event.key === 'Escape') {
        clearForm();
    }
});

// Auto-focus on task name input and draw subtle attention
(function(){
    const tn = document.getElementById('taskName');
    if (tn) {
        tn.focus();
        triggerTaskNameAttention(1800);
    }
})();

// Project duplication
function duplicateProject(index) {
    if (!projects[index]) return;
    // Ensure current active project is saved
    syncActiveFromGlobals();
    const source = projects[index];
    const clone = JSON.parse(JSON.stringify(source));
    // Generate unique name
    const base = (source.name || 'Untitled') + ' (copy)';
    let name = base;
    let n = 2;
    const names = new Set(projects.map(p => p.name || 'Untitled'));
    while (names.has(name)) { name = `${base} ${n++}`; }
    clone.name = name;
    // Insert duplicate after source and activate it
    projects.splice(index + 1, 0, clone);
    activeProjectIndex = index + 1;
    // Refresh UI/state for the new active project
    syncGlobalsFromActive();
    renderProjectTabs();
    updateGroupsList();
    updateChart();
    updateGroupSuggestions();
    markAsChanged();
    if (typeof showNotification === 'function') showNotification(`Duplicated project as "${name}"`, 'success');
}

// ── PPT Mixed Export ──────────────────────────────────────────────────────────
function openPptMixedModal() {
    const list = document.getElementById('pptMixedGroupList');
    if (!list) return;
    list.innerHTML = '';
    (groupOrder.length ? groupOrder : Object.keys(groups)).forEach(g => {
        const id = `pptmix_${CSS.escape(g)}`;
        const row = document.createElement('label');
        row.style.cssText = 'display:flex;align-items:center;gap:10px;cursor:pointer;font-size:14px;padding:4px 2px;';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = true;
        cb.dataset.group = g;
        cb.id = id;
        cb.style.cssText = 'width:16px;height:16px;cursor:pointer;accent-color:#667eea;';
        const dot = document.createElement('span');
        dot.style.cssText = `display:inline-block;width:10px;height:10px;border-radius:50%;background:${groups[g]?.color || '#808080'};flex-shrink:0;`;
        const text = document.createElement('span');
        text.textContent = g;
        row.appendChild(cb);
        row.appendChild(dot);
        row.appendChild(text);
        list.appendChild(row);
    });
    document.getElementById('pptMixedModal').style.display = 'flex';
}

function setPptMixedAll(checked) {
    document.querySelectorAll('#pptMixedGroupList input[type=checkbox]').forEach(cb => cb.checked = checked);
}

function confirmPptMixedExport() {
    const expanded = new Set();
    document.querySelectorAll('#pptMixedGroupList input[type=checkbox]').forEach(cb => {
        if (cb.checked) expanded.add(cb.dataset.group);
    });
    document.getElementById('pptMixedModal').style.display = 'none';
    exportToPPT('mixed', expanded);
}

// ── PPT Groups-Only Export ────────────────────────────────────────────────────
function openPptGroupsOnlyModal() {
    const list = document.getElementById('pptGroupsOnlyGroupList');
    if (!list) return;
    list.innerHTML = '';
    (groupOrder.length ? groupOrder : Object.keys(groups)).forEach(g => {
        const id = `pptgo_${CSS.escape(g)}`;
        const row = document.createElement('label');
        row.style.cssText = 'display:flex;align-items:center;gap:10px;cursor:pointer;font-size:14px;padding:4px 2px;';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = true;
        cb.dataset.group = g;
        cb.id = id;
        cb.style.cssText = 'width:16px;height:16px;cursor:pointer;accent-color:#667eea;';
        const dot = document.createElement('span');
        dot.style.cssText = `display:inline-block;width:10px;height:10px;border-radius:50%;background:${groups[g]?.color || '#808080'};flex-shrink:0;`;
        const text = document.createElement('span');
        text.textContent = g;
        row.appendChild(cb);
        row.appendChild(dot);
        row.appendChild(text);
        list.appendChild(row);
    });
    document.getElementById('pptGroupsOnlyModal').style.display = 'flex';
}

function setPptGroupsOnlyAll(checked) {
    document.querySelectorAll('#pptGroupsOnlyGroupList input[type=checkbox]').forEach(cb => cb.checked = checked);
}

function confirmPptGroupsOnlyExport() {
    const selected = new Set();
    document.querySelectorAll('#pptGroupsOnlyGroupList input[type=checkbox]').forEach(cb => {
        if (cb.checked) selected.add(cb.dataset.group);
    });
    if (!selected.size) {
        showNotification('Select at least one group', 'error');
        return;
    }
    document.getElementById('pptGroupsOnlyModal').style.display = 'none';
    exportToPPT('groups-only', null, selected);
}
