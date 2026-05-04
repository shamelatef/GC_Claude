/**
 * Visibility / Show-Hide panel
 * Controls which UI elements are visible in the Gantt chart.
 * Settings are persisted to localStorage.
 */

const VISIBILITY_KEY = 'gantt_visibility_v1';

const VISIBILITY_DEFAULTS = {
    showProgressBar:    true,
    showProgressLabel:  true,
    showMilestones:     true,
    showTodayMarker:    true,
    showGroupBar:       true,
    showStatusColor:    true,
    showTaskDates:      false,
    showAssignee:       false,
    showSubtasks:       true,
    showGridLines:      true,
    showLegend:         true,
};

let visibilitySettings = { ...VISIBILITY_DEFAULTS };

function loadVisibilitySettings() {
    try {
        const raw = localStorage.getItem(VISIBILITY_KEY);
        if (raw) {
            const saved = JSON.parse(raw);
            visibilitySettings = { ...VISIBILITY_DEFAULTS, ...saved };
        }
    } catch (_) {}
    applyVisibilityCSS();
}

function saveVisibilitySettings() {
    try {
        localStorage.setItem(VISIBILITY_KEY, JSON.stringify(visibilitySettings));
    } catch (_) {}
}

function setVisibility(key, value) {
    visibilitySettings[key] = value;
    saveVisibilitySettings();
    applyVisibilityCSS();
    // Re-render chart to pick up new settings
    if (typeof updateChart === 'function') updateChart();
}

function getVisibility(key) {
    return visibilitySettings[key] !== undefined ? visibilitySettings[key] : VISIBILITY_DEFAULTS[key];
}

/**
 * Inject a <style> block that shows/hides CSS-controlled elements.
 * JS-rendered elements (progress bars, dates inside bars) are controlled
 * directly in createTaskRow / createGroupBar.
 */
function applyVisibilityCSS() {
    let el = document.getElementById('visibilityOverrideStyle');
    if (!el) {
        el = document.createElement('style');
        el.id = 'visibilityOverrideStyle';
        document.head.appendChild(el);
    }

    const rules = [];

    if (!visibilitySettings.showMilestones) {
        rules.push('.milestones-overlay .milestone-marker { display: none !important; }');
    }
    if (!visibilitySettings.showTodayMarker) {
        rules.push('.today-marker { display: none !important; }');
    }
    if (!visibilitySettings.showGridLines) {
        rules.push('.timeline-grid .vline, .timeline-grid .qline { display: none !important; }');
    }
    if (!visibilitySettings.showGroupBar) {
        rules.push('.group-bar { display: none !important; }');
    }
    if (!visibilitySettings.showLegend) {
        rules.push('#statusLegend { display: none !important; }');
    }

    el.textContent = rules.join('\n');
}

// ─── Panel rendering ────────────────────────────────────────────────────────

function renderVisibilityPanel() {
    const existing = document.getElementById('visibilityPanel');
    if (existing) { existing.remove(); return; }

    const panel = document.createElement('div');
    panel.id = 'visibilityPanel';
    panel.className = 'visibility-panel';
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-label', 'Show / Hide options');

    const items = [
        { key: 'showProgressBar',   label: 'Progress Bars',        icon: '📊' },
        { key: 'showProgressLabel', label: 'Progress % Label',     icon: '🔢' },
        { key: 'showSubtasks',      label: 'Subtask Count',        icon: '☑️' },
        { key: 'showMilestones',    label: 'Milestones',           icon: '🔷' },
        { key: 'showTodayMarker',   label: 'Today Marker',         icon: '📍' },
        { key: 'showGroupBar',      label: 'Group Summary Bar',    icon: '▬' },
        { key: 'showGridLines',     label: 'Grid Lines',           icon: '⌗' },
        { key: 'showStatusColor',   label: 'Status Color Bars',    icon: '🎨' },
        { key: 'showLegend',        label: 'Status Legend',        icon: '🏷️' },
    ];

    panel.innerHTML = `
        <div class="vp-header">
            <span class="vp-title">Show / Hide</span>
            <button class="vp-close" onclick="document.getElementById('visibilityPanel').remove()" aria-label="Close">✕</button>
        </div>
        <div class="vp-body">
            ${items.map(({ key, label, icon }) => `
                <label class="vp-row">
                    <span class="vp-icon">${icon}</span>
                    <span class="vp-label">${label}</span>
                    <div class="vp-toggle ${visibilitySettings[key] ? 'on' : 'off'}" data-key="${key}" onclick="toggleVisibilityItem('${key}', this)" role="switch" aria-checked="${visibilitySettings[key]}" tabindex="0">
                        <div class="vp-knob"></div>
                    </div>
                </label>
            `).join('')}
        </div>
        <div class="vp-footer">
            <button class="vp-reset" onclick="resetVisibility()">Reset to defaults</button>
        </div>
    `;

    // Close on outside click
    setTimeout(() => {
        document.addEventListener('click', function outsideClose(e) {
            const p = document.getElementById('visibilityPanel');
            const btn = document.getElementById('visibilityToggleBtn');
            if (p && !p.contains(e.target) && e.target !== btn && !btn?.contains(e.target)) {
                p.remove();
                document.removeEventListener('click', outsideClose);
            }
        });
    }, 0);

    document.querySelector('.controls')?.appendChild(panel) || document.body.appendChild(panel);
}

function toggleVisibilityItem(key, el) {
    const newVal = !visibilitySettings[key];
    setVisibility(key, newVal);
    el.classList.toggle('on', newVal);
    el.classList.toggle('off', !newVal);
    el.setAttribute('aria-checked', String(newVal));
}

function resetVisibility() {
    visibilitySettings = { ...VISIBILITY_DEFAULTS };
    saveVisibilitySettings();
    applyVisibilityCSS();
    if (typeof updateChart === 'function') updateChart();
    // Rebuild panel
    const p = document.getElementById('visibilityPanel');
    if (p) p.remove();
    renderVisibilityPanel();
}

function toggleVisibilityPanel() {
    renderVisibilityPanel();
}

// Init on load
document.addEventListener('DOMContentLoaded', loadVisibilitySettings);
