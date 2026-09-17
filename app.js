// --- CONFIGURATIONS ---

const firebaseConfig = {
    apiKey: "AIzaSyCtFf85MUkNSsSsT6Nv8M_09Fphm2DcQOU",
    authDomain: "dailybriefing-fe7df.firebaseapp.com",
    projectId: "dailybriefing-fe7df",
    storageBucket: "dailybriefing-fe7df.firebasestorage.app", // Ensure this matches console!
    messagingSenderId: "",
    appId: ""
};
firebase.initializeApp(firebaseConfig);

// --- GLOBAL VARIABLES ---
const db = firebase.firestore();
let targetId = null;

// GLOBALS FOR EXPORT AND ANALYTICS
let currentReportData = null;

// STORE QUILL INSTANCES
let commentEditors = {};

let collapsedStatusFilters = {
    project: 'ALL',
    active: 'ALL'
};

// --- HELPER FUNCTIONS ---

function linkify(htmlContent) {
    if (!htmlContent) return "";
    let newText = htmlContent;
    const urlPattern = /(\b(https?:\/\/[-\w+&@#\/%?=~_|!:,.;&amp;]*[-\w+&@#\/%=~_|])|(\bwww\.[-\w+&@#\/%?=~_|!:,.;&amp;]*[-\w+&@#\/%=~_|]))/gi;
    const emailPattern = /(\b[\w.-]+@[\w.-]+\.\w{2,4}\b)/gi;

    newText = newText.replace(urlPattern, function(match) {
        let href = match.replace(/&amp;/g, '&');
        if (match.startsWith('www.')) { href = 'http://' + href; }
        return '<a href="' + href.replace(/"/g, '&quot;') + '" target="_blank">' + match + '</a>';
    });

    newText = newText.replace(emailPattern, '<a href="mailto:$1">$1</a>');
    return newText;
}

// STRIP HTML HELPER FOR EXCEL EXPORT
function stripHtml(html) {
   if (!html) return "";
   let tmp = document.createElement("DIV");
   tmp.innerHTML = html;
   return tmp.textContent || tmp.innerText || "";
}

// Toggle comment visibility and Init Quill
window.toggleComments = function(id) {
    const el = document.getElementById(id);
    if (el) {
        el.classList.toggle('open');
        
        // --- QUILL INIT LOGIC ---
        const uniqueId = id.replace('comments-', '');
        const editorContainerId = 'editor-container-' + uniqueId;
        
        if (el.classList.contains('open') && !commentEditors[uniqueId]) {
            if (document.getElementById(editorContainerId)) {
                const quill = new Quill('#' + editorContainerId, {
                    theme: 'snow',
                    placeholder: 'Type a question or comment...',
                    modules: {
                        toolbar: [
                            ['bold', 'italic', 'underline'],
                            [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                            ['link', 'clean']
                        ]
                    }
                });
                commentEditors[uniqueId] = quill;
            }
        }

        const nameInput = el.querySelector('.comment-input-name');
        if(nameInput && !nameInput.value) {
            nameInput.value = localStorage.getItem('commenterName') || '';
        }
    }
};

// Post a new comment
window.postComment = function(listType, itemIndex, uniqueId) {
    const container = document.getElementById('comments-' + uniqueId);
    const nameVal = container.querySelector('.comment-input-name').value.trim();
    
    let textVal = "";
    if (commentEditors[uniqueId]) {
        const editorContent = commentEditors[uniqueId].root.innerHTML;
        const textOnly = commentEditors[uniqueId].getText().trim();
        if(textOnly.length > 0) {
            textVal = editorContent;
        }
    }

    if (!nameVal || !textVal) {
        alert("Please enter both your Name and a Comment.");
        return;
    }

    localStorage.setItem('commenterName', nameVal);

    const docRef = db.collection('briefings').doc(targetId);

    docRef.get().then(doc => {
        if (!doc.exists) return;
        const data = doc.data();
        
        let listKey = '';
        let listData = [];
        let containerId = '';
        
        if (listType === 'daily') {
            listKey = 'structuredDailyTasks';
            listData = data.structuredDailyTasks || data.dailyTasks;
            containerId = 'content-tasks';
        } else if (listType === 'project') {
            listKey = 'structuredProjects';
            listData = data.structuredProjects || data.projects;
            containerId = 'content-projects';
        } else if (listType === 'active') {
            listKey = 'structuredActiveTasks';
            listData = data.structuredActiveTasks || data.activeTasks;
            containerId = 'content-active';
        }

        if (!listData[itemIndex]) return;

        const newComment = {
            author: nameVal,
            text: textVal,
            timestamp: new Date().toISOString()
        };

        if (!listData[itemIndex].publicComments) {
            listData[itemIndex].publicComments = [];
        }

        listData[itemIndex].publicComments.push(newComment);
        
        listData[itemIndex].lastUpdated = new Date().toLocaleString();

        docRef.update({
            [listKey]: listData,
            lastUpdated: new Date().toISOString()
        }).then(() => {
            delete commentEditors[uniqueId];

            renderList(containerId, listData);
            
            const newContainer = document.getElementById('comments-' + uniqueId);
            if (newContainer) {
                window.toggleComments('comments-' + uniqueId);
                newContainer.querySelector('.comment-input-name').value = nameVal;
            }
        });
    });
};

function getUniqueStatuses(items) {
    if (!Array.isArray(items)) return [];

    return [...new Set(
        items
            .filter(item => typeof item === 'object' && item !== null && item.status)
            .map(item => item.status)
    )].sort();
}

function getListTypeFromContentId(id) {
    if (id === 'content-projects') return 'project';
    if (id === 'content-active') return 'active';
    return 'daily';
}

window.setCollapsedStatusFilter = function(listType, status) {
    collapsedStatusFilters[listType] = status;

    if (!currentReportData) return;

    if (listType === 'project') {
        const projectsData = currentReportData.structuredProjects || currentReportData.projects || [];
        renderList('content-projects', projectsData);
    } else if (listType === 'active') {
        const activeData = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
        renderList('content-active', activeData);
    }
};

function getStatusCount(items, status) {
    if (!Array.isArray(items)) return 0;
    if (status === 'ALL') return items.length;

    return items.filter(item =>
        typeof item === 'object' &&
        item !== null &&
        (item.status || '') === status
    ).length;
}

function buildStatusFilterButtons(listType, items, showOnlySelected = false) {
    if (listType !== 'project' && listType !== 'active') return '';

    const statuses = getUniqueStatuses(items);
    if (!statuses.length) return '';

    const currentFilter = collapsedStatusFilters[listType] || 'ALL';
    const buttonStatuses = ['ALL', ...statuses];

    const visibleStatuses = showOnlySelected
        ? (currentFilter === 'ALL'
            ? buttonStatuses
            : ['ALL', currentFilter])
        : buttonStatuses;

    return `
        <div class="collapsed-status-filters ${showOnlySelected && currentFilter !== 'ALL' ? 'selected-only' : ''}">
            ${visibleStatuses.map(status => {
                const count = getStatusCount(items, status);
                const label = `${status} (${count})`;

                return `
                    <button
                        type="button"
                        class="status-filter-btn ${currentFilter === status ? 'active' : ''}"
                        onclick="event.stopPropagation(); setCollapsedStatusFilter('${listType}', '${status}')"
                    >${label}</button>
                `;
            }).join('')}
        </div>
    `;
}



const renderList = (id, items) => {
    const el = document.getElementById(id);
    const headerEl = document.getElementById('header-' + id.replace('content-', ''));
    
    let listType = 'daily';
    if (id === 'content-projects') listType = 'project';
    if (id === 'content-active') listType = 'active';

    const totalCount = Array.isArray(items) ? items.length : 0;

const originalItems = Array.isArray(items) ? items : [];
const isCollapsedSection = el.classList.contains('collapsed');
const selectedStatus = collapsedStatusFilters[listType] || 'ALL';

let filteredItems = originalItems;
if ((listType === 'project' || listType === 'active') && selectedStatus !== 'ALL') {
    filteredItems = originalItems.filter(item =>
        typeof item === 'object' &&
        item !== null &&
        (item.status || '') === selectedStatus
    );
}
    
if (headerEl) {
    const titleMap = {
        'header-tasks': 'Daily Tasks',
        'header-projects': 'Active Projects',
        'header-active': 'Quick Tasks'
    };

    const baseTitle = titleMap[headerEl.id] || headerEl.textContent.split('(')[0].trim();

    const displayCount =
        (listType === 'project' || listType === 'active')
            ? filteredItems.length
            : totalCount;

    headerEl.textContent = baseTitle + ` (${displayCount})`;
}


if (!originalItems.length) {
    el.innerHTML = "<em>No items.</em>";
    return;
}

let html = '';

if (listType === 'project' || listType === 'active') {
    html += buildStatusFilterButtons(listType, originalItems, !isCollapsedSection);
}

if (!filteredItems.length) {
    el.innerHTML = html + "<em>No items for this status.</em>";
    return;
}

html += '<ol style="padding-left:20px;">';

filteredItems.forEach((item) => {
    const index = originalItems.indexOf(item);
        let name, notes = '', status, priority = '', assignedTo = '', milestone = '', milestones = [], tester = '', collaborators = '', startDate = '', endDate = '', lastUpdated = '', goal = '', attachments = [], publicComments = [];
        let dailyChecks = item.dailyChecks || [];    
        
        if (typeof item === 'object' && item !== null && item.name) {
            name = item.name;
            notes = item.notes || '';
            status = item.status;
            priority = item.priority;
            assignedTo = item.assignedTo;
            milestone = item.milestone;
            milestones = Array.isArray(item.milestones) ? item.milestones : [];
            tester = item.tester;
            collaborators = item.collaborators;
            startDate = item.startDate;
            endDate = item.endDate;
            lastUpdated = item.lastUpdated;
            goal = item.goal;
            
            if (item.attachments && Array.isArray(item.attachments)) {
                attachments = item.attachments;
            } else if (item.attachment) {
                attachments = [{ name: "Resource", url: item.attachment }];
            }

            publicComments = item.publicComments || [];

        } else if (typeof item === 'string') {
            name = item;
        } else {
            return;
        }

        const safeName = linkify(name);

let safeNotes = notes.trim().startsWith('<') ? notes : linkify(notes);

// Remove ALL empty Quill-style paragraph blocks
if (safeNotes) {

    // Remove repeated empty paragraphs
    safeNotes = safeNotes.replace(/(<p><br><\/p>\s*)+/gi, '');

    // Remove empty paragraph wrappers
    safeNotes = safeNotes.replace(/(<p>\s*<\/p>\s*)+/gi, '');

    // Remove standalone <br>
    safeNotes = safeNotes.replace(/(<br\s*\/?>\s*)+/gi, '');

    // Remove &nbsp;
    safeNotes = safeNotes.replace(/(&nbsp;\s*)+/gi, '');

    // Final safety check: strip tags and confirm real content exists
    if (safeNotes.replace(/<[^>]*>/g, '').trim() === '') {
        safeNotes = '';
    }
}       

        const uniqueId = `${listType}-${index}`;

        // --- BADGES ---
        let updatedBadge = '';
        if (lastUpdated) {
            const upDate = new Date(lastUpdated);
            const today = new Date();
            const diffDays = Math.ceil(Math.abs(today - upDate) / (1000 * 60 * 60 * 24)); 

            if (upDate.toDateString() === today.toDateString()) {
                updatedBadge = '<span class="badge-updated">✨ Updated Today</span>';
            } else if (diffDays <= 7) {
                updatedBadge = '<span class="badge-recent">🔄 Updated Recently</span>';
            }
        }

        // --- DAILY CHECK LOGIC ---
        let checkHtml = '';
        if (dailyChecks.length > 0) {
            const sorted = [...dailyChecks].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
            const latest = sorted[0];
            let badgeColor = latest.status === 'Verified' ? '#28a745' : (latest.status === 'Issues Found' ? '#dc3545' : '#6c757d');
            let icon = latest.status === 'Verified' ? '✅' : (latest.status === 'Issues Found' ? '⚠️' : '⚪');
            const timeStr = new Date(latest.timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            
            checkHtml = `
                <div style="margin-top:6px; background:#fff; border:1px solid ${badgeColor}; border-left: 5px solid ${badgeColor}; padding:6px 10px; border-radius:4px; display:flex; align-items:center; gap:8px;">
                    <span style="font-size:1.2em;">${icon}</span>
                    <div style="line-height:1.2;">
                        <div style="font-weight:bold; color:${badgeColor}; font-size:0.9em; text-transform:uppercase;">${latest.status}</div>
                        <div style="font-size:0.85em; color:#555;">${latest.note ? linkify(latest.note) : 'System Operational'}</div>
                    </div>
                </div>`;
        }

        // Meta Data
        let metaHtml = '';
        let dateParts = [];
        if (startDate) dateParts.push(`Start: ${startDate}`);
        if (lastUpdated) dateParts.push(`Updated: ${lastUpdated}`);
        if (endDate) dateParts.push(`End: ${endDate}`);
        if (dateParts.length > 0) metaHtml += `<div style="font-size:0.8em; color:#777; margin-top:2px;">📅 ${dateParts.join(' | ')}</div>`;
        if (assignedTo) metaHtml += `<div style="margin-top:2px;"><span style="font-size:0.75em; background:#e8f5e9; color:#2e7d32; padding:1px 6px; border-radius:4px; border:1px solid #c8e6c9;">👤 Assigned To: ${assignedTo}</span></div>`;
        if (tester) metaHtml += `<div style="margin-top:2px;"><span style="font-size:0.75em; background:#eef; color:#336; padding:1px 6px; border-radius:4px; border:1px solid #dde;">👤 Tester: ${tester}</span></div>`;

        // Attachments
        let attachmentHtml = '';
        if (attachments.length > 0) {
            attachmentHtml = '<div style="margin-top:4px;">' + attachments.map(att => `<div class="item-attachment">🔗 <a href="${att.url}" target="_blank">${att.name}</a></div>`).join(' ') + '</div>';
        }

        // Milestones (multi-milestone table, falls back to legacy single field)
        let milestoneHtml = '';
        if (milestones.length > 0) {
            const doneCount = milestones.filter(m => m.done).length;
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const rowsHtml = milestones.map(m => {
                const isOverdue = !m.done && !m.onHold && !m.roadblock && m.estActDate && new Date(m.estActDate + 'T00:00:00') < today;
                let icon = '⚪';
                if (m.done) icon = '✅';
                else if (m.roadblock) icon = '🚧';
                else if (m.onHold) icon = '⏸️';
                else if (isOverdue) icon = '⚠️';
                else if (m.inProgress) icon = '⏳';
                const nameStyle = m.done ? 'color:#888; text-decoration:line-through;' : '';
                const estStyle = isOverdue ? 'color:#dc3545; font-weight:bold;' : 'color:#555;';
                const holdNote = m.onHold && m.holdReason ? `<div style="font-size:0.9em; color:#c0392b; font-style:italic;">On hold: ${linkify(m.holdReason)}</div>` : '';
                const blockNote = m.roadblock ? `<div style="font-size:0.9em; color:#b02a2a; font-weight:bold;">🚧 Roadblock${m.blockReason ? ': ' + linkify(m.blockReason) : ''}</div>` : '';
                return `<tr>
                    <td style="padding:3px 8px; border:1px solid #eee; ${nameStyle}">${icon} ${linkify(m.name || '')}${holdNote}${blockNote}</td>
                    <td style="padding:3px 8px; border:1px solid #eee; text-align:center; color:#555;">${m.baseDate || '—'}</td>
                    <td style="padding:3px 8px; border:1px solid #eee; text-align:center; ${estStyle}">${m.estActDate || '—'}</td>
                </tr>`;
            }).join('');
            milestoneHtml = `
                <div style="margin-top:6px;">
                    <div style="font-size:0.8em; color:#666; font-weight:bold;">🏁 Milestones (${doneCount}/${milestones.length} complete)</div>
                    <table style="border-collapse:collapse; font-size:0.8em; margin-top:2px;">
                        <thead>
                            <tr style="background:#f8f9fa; color:#666;">
                                <th style="padding:3px 8px; border:1px solid #eee; text-align:left;">Milestone</th>
                                <th style="padding:3px 8px; border:1px solid #eee;">Base Complete</th>
                                <th style="padding:3px 8px; border:1px solid #eee;">Est/Act Complete</th>
                            </tr>
                        </thead>
                        <tbody>${rowsHtml}</tbody>
                    </table>
                </div>`;
        } else if (milestone) {
            milestoneHtml = `<small style="color:#999; font-size:0.8em; display:block;">🏁 Next: ${linkify(milestone)}</small>`;
        }

        // Comments logic - Form at bottom
        const commentCount = publicComments.length;
        const commentLabel = commentCount > 0 ? `💬 View/Add Comments (${commentCount})` : `💬 Add Question/Comment`;
        let commentsListHtml = publicComments.map(c => {
            const d = c.timestamp ? new Date(c.timestamp) : null;
            const timeStr = d ? d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) : '';
            return `<div class="comment-bubble">
                        <div class="comment-header"><span class="comment-author">${linkify(c.author)}</span><span>${timeStr}</span></div>
                        <div>${c.text.trim().startsWith('<') ? c.text : linkify(c.text)}</div>
                    </div>`;
        }).join('');

        const commentsSectionHtml = `
            <div class="comments-section">
                <button class="comment-toggle" onclick="toggleComments('comments-${uniqueId}')">${commentLabel}</button>
                <div id="comments-${uniqueId}" class="comments-container">
                    ${commentsListHtml}
                    <div class="comment-form" style="margin-top: 10px;">
                        <input type="text" class="comment-input-name" placeholder="Your Name" maxlength="20">
                        <div id="editor-container-${uniqueId}"></div>
                        <button class="btn-post" onclick="postComment('${listType}', ${index}, '${uniqueId}')">Post</button>
                    </div>
                </div>
            </div>`;

        // Render Item
        if (typeof item === 'object') {
            let color = { 'On Track': 'green', 'In Progress': '#2e8b57', 'Planning': '#5b9bd5', 'Review': '#b8860b', 'Testing': '#ff9f43', 'Delayed': 'red', 'Completed': '#800080' }[status] || 'grey';
            let pColor = priority === 'High' ? '#d9534f' : (priority === 'Medium' ? '#f0ad4e' : '#5cb85c');

            const priorityBadge = priority ? `<span style="font-size:0.8em; color:${pColor}; border:1px solid ${pColor}; padding:0 4px; border-radius:4px; margin-left:5px;">${priority}</span>` : '';
            const statusBadge = status ? `<span style="font-size:0.8em; color:${color}; border:1px solid ${color}; padding:0 4px; border-radius:4px; margin-left:5px;">${status}</span>` : '';

            // This "li" now contains a hidden content section for notes and attachments
            html += `
<li style="margin-bottom:15px;">
    <div style="margin-bottom:2px; cursor:pointer;"
     onclick="this.parentElement.querySelector('.item-details').classList.toggle('collapsed')">
        <strong>${safeName}</strong>${statusBadge}${priorityBadge}${updatedBadge}
    </div>
    ${checkHtml}
    <div class="item-details collapsed" style="margin-top: 5px;">
        ${safeNotes ? `<div style="color:#666; margin-bottom:2px;">${safeNotes}</div>` : ''}
        ${goal ? `<div class="item-goal">🎯 <strong>Goal:</strong> ${linkify(goal)}</div>` : ''}
        ${attachmentHtml}
        ${milestoneHtml}
        ${metaHtml}
    </div>
    ${commentsSectionHtml}
</li>`;
        } else {
             html += `<li style="margin-bottom:5px;"><strong>${safeName}</strong></li>`;
        }
    });
    
    el.innerHTML = html + '</ol>';
};

// --- CORE RENDER FUNCTION ---
function renderReport(data) {
    document.getElementById('default-message').classList.add('hidden');
    document.getElementById('loading-overlay').classList.add('hidden');
    document.getElementById('nav-links').classList.remove('hidden');
    document.getElementById('report-body').classList.remove('hidden');
    document.getElementById('report-subtitle').textContent = "Report: " +
        (data.displayName ? data.displayName + " (" + data.reportId + ")" : data.reportId) +
        (data.department ? " · " + data.department : "");

    currentReportData = data;

    const projectsData = data.structuredProjects || data.projects;
    const activeData = data.structuredActiveTasks || data.activeTasks;
    const dailyTasksData = data.structuredDailyTasks || data.dailyTasks;

    renderList('content-projects', projectsData || []);
    renderList('content-active', activeData || []);
    renderList('content-tasks', dailyTasksData || []);
}


// --- CORE APPLICATION LOGIC (Initialization) ---
function viewReport() {
    const input = document.getElementById('userid-input');
    const val = input.value.toLowerCase().trim();
    if (val) {
        window.location.href = "?report=" + encodeURIComponent(val);
    } else {
        alert("Please enter a User ID.");
        input.focus();
    }
}

function setLastGenerated() {
    const now = new Date();
    let hour = now.getHours();
    const min = String(now.getMinutes()).padStart(2, '0');
    const sec = String(now.getSeconds()).padStart(2, '0');
    const ampm = hour >= 12 ? 'PM' : 'AM';
    hour = hour % 12;
    hour = hour ? hour : 12;
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day   = String(now.getDate()).padStart(2, '0');
    const year  = now.getFullYear();
    const tz = Intl.DateTimeFormat('en-US', { timeZoneName: 'short' }).formatToParts(now).find(p => p.type === 'timeZoneName').value;
    const formatted = `${month}/${day}/${year} ${hour}:${min}:${sec} ${ampm} ${tz}`;
    const el = document.getElementById('last-updated');
    if (el) el.textContent = `Report generated: ${formatted}`;
}

// --- ANALYTICS VIEW LOGIC ---
let statusChart = null;
let priorityChart = null;
let workloadChart = null;
let milestoneChart = null;

function toggleAnalyticsView() {
    document.getElementById('container-projects').classList.add('hidden');
    document.getElementById('container-active').classList.add('hidden');
    document.getElementById('container-tasks').classList.add('hidden');
    document.getElementById('container-analytics').classList.remove('hidden');
    renderPublicAnalytics();
}

function closeAnalytics() {
    document.getElementById('container-analytics').classList.add('hidden');
    document.getElementById('container-projects').classList.remove('hidden');
    document.getElementById('container-active').classList.remove('hidden');
    document.getElementById('container-tasks').classList.remove('hidden');
}

// Classify a milestone into a single state (same precedence as the list view icons)
function getMilestoneState(m) {
    if (m.done) return 'Completed';
    if (m.roadblock) return 'Roadblock';
    if (m.onHold) return 'On-Hold';
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (m.estActDate && new Date(m.estActDate + 'T00:00:00') < today) return 'Overdue';
    if (m.inProgress) return 'In Progress';
    return 'Not Started';
}

// Milestone states in fixed display order, with the site's status colors
const MILESTONE_STATES = [
    { key: 'Completed',   color: '#28a745' },
    { key: 'In Progress', color: '#5b9bd5' },
    { key: 'Overdue',     color: '#dc3545' },
    { key: 'On-Hold',     color: '#6c757d' },
    { key: 'Roadblock',   color: '#fd7e14' },
    { key: 'Not Started', color: '#adb5bd' }
];

// Counts of milestone states across a list of projects
function getMilestoneStats(projects) {
    const counts = {};
    MILESTONE_STATES.forEach(s => counts[s.key] = 0);
    let total = 0;
    projects.forEach(p => {
        (Array.isArray(p.milestones) ? p.milestones : []).forEach(m => {
            counts[getMilestoneState(m)]++;
            total++;
        });
    });
    return { counts, total, completed: counts['Completed'] };
}

function renderPublicAnalytics() {
    if (!currentReportData) return;

    if (typeof ChartDataLabels !== 'undefined') {
        Chart.register(ChartDataLabels);
    }

    const projects = currentReportData.structuredProjects || currentReportData.projects || [];
    const active = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
    const daily = currentReportData.structuredDailyTasks || currentReportData.dailyTasks || [];
    
    const allItems = [...projects, ...active];

    // Status order matches admin.html STATUS_ORDER
    const STATUS_ORDER_KEYS = [
        'Delayed', 'Requirement Gathering', 'Planning', 'On Track', 'In Progress', 'Development',
        'Testing', 'Review', 'Training', 'Follow-Up', 'Future',
        'On-Hold', 'Completed', 'Maintenance', 'Other'
    ];
    const STATUS_COLORS = [
        '#dc3545', // Delayed
        '#17a2b8', // Requirement Gathering
        '#5b9bd5', // Planning
        '#28a745', // On Track
        '#2e8b57', // In Progress
        '#6610f2', // Development
        '#ff9f43', // Testing
        '#b8860b', // Review
        '#20c997', // Training
        '#fd7e14', // Follow-Up
        '#adb5bd', // Future
        '#6c757d', // On-Hold
        '#800080', // Completed
        '#9b59b6', // Maintenance
        '#e2e6ea'  // Other
    ];

    const statusCounts = {};
    STATUS_ORDER_KEYS.forEach(k => statusCounts[k] = 0);
    allItems.forEach(item => {
        const s = item.status || 'Other';
        if (statusCounts.hasOwnProperty(s)) statusCounts[s]++;
        else statusCounts['Other']++;
    });

    const prioCounts = { 'High': 0, 'Medium': 0, 'Low': 0 };
    allItems.forEach(item => {
        const p = item.priority || 'Low';
        if (prioCounts.hasOwnProperty(p)) prioCounts[p]++;
    });

    // Filter to only statuses that have items (keep chart clean)
    const activeStatusKeys   = STATUS_ORDER_KEYS.filter(k => statusCounts[k] > 0);
    const activeStatusValues = activeStatusKeys.map(k => statusCounts[k]);
    const activeStatusColors = activeStatusKeys.map(k => STATUS_COLORS[STATUS_ORDER_KEYS.indexOf(k)]);

    const totalStatus = allItems.length;
    const statusHeader = document.getElementById('chart-header-status');
    if (statusHeader) statusHeader.textContent = `Project & Task Status (Total: ${totalStatus})`;

    const totalWorkload = daily.length + projects.length + active.length;
    const workloadHeader = document.getElementById('chart-header-workload');
    if (workloadHeader) workloadHeader.textContent = `Workload Distribution (Total: ${totalWorkload})`;

    if (statusChart)    statusChart.destroy();
    if (priorityChart)  priorityChart.destroy();
    if (workloadChart)  workloadChart.destroy();
    if (milestoneChart) milestoneChart.destroy();

    statusChart = new Chart(document.getElementById('chartStatus'), {
        type: 'doughnut',
        data: {
            labels: activeStatusKeys,
            datasets: [{
                data: activeStatusValues,
                backgroundColor: activeStatusColors
            }]
        },
        options: {
            plugins: {
                datalabels: {
                    color: '#ffffff',
                    font: { weight: 'bold' },
                    formatter: (value, ctx) => value > 0 ? value : ''
                },
                legend: { position: 'right' }
            }
        }
    });

    const pKeys = Object.keys(prioCounts);
    const pValues = Object.values(prioCounts);
    const pLabels = pKeys.map((key, i) => `${key} (${pValues[i]})`);

    priorityChart = new Chart(document.getElementById('chartPriority'), {
        type: 'bar',
        data: {
            labels: pLabels,
            datasets: [{
                label: 'Count',
                data: pValues,
                backgroundColor: ['#dc3545', '#ffc107', '#28a745']
            }]
        },
        options: { 
            scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } },
            plugins: {
                legend: { display: false }, 
                datalabels: {
                    display: false
                }
            }
        }
    });

    workloadChart = new Chart(document.getElementById('chartWorkload'), {
        type: 'pie',
        data: {
            labels: ['Active Projects', 'Quick Tasks', 'Daily Tasks'],
            datasets: [{
                data: [projects.length, active.length, daily.length],
                backgroundColor: ['#9b59b6', '#fd7e14', '#007bff']
            }]
        },
        options: {
            plugins: {
                datalabels: {
                    color: '#ffffff',
                    font: { weight: 'bold' },
                    formatter: (value) => value > 0 ? value : ''
                }
            }
        }
    });

    // --- PROJECT MILESTONES CHART ---
    const msStats = getMilestoneStats(projects);
    const msHeader = document.getElementById('chart-header-milestones');
    if (msHeader) {
        msHeader.textContent = msStats.total > 0
            ? `Project Milestones (${msStats.completed}/${msStats.total} complete)`
            : 'Project Milestones';
    }

    const msActiveStates = MILESTONE_STATES.filter(s => msStats.counts[s.key] > 0);
    milestoneChart = new Chart(document.getElementById('chartMilestones'), {
        type: 'doughnut',
        data: {
            labels: msActiveStates.map(s => `${s.key} (${msStats.counts[s.key]})`),
            datasets: [{
                data: msActiveStates.map(s => msStats.counts[s.key]),
                backgroundColor: msActiveStates.map(s => s.color)
            }]
        },
        options: {
            plugins: {
                datalabels: {
                    color: '#ffffff',
                    font: { weight: 'bold' },
                    formatter: (value) => value > 0 ? value : ''
                },
                legend: { position: 'right' }
            }
        }
    });

    // --- MILESTONE PROGRESS BY PROJECT TABLE ---
    const msTableEl = document.getElementById('analytics-milestone-table');
    if (msTableEl) {
        const projectsWithMs = projects.filter(p => Array.isArray(p.milestones) && p.milestones.length > 0);
        if (projectsWithMs.length === 0) {
            msTableEl.innerHTML = '';
            msTableEl.style.display = 'none';
        } else {
            msTableEl.style.display = 'block';
            let mHtml = `
                <table style="width:100%; border-collapse:collapse; font-size:0.92em;">
                    <thead>
                        <tr style="background:#f8f9fa; border-bottom:2px solid #dee2e6;">
                            <th style="padding:10px 14px; text-align:left; color:#555; font-weight:700;" colspan="5">🏁 Milestone Progress by Project</th>
                        </tr>
                        <tr style="background:#f8f9fa; border-bottom:2px solid #dee2e6;">
                            <th style="padding:8px 14px; text-align:left; color:#555; font-weight:700;">Project</th>
                            <th style="padding:8px 14px; text-align:left; color:#555; font-weight:700;">Progress</th>
                            <th style="padding:8px 14px; text-align:left; color:#555; font-weight:700;">Next Milestone</th>
                            <th style="padding:8px 14px; text-align:center; color:#555; font-weight:700;">Est/Act Date</th>
                            <th style="padding:8px 14px; text-align:center; color:#555; font-weight:700;">Flags</th>
                        </tr>
                    </thead>
                    <tbody>`;

            projectsWithMs.forEach((p, i) => {
                const list = p.milestones;
                const doneCount = list.filter(m => m.done).length;
                const pct = Math.round((doneCount / list.length) * 100);
                const next = list.find(m => !m.done);
                const nextName = next ? (next.name || '—') : 'All complete 🎉';
                const nextDate = next ? (next.estActDate || next.baseDate || '—') : '—';

                const flags = [];
                list.forEach(m => {
                    const state = getMilestoneState(m);
                    if (state === 'Overdue') flags.push('⚠️ Overdue');
                    if (state === 'Roadblock') flags.push('🚧 Roadblock');
                    if (state === 'On-Hold') flags.push('⏸️ On-Hold');
                });
                const flagStr = [...new Set(flags)].join(' ');
                const hasProblem = flagStr.includes('Overdue') || flagStr.includes('Roadblock');

                const rowBg = i % 2 === 0 ? '#ffffff' : '#f9fafb';
                const barColor = pct === 100 ? '#28a745' : (hasProblem ? '#dc3545' : '#0078D4');
                mHtml += `
                        <tr style="background:${rowBg}; border-bottom:1px solid #e9ecef;">
                            <td style="padding:9px 14px;"><strong>${p.name || ''}</strong></td>
                            <td style="padding:9px 14px; min-width:140px;">
                                <div style="display:flex; align-items:center; gap:8px;">
                                    <div style="flex:1; background:#e9ecef; border-radius:4px; height:10px; overflow:hidden; min-width:60px;">
                                        <div style="width:${pct}%; background:${barColor}; height:100%;"></div>
                                    </div>
                                    <span style="white-space:nowrap; color:#555;">${doneCount}/${list.length}</span>
                                </div>
                            </td>
                            <td style="padding:9px 14px;">${nextName}</td>
                            <td style="padding:9px 14px; text-align:center; color:#555;">${nextDate}</td>
                            <td style="padding:9px 14px; text-align:center; ${hasProblem ? 'color:#dc3545; font-weight:bold;' : 'color:#555;'}">${flagStr || '—'}</td>
                        </tr>`;
            });

            mHtml += `
                    </tbody>
                </table>`;
            msTableEl.innerHTML = mHtml;
        }
    }

    // --- STATUS SUMMARY TABLE ---
    const tableEl = document.getElementById('analytics-status-table');
    if (tableEl) {
        const rowsWithData = STATUS_ORDER_KEYS.filter(k => k !== 'Other' && statusCounts[k] > 0);
        if (rowsWithData.length === 0) {
            tableEl.innerHTML = '';
        } else {
            const headerBg = '#f8f9fa';
            let tHtml = `
                <table style="width:100%; border-collapse:collapse; font-size:0.92em;">
                    <thead>
                        <tr style="background:${headerBg}; border-bottom:2px solid #dee2e6;">
                            <th style="padding:10px 14px; text-align:left; color:#555; font-weight:700;">Status</th>
                            <th style="padding:10px 14px; text-align:center; color:#555; font-weight:700;">Projects</th>
                            <th style="padding:10px 14px; text-align:center; color:#555; font-weight:700;">Quick Tasks</th>
                            <th style="padding:10px 14px; text-align:center; color:#555; font-weight:700;">Total</th>
                        </tr>
                    </thead>
                    <tbody>`;

            rowsWithData.forEach((status, i) => {
                const pCount = projects.filter(item => (item.status || '') === status).length;
                const aCount = active.filter(item   => (item.status || '') === status).length;
                const total  = pCount + aCount;
                const color  = STATUS_COLORS[STATUS_ORDER_KEYS.indexOf(status)];
                const rowBg  = i % 2 === 0 ? '#ffffff' : '#f9fafb';
                tHtml += `
                        <tr style="background:${rowBg}; border-bottom:1px solid #e9ecef;">
                            <td style="padding:9px 14px;">
                                <span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:${color}; margin-right:7px; vertical-align:middle;"></span>
                                <strong>${status}</strong>
                            </td>
                            <td style="padding:9px 14px; text-align:center;">${pCount > 0 ? pCount : '<span style="color:#ccc;">—</span>'}</td>
                            <td style="padding:9px 14px; text-align:center;">${aCount > 0 ? aCount : '<span style="color:#ccc;">—</span>'}</td>
                            <td style="padding:9px 14px; text-align:center; font-weight:bold;">${total}</td>
                        </tr>`;
            });

            // Totals row
            const totalProjects = projects.length;
            const totalActive   = active.length;
            tHtml += `
                        <tr style="background:#f0f4f8; border-top:2px solid #dee2e6; font-weight:700;">
                            <td style="padding:10px 14px;">Total</td>
                            <td style="padding:10px 14px; text-align:center;">${totalProjects}</td>
                            <td style="padding:10px 14px; text-align:center;">${totalActive}</td>
                            <td style="padding:10px 14px; text-align:center;">${totalProjects + totalActive}</td>
                        </tr>
                    </tbody>
                </table>`;
            tableEl.innerHTML = tHtml;
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('userid-input').addEventListener("keypress", function(event) {
        if (event.key === "Enter") viewReport();
    });

    const urlParams = new URLSearchParams(window.location.search);
    // ?daily= is the retired Daily Briefing URL — old bookmarks fall through to the report view.
    const reportId = urlParams.get('report') || urlParams.get('daily');
    const isJsonView = urlParams.get('json') === 'true';

    targetId = reportId ? reportId.toLowerCase() : null;

    if (targetId) {
        db.collection('briefings').doc(targetId).get().then(doc => {
            if (doc.exists) {
                const data = doc.data();
                if (isJsonView) {
                    document.body.innerHTML = `<pre style="background:#1e1e1e; color:#d4d4d4; padding:20px; font-family:monospace; line-height:1.5; overflow:auto;">${JSON.stringify(data, null, 2)}</pre>`;
                    document.title = `JSON Report - ${targetId}`;
                    return; // Stop further rendering
                }
                renderReport(data);
                setLastGenerated();
            } else {
                document.getElementById('loading-overlay').classList.add('hidden');
                document.getElementById('default-message').classList.remove('hidden');
                document.getElementById('report-subtitle').innerText = "No briefing found for \"" + targetId + "\". New user? Sign in to the Dashboard once to set it up.";
            }
        }).catch(error => {
            console.error("Error getting document:", error);
            document.getElementById('loading-overlay').classList.add('hidden');
            document.getElementById('default-message').classList.remove('hidden');
        });
    } else {
        document.getElementById('report-subtitle').textContent = "Please enter ID below";
        document.getElementById('loading-overlay').classList.add('hidden');
        document.getElementById('default-message').classList.remove('hidden');
    }
});

// --- EXPORT FUNCTION (Requires xlsx-js-style library) ---
// --- UPDATED EXPORT FUNCTION ---
function exportReportToExcel() {
    if (!currentReportData) {
        alert("No data loaded to export.");
        return;
    }

    // --- HELPER 1: FORMAT DATA ---
    const formatForExcel = (list) => {
        if (!Array.isArray(list)) return [];
        return list.map(item => {
            let commentsStr = "";
            if (item.publicComments && item.publicComments.length > 0) {
                commentsStr = [...item.publicComments].reverse().map(c => {
                    let timeStr = "N/A";
                    if (c.timestamp) {
                        const d = new Date(c.timestamp);
                        timeStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
                    }
                    return `[${timeStr}]: [${c.author}]: ${stripHtml(c.text)}`;
                }).join("\r\n");
            }

            let checkStatus = "";
            let checkNote = "";
            let checkHistoryStr = "";
            if (item.dailyChecks && item.dailyChecks.length > 0) {
                const sortedChecks = [...item.dailyChecks].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
                const latest = sortedChecks[0];
                checkStatus = latest.status;
                checkNote = latest.note || "";
                checkHistoryStr = sortedChecks.map(c => {
                     let d = "N/A";
                     if (c.timestamp) {
                         const dateObj = new Date(c.timestamp);
                         d = dateObj.toLocaleDateString() + ' ' + dateObj.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
                     }
                     const n = c.note ? ` - ${stripHtml(c.note)}` : "";
                     return `[${d}] [${c.status}]${n}`;
                }).join("\r\n");
            }
            
            let attachmentStr = "";
            const atts = item.attachments || (item.attachment ? [{name:'Attachment', url:item.attachment}] : []);
            if (atts.length > 0) {
                attachmentStr = atts.map(a => `[${a.name}] ${a.url}`).join("\r\n");
            }

            let milestonesStr = "";
            if (item.milestones && item.milestones.length > 0) {
                milestonesStr = item.milestones.map(m => {
                    let state = "Not Started";
                    if (m.done) state = "Done";
                    else if (m.roadblock) state = "Roadblock";
                    else if (m.onHold) state = "On-Hold";
                    else if (m.inProgress) state = "In Progress";
                    let dates = [];
                    if (m.baseDate) dates.push(`Base: ${m.baseDate}`);
                    if (m.estActDate) dates.push(`Est/Act: ${m.estActDate}`);
                    if (m.onHold && m.holdReason) dates.push(`Reason: ${m.holdReason}`);
                    if (m.roadblock && m.blockReason) dates.push(`Reason: ${m.blockReason}`);
                    return `[${state}] ${m.name}${dates.length ? ' — ' + dates.join(' | ') : ''}`;
                }).join("\r\n");
            }

            let row = {
                Name: item.name,
                Status: item.status || "",
                Daily_Check_Status: checkStatus,
                Daily_Check_Note: stripHtml(checkNote),
                Priority: item.priority || "",
                Assigned_To: item.assignedTo || "N/A",
                Goal: item.goal || "",
                Milestone: item.milestone || "",
                Milestones: milestonesStr,
                Start: item.startDate || "",
                End: item.endDate || "",
                Updated: item.lastUpdated || "",
                Tester: item.tester || "",
                Collaborators: item.collaborators || "",
                Notes: stripHtml(item.notes || ""),
                Public_Comments: commentsStr,
                Daily_Check_History: checkHistoryStr
            };
            row.Attachment = attachmentStr;
            return row;
        });
    };

    // --- HELPER 2: GLOBAL STYLING (TOP ALIGNMENT) ---
// --- HELPER 2: GLOBAL STYLING (TOP ALIGNMENT & GRIDLINES) ---
const applyGlobalStyles = (ws) => {
    if (!ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const COLUMN_WIDTH_CHARS = 45;
    const DEFAULT_ROW_HEIGHT = 20; 
    const LINE_HEIGHT_PTS = 15;

    // Status Color Map (RGB Hex for Excel) — matches admin.html STATUS_ORDER
    const statusColors = {
        'Future':                'F3F4F6',
        'Planning':              'DBEAFE',
        'In Progress':           'C6EFCE',
        'Review':                'FFF2CC',
        'Requirement Gathering': 'DBEAFE',
        'On Track':              'C6EFCE',
        'Development':           'E0E7FF',
        'Testing':               'FFEB9C',
        'Training':              'D1FAE5',
        'Completed':             'E9D5FF',
        'Follow-Up':             'FFEDD5',
        'Maintenance':           'F5F3FF',
        'Delayed':               'FFC7CE',
        'On-Hold':               'E5E7EB'
    };

    if (!ws['!cols']) ws['!cols'] = [];
    if (!ws['!rows']) ws['!rows'] = [];

    let statusColIdx = -1;
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const header = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
        if (header && header.v === "Status") { statusColIdx = C; break; }
    }

    for (let R = range.s.r; R <= range.e.r; ++R) {
        let rowColor = 'FFFFFF'; 
        if (R > 0 && statusColIdx !== -1) {
            const statusCell = ws[XLSX.utils.encode_cell({ r: R, c: statusColIdx })];
            if (statusCell && statusColors[statusCell.v]) {
                rowColor = statusColors[statusCell.v];
            }
        }

        for (let C = range.s.c; C <= range.e.c; ++C) {
            const address = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = ws[address];
            
            // If cell doesn't exist, create an empty one to ensure gridlines show
            if (!cell) {
                ws[address] = { v: "", t: "s", s: {} };
            }
            if (!ws[address].s) ws[address].s = {};
            const s = ws[address].s;

            // --- ADDED: GRIDLINES (BORDERS) ---
            s.border = {
                top: { style: "thin", color: { rgb: "D1D5DB" } },
                bottom: { style: "thin", color: { rgb: "D1D5DB" } },
                left: { style: "thin", color: { rgb: "D1D5DB" } },
                right: { style: "thin", color: { rgb: "D1D5DB" } }
            };

            // Basic Alignment
            s.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

            // Apply Status Fill
            s.fill = { patternType: "solid", fgColor: { rgb: rowColor } };

            // Header Row Formatting
            if (R === 0) {
                s.font = { bold: true };
                s.fill = { patternType: "solid", fgColor: { rgb: "F2F2F2" } };
            }

            // High Priority Override: Issues Found
            const headerCell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
            const headerVal = headerCell ? headerCell.v : "";
            if (headerVal === "Daily_Check_Status" && ws[address].v === "Issues Found") {
                s.font = { color: { rgb: "9C0006" }, bold: true };
                s.fill = { patternType: "solid", fgColor: { rgb: "FFC7CE" } };
            }

            // Column Widths
            if (["Notes", "Public_Comments", "Daily_Check_History", "Attachment", "Milestones"].includes(headerVal)) {
                ws['!cols'][C] = { wch: COLUMN_WIDTH_CHARS };
            }

            // Row Height Calculation
            if (R > 0) {
                const cellValue = ws[address].v ? String(ws[address].v) : "";
                const lineCount = (cellValue.match(/\r\n/g) || []).length + 1;
                const targetLines = Math.min(lineCount, 5); 
                const calculatedHeight = Math.max(DEFAULT_ROW_HEIGHT, targetLines * LINE_HEIGHT_PTS);

                if (!ws['!rows'][R]) ws['!rows'][R] = { hpt: DEFAULT_ROW_HEIGHT };
                if (calculatedHeight > ws['!rows'][R].hpt) {
                    ws['!rows'][R].hpt = calculatedHeight;
                }
            }
        }
    }
};

    // --- MAIN EXPORT LOGIC ---
    const wb = XLSX.utils.book_new();

    // 1. Analytics Overview (with full status breakdown)
    const analyticsData = (function() {
        const projects = currentReportData.structuredProjects || currentReportData.projects || [];
        const active   = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
        const daily    = currentReportData.structuredDailyTasks || currentReportData.dailyTasks || [];

        const statusOrder = [
            'Delayed', 'Requirement Gathering', 'Planning', 'On Track', 'In Progress', 'Development',
            'Testing', 'Review', 'Training', 'Follow-Up', 'Future',
            'On-Hold', 'Completed', 'Maintenance'
        ];

        const rows = [
            // Workload summary
            { Category: 'WORKLOAD SUMMARY', Metric: 'Active Projects', Count: projects.length, Projects: '',    Quick_Tasks: '' },
            { Category: 'WORKLOAD SUMMARY', Metric: 'Quick Tasks',    Count: active.length,   Projects: '',    Quick_Tasks: '' },
            { Category: 'WORKLOAD SUMMARY', Metric: 'Daily Tasks',    Count: daily.length,    Projects: '',    Quick_Tasks: '' },
            { Category: '',                 Metric: '',               Count: '',              Projects: '',    Quick_Tasks: '' },
            // Status breakdown header
            { Category: 'STATUS BREAKDOWN', Metric: 'Status',         Count: 'Combined Total', Projects: 'Projects', Quick_Tasks: 'Quick Tasks' }
        ];

        statusOrder.forEach(status => {
            const pCount = projects.filter(i => (i.status || '') === status).length;
            const aCount = active.filter(i   => (i.status || '') === status).length;
            if (pCount + aCount > 0) {
                rows.push({ Category: '', Metric: status, Count: pCount + aCount, Projects: pCount, Quick_Tasks: aCount });
            }
        });

        // Project milestone summary (matches the Analytics view's milestone chart)
        const msStats = getMilestoneStats(projects);
        if (msStats.total > 0) {
            rows.push({ Category: '', Metric: '', Count: '', Projects: '', Quick_Tasks: '' });
            rows.push({ Category: 'PROJECT MILESTONES', Metric: 'Total Milestones', Count: msStats.total, Projects: '', Quick_Tasks: '' });
            MILESTONE_STATES.forEach(s => {
                if (msStats.counts[s.key] > 0) {
                    rows.push({ Category: '', Metric: s.key, Count: msStats.counts[s.key], Projects: '', Quick_Tasks: '' });
                }
            });
        }

        rows.push({ Category: '', Metric: '', Count: '', Projects: '', Quick_Tasks: '' });
        rows.push({ Category: 'REPORT INFO', Metric: 'User',     Count: targetId,                    Projects: '', Quick_Tasks: '' });
        rows.push({ Category: 'REPORT INFO', Metric: 'Exported', Count: new Date().toLocaleString(), Projects: '', Quick_Tasks: '' });

        return rows;
    })();

    const wsAnalytics = XLSX.utils.json_to_sheet(analyticsData);
    applyGlobalStyles(wsAnalytics);
    // Widen key columns for readability
    wsAnalytics['!cols'] = [
        { wch: 20 }, // Category
        { wch: 28 }, // Metric
        { wch: 16 }, // Count
        { wch: 14 }, // Projects
        { wch: 14 }  // Quick_Tasks
    ];
    XLSX.utils.book_append_sheet(wb, wsAnalytics, 'Analytics Overview');

    // 2. Data Sheets (same order as the report: Projects, Quick Tasks, Daily Tasks)
    const sheetsToProcess = [
        { name: "Projects", data: currentReportData.structuredProjects || currentReportData.projects },
        { name: "Quick Tasks", data: currentReportData.structuredActiveTasks || currentReportData.activeTasks },
        { name: "Daily Tasks", data: currentReportData.structuredDailyTasks || currentReportData.dailyTasks }
    ];

    sheetsToProcess.forEach(sheetObj => {
        const formattedData = formatForExcel(sheetObj.data);
        if (formattedData.length > 0) {
            const ws = XLSX.utils.json_to_sheet(formattedData);
            applyGlobalStyles(ws);
            XLSX.utils.book_append_sheet(wb, ws, sheetObj.name);
        }
    });

    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Briefing_${targetId}_${dateStr}.xlsx`);
}



