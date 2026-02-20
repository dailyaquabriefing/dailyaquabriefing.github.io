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
let pendingData = null;
let targetId = null;
let currentShowPrivate = false;

// NEW GLOBALS FOR EXPORT AND ANALYTICS
let currentReportData = null;
let currentOutlookData = null;

// STORE QUILL INSTANCES
let commentEditors = {};

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
            
            renderList(containerId, listData, currentShowPrivate);
            
            const newContainer = document.getElementById('comments-' + uniqueId);
            if (newContainer) {
                window.toggleComments('comments-' + uniqueId);
                newContainer.querySelector('.comment-input-name').value = nameVal;
            }
        });
    });
};

const renderList = (id, items, showPrivate = false) => {
    const el = document.getElementById(id);
    const headerEl = document.getElementById('header-' + id.replace('content-', ''));
    
    let listType = 'daily';
    if (id === 'content-projects') listType = 'project';
    if (id === 'content-active') listType = 'active';

    const totalCount = Array.isArray(items) ? items.length : 0;
    
    if (headerEl) {
        const titleMap = {
            'header-tasks': 'Daily Tasks',
            'header-projects': 'Active Projects',
            'header-active': 'Active Tasks'
        };
        const baseTitle = titleMap[headerEl.id] || headerEl.textContent.split('(')[0].trim();
        headerEl.textContent = baseTitle + ` (${totalCount})`;
    }

    if (!Array.isArray(items) || !items.length) {
        el.innerHTML = "<em>No items.</em>";
        return;
    }
    
    let html = '<ol style="padding-left:20px;">';
    
    items.forEach((item, index) => {
    let name, notes = '', status, priority = '', milestone = '', tester = '', collaborators = '', startDate = '', endDate = '', lastUpdated = '', goal = '', attachments = [], itComments = '', publicComments = [];
    let dailyChecks = item.dailyChecks || [];    
        if (typeof item === 'object' && item !== null && item.name) {
            name = item.name;
            notes = item.notes || '';
            status = item.status;
            priority = item.priority; 
            milestone = item.milestone;
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

            itComments = item.itComments;
            publicComments = item.publicComments || [];

        } else if (typeof item === 'string') {
            name = item;
        } else {
            return;
        }

        const safeName = linkify(name);
        
        let safeNotes = notes;
        if (!notes.trim().startsWith('<')) {
             safeNotes = linkify(notes);
        }

        const uniqueId = `${listType}-${index}`;

        // --- UPDATED BADGES ---
        let updatedBadge = '';
        if (lastUpdated) {
            const upDate = new Date(lastUpdated);
            const today = new Date();
            const upDateString = upDate.toDateString();
            const todayString = today.toDateString();
            const diffTime = Math.abs(today - upDate);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 

            if (upDateString === todayString) {
                updatedBadge = '<span class="badge-updated">✨ Updated Today</span>';
            } else if (diffDays <= 7) {
                updatedBadge = '<span class="badge-recent">🔄 Updated Recently</span>';
            }
        }

        // --- DAILY CHECK LOGIC ---
        let checkHtml = '';
        if (dailyChecks.length > 0) {
            dailyChecks.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
            const latest = dailyChecks[0];
            const checkDate = new Date(latest.timestamp).toDateString();
            const todayDate = new Date().toDateString();
            const isToday = checkDate === todayDate;

            let badgeColor = '#6c757d'; 
            let icon = '⚪';
            
            if (latest.status === 'Verified') {
                badgeColor = '#28a745'; 
                icon = '✅';
            } else if (latest.status === 'Issues Found') {
                badgeColor = '#dc3545'; 
                icon = '⚠️';
            }

            const timeStr = new Date(latest.timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            const dateDisplay = isToday ? `Today at ${timeStr}` : new Date(latest.timestamp).toLocaleDateString();

            checkHtml = `
                <div style="margin-top:6px; background:#fff; border:1px solid ${badgeColor}; border-left: 5px solid ${badgeColor}; padding:6px 10px; border-radius:4px; display:flex; align-items:center; gap:8px;">
                    <span style="font-size:1.2em;">${icon}</span>
                    <div style="line-height:1.2;">
                        <div style="font-weight:bold; color:${badgeColor}; font-size:0.9em; text-transform:uppercase;">${latest.status}</div>
                        <div style="font-size:0.85em; color:#555;">${latest.note ? linkify(latest.note) : 'System Operational'}</div>
                        <div style="font-size:0.75em; color:#999;">Checked: ${dateDisplay}</div>
                    </div>
                </div>
            `;
        }

        // Meta Data HTML
        let metaHtml = '';
        let dateParts = [];
        const fmtDate = (d) => d;

        if (startDate) dateParts.push(`Start: ${fmtDate(startDate)}`);
        if (lastUpdated) dateParts.push(`Updated: ${fmtDate(lastUpdated)}`);
        if (endDate) dateParts.push(`End: ${fmtDate(endDate)}`);
        
        if (dateParts.length > 0) {
            metaHtml += `<div style="font-size:0.8em; color:#777; margin-top:2px;">📅 ${dateParts.join(' <span style="color:#ccc;">|</span> ')}</div>`;
        }

        if (tester) {
          metaHtml += `<div style="margin-top:2px;"><span style="font-size:0.75em; background:#eef; color:#336; padding:1px 6px; border-radius:4px; border:1px solid #dde;">👤 Tester: ${tester}</span></div>`;
        }

        if (collaborators) {
            metaHtml += `<div style="margin-top:2px;"><span style="font-size:0.75em; background:#fff3cd; color:#856404; padding:1px 6px; border-radius:4px; border:1px solid #ffeeba;">👥 Team: ${collaborators}</span></div>`;
        }
        
        // Goal
        let goalHtml = '';
        if (goal) {
            goalHtml = `<div class="item-goal">🎯 <strong>Goal:</strong> ${linkify(goal)}</div>`;
        }
        
        // Attachments
        let attachmentHtml = '';
        if (attachments.length > 0) {
            attachmentHtml = '<div style="margin-top:4px;">';
            attachments.forEach(att => {
                attachmentHtml += `<div class="item-attachment">🔗 <a href="${att.url}" target="_blank">${att.name}</a></div> `;
            });
            attachmentHtml += '</div>';
        }

        // Private Comments HTML
        let itCommentsHtml = '';
        if (itComments && showPrivate) {
            let safeIT = itComments;
            if (!itComments.trim().startsWith('<')) {
                safeIT = linkify(itComments);
            }
            itCommentsHtml = `<div class="item-it-comment">🔒 <strong>Private Note:</strong> ${safeIT}</div>`;
        }

        // Comments HTML
        const commentCount = publicComments.length;
        const commentLabel = commentCount > 0 ? `💬 View/Add Comments (${commentCount})` : `💬 Add Question/Comment`;
        
        let commentsListHtml = '';
        publicComments.forEach(c => {
            let timeStr = '';
            if(c.timestamp) {
                const d = new Date(c.timestamp);
                timeStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            }
            
            let safeComment = c.text;
            if (!safeComment.trim().startsWith('<')) {
                safeComment = linkify(safeComment);
            }

            commentsListHtml += `
                <div class="comment-bubble">
                    <div class="comment-header">
                        <span class="comment-author">${linkify(c.author)}</span>
                        <span>${timeStr}</span>
                    </div>
                    <div>${safeComment}</div>
                </div>
            `;
        });

        const commentsSectionHtml = `
            <div class="comments-section">
                <button class="comment-toggle" onclick="toggleComments('comments-${uniqueId}')">${commentLabel}</button>
                <div id="comments-${uniqueId}" class="comments-container">
                    ${commentsListHtml}
                    <div class="comment-form">
                        <input type="text" class="comment-input-name" placeholder="Your Name" maxlength="20">
                        <div id="editor-container-${uniqueId}"></div>
                        <button class="btn-post" onclick="postComment('${listType}', ${index}, '${uniqueId}')">Post</button>
                    </div>
                </div>
            </div>
        `;

        // Render Item
        if (typeof item === 'object') {
            let color;
            switch (status) {
                case 'On Track': color = 'green'; break;
                case 'Testing': color = '#ff9f43'; break;
                case 'Delayed': color = 'red'; break;
                case 'On-Hold': color = 'grey'; break;
                case 'Completed': color = '#800080'; break;
                case 'Follow-Up': color = '#fd7e14'; break;
                case 'Maintenance': color = '#6610f2'; break;
                case 'Stable': color = '#007bff'; break;
                default: color = 'grey';
            }

            let pColor = '#777';
            if (priority === 'High') pColor = '#d9534f'; 
            if (priority === 'Medium') pColor = '#f0ad4e'; 
            if (priority === 'Low') pColor = '#5cb85c'; 

            const priorityBadge = priority ? `<span style="font-size:0.8em; color:${pColor}; border:1px solid ${pColor}; padding:0 4px; border-radius:4px; margin-left:5px;">${priority}</span>` : '';

            const statusBadge = status ? `<span style="font-size:0.8em; color:${color}; border:1px solid ${color}; padding:0 4px; border-radius:4px; margin-left:5px;">${status}</span>` : '';
            const milestoneHtml = milestone ? `<small style="color:#999; font-size:0.8em; display:block;">🏁 Next Milestone: ${linkify(milestone)}</small>` : '';

            html += `<li style="margin-bottom:15px;">
                <div style="margin-bottom:2px;">
                    <strong>${safeName}</strong>${statusBadge}${priorityBadge}${updatedBadge}
                </div>
                ${checkHtml}
                <div style="color:#666; margin-bottom:2px;">${safeNotes}</div>
                ${goalHtml}
                ${attachmentHtml}
                ${itCommentsHtml}
                ${milestoneHtml}
                ${metaHtml}
                ${commentsSectionHtml}
            </li>`;
        } else {
             html += `<li style="margin-bottom:5px;"><strong>${safeName}</strong></li>`;
        }
    });
    
    el.innerHTML = html + '</ol>';
};

// --- OUTLOOK DATA LOADER ---
function loadOutlookData(reportId) {
    const outlookDocId = reportId + "_outlook";
    db.collection("briefings").doc(outlookDocId).onSnapshot((doc) => {
        const headerMeetings = document.getElementById('header-meetings');
        const headerEmails = document.getElementById('header-emails');
        const meetingsContent = document.getElementById('content-meetings');
        const emailsContent = document.getElementById('content-emails');

        if (doc.exists) {
            const data = doc.data();
            currentOutlookData = data;
            headerMeetings.textContent = `Meetings (${data.meetings_count || 0})`;
            headerEmails.textContent   = `Emails (${data.emails_count || 0})`;
            meetingsContent.innerHTML = linkify(data.meetings || "") || "<i>No meetings found.</i>";
            emailsContent.innerHTML   = linkify(data.emails || "")   || "<i>No unread emails.</i>";
        } else {
            headerMeetings.textContent = `Meetings (0)`;
            headerEmails.textContent   = `Unread Emails (0)`;
            meetingsContent.innerHTML = "<i>Waiting for Outlook Sync...</i>";
            emailsContent.innerHTML   = "<i>Waiting for Outlook Sync...</i>";
        }
    }, (error) => {
         console.error("Error Outlook Sync:", error);
         document.getElementById('content-meetings').innerHTML = "<i>Error loading meetings.</i>";
         document.getElementById('content-emails').innerHTML = "<i>Error loading emails.</i>";
    });
}


// --- CORE RENDER FUNCTION ---
function renderReport(data, isDailyMode) {
    document.getElementById('default-message').classList.add('hidden');
    document.getElementById('lock-screen').classList.add('hidden');
    document.getElementById('loading-overlay').classList.add('hidden');
    document.getElementById('nav-links').classList.remove('hidden');
    document.getElementById('report-body').classList.remove('hidden');
    document.getElementById('report-subtitle').textContent = "Report: " + data.reportId;
    
    currentReportData = data;
    
    const hasPasscode = (data.passcode && data.passcode.trim() !== "");
    const showPrivate = isDailyMode && hasPasscode;
    currentShowPrivate = showPrivate;

    let updateTime = "Unknown";
    if (data.lastUpdated) {
        const dateObj = data.lastUpdated.toDate ? data.lastUpdated.toDate() : new Date(data.lastUpdated);
        if (!isNaN(dateObj)) {
            updateTime = dateObj.toLocaleString('en-US', {
                year: 'numeric', month: 'long', day: 'numeric',
                hour: '2-digit', minute: '2-digit'
            });
        } else {
            updateTime = data.lastUpdated;
        }
    }

    if (isDailyMode) {
        document.getElementById('container-meetings').classList.remove('hidden');
        document.getElementById('container-emails').classList.remove('hidden');
    } else {
        document.getElementById('container-meetings').classList.add('hidden');
        document.getElementById('container-emails').classList.add('hidden');
    }

    const dailyTasksData = data.structuredDailyTasks || data.dailyTasks;
    const projectsData = data.structuredProjects || data.projects;
    const activeData = data.structuredActiveTasks || data.activeTasks;

    renderList('content-tasks', dailyTasksData || [], showPrivate);
    renderList('content-projects', projectsData || [], showPrivate);
    renderList('content-active', activeData || [], showPrivate);
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

function attemptUnlock() {
    const entered = document.getElementById('unlock-pass').value;
    if (entered === pendingData.passcode) {
        document.getElementById('lock-screen').classList.add('hidden');
        renderReport(pendingData, true);
        loadOutlookData(targetId);
    } else {
        document.getElementById('unlock-error').style.display = 'block';
    }
}

function cancelUnlock() {
    window.location.href = window.location.pathname;
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

function toggleAnalyticsView() {
    document.getElementById('container-tasks').classList.add('hidden');
    document.getElementById('container-projects').classList.add('hidden');
    document.getElementById('container-active').classList.add('hidden');
    document.getElementById('container-meetings').classList.add('hidden');
    document.getElementById('container-emails').classList.add('hidden');
    document.getElementById('container-analytics').classList.remove('hidden');
    renderPublicAnalytics();
}

function closeAnalytics() {
    document.getElementById('container-analytics').classList.add('hidden');
    document.getElementById('container-tasks').classList.remove('hidden');
    document.getElementById('container-projects').classList.remove('hidden');
    document.getElementById('container-active').classList.remove('hidden');
    
    if (currentShowPrivate) {
        document.getElementById('container-meetings').classList.remove('hidden');
        document.getElementById('container-emails').classList.remove('hidden');
    }
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

    const statusCounts = { 
        'On Track': 0, 
        'Requirement Gathering': 0,
        'Development': 0,
        'Testing': 0, 
        'Delayed': 0, 
        'Future': 0,
        'On-Hold': 0, 
        'Completed': 0, 
        'Follow-Up': 0, 
        'Maintenance': 0, 
        'Stable': 0,
        'Other': 0 
    };
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
    
    const totalStatus = Object.values(statusCounts).reduce((a, b) => a + b, 0);
    const statusHeader = document.getElementById('chart-header-status');
    if(statusHeader) statusHeader.textContent = `Project Status (Total: ${totalStatus})`;

    const totalWorkload = daily.length + projects.length + active.length;
    const workloadHeader = document.getElementById('chart-header-workload');
    if(workloadHeader) workloadHeader.textContent = `Workload Distribution (Total: ${totalWorkload})`;

    if (statusChart) statusChart.destroy();
    if (priorityChart) priorityChart.destroy();
    if (workloadChart) workloadChart.destroy();

   statusChart = new Chart(document.getElementById('chartStatus'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(statusCounts),
            datasets: [{
                data: Object.values(statusCounts),
                // UPDATED: Colors mapped to your CSS classes in admin.html
                backgroundColor: [
                    '#28a745', // On Track (Green)
                    '#17a2b8', // Requirement Gathering (Teal)
                    '#6610f2', // Development (Indigo)
                    '#ff9f43', // Testing (Orange)
                    '#dc3545', // Delayed (Red)
                    '#adb5bd', // Future (Grey-Blue)
                    '#6c757d', // On-Hold (Grey)
                    '#800080', // Completed (Purple)
                    '#fd7e14', // Follow-Up (Orange)
                    '#6610f2', // Maintenance (Violet)
                    '#007bff', // Stable (Blue)
                    '#e2e6ea'  // Other
                ]
            }]
        },
        options: {
            plugins: {
                datalabels: {
                    color: '#ffffff',
                    font: { weight: 'bold' }
                }
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
            labels: ['Daily Priorities', 'Projects', 'Active Tasks'],
            datasets: [{
                data: [daily.length, projects.length, active.length],
                backgroundColor: ['#007bff', '#6610f2', '#fd7e14']
            }]
        },
        options: {
            plugins: {
                datalabels: {
                    color: '#ffffff',
                    font: { weight: 'bold' }
                }
            }
        }
    });
}

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('userid-input').addEventListener("keypress", function(event) {
        if (event.key === "Enter") viewReport();
    });

    document.getElementById('unlock-pass').addEventListener("keypress", function(event) {
        if (event.key === "Enter") attemptUnlock();
    });
    
    const urlParams = new URLSearchParams(window.location.search);
    const dailyId = urlParams.get('daily');
    const reportId = urlParams.get('report');
    let potentialTargetId = dailyId || reportId;
    targetId = potentialTargetId ? potentialTargetId.toLowerCase() : null;
    const isDailyMode = !!dailyId;

    if (targetId) {
        document.getElementById('link-weekly').href = "?report=" + encodeURIComponent(targetId);
        document.getElementById('link-daily').href = "?daily=" + encodeURIComponent(targetId);

        db.collection('briefings').doc(targetId).get().then(doc => {
            if (doc.exists) {
                const data = doc.data();
                if (isDailyMode) {
                    document.getElementById('report-title').textContent = "Daily Briefing";
                    if (data.passcode && data.passcode.trim() !== "") {
                        pendingData = data;
                        document.getElementById('loading-overlay').classList.add('hidden');
                        document.getElementById('lock-screen').classList.remove('hidden');
                        document.getElementById('report-subtitle').textContent = "Protected";
                        
                        setTimeout(() => {
                            document.getElementById('unlock-pass').focus();
                        }, 100);

                    } else {
                        renderReport(data, true);
                        loadOutlookData(targetId);
                        setLastGenerated();
                    }
                } else {
                    document.getElementById('report-title').textContent = "Weekly Report";
                    renderReport(data, false);
                    setLastGenerated();
                }
            } else {
                document.getElementById('loading-overlay').classList.add('hidden');
                document.getElementById('default-message').classList.remove('hidden');
                document.getElementById('report-subtitle').innerText = "ID not found";
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

            let row = {
                Name: item.name,
                Status: item.status || "",
                Daily_Check_Status: checkStatus,
                Daily_Check_Note: stripHtml(checkNote),
                Priority: item.priority || "",
                Goal: item.goal || "",
                Milestone: item.milestone || "",
                Start: item.startDate || "",
                End: item.endDate || "",
                Updated: item.lastUpdated || "",
                Tester: item.tester || "",
                Collaborators: item.collaborators || "",
                Notes: stripHtml(item.notes || ""),
                Public_Comments: commentsStr,
                Daily_Check_History: checkHistoryStr
            };

            if (currentShowPrivate) {
                row.Private_Comments = stripHtml(item.itComments || "");
            }
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

    // Status Color Map (RGB Hex for Excel)
    const statusColors = {
        'On Track': 'C6EFCE',
        'Requirement Gathering': 'DBEAFE',
        'Development': 'E0E7FF',
        'Testing': 'FFEB9C',
        'Delayed': 'FFC7CE',
        'Future': 'F3F4F6',
        'On-Hold': 'E5E7EB',
        'Completed': 'E9D5FF',
        'Follow-Up': 'FFEDD5',
        'Maintenance': 'F5F3FF',
        'Stable': 'DBEAFE'
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
            if (["Notes", "Public_Comments", "Daily_Check_History", "Attachment", "Private_Comments"].includes(headerVal)) {
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

    // 1. Analytics
    const analyticsData = (function() {
        const projects = currentReportData.structuredProjects || currentReportData.projects || [];
        const active = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
        const daily = currentReportData.structuredDailyTasks || currentReportData.dailyTasks || [];
        const allItems = [...projects, ...active];
        const rows = [
            { Category: "WORKLOAD", Metric: "Daily Tasks", Count: daily.length },
            { Category: "WORKLOAD", Metric: "Active Projects", Count: projects.length },
            { Category: "WORKLOAD", Metric: "Active Tasks", Count: active.length },
            { Category: "", Metric: "", Count: "" },
            { Category: "REPORT INFO", Metric: "User", Count: targetId },
            { Category: "REPORT INFO", Metric: "Exported", Count: new Date().toLocaleString() }
        ];
        return rows;
    })();
    
    const wsAnalytics = XLSX.utils.json_to_sheet(analyticsData);
    applyGlobalStyles(wsAnalytics);
    XLSX.utils.book_append_sheet(wb, wsAnalytics, "Analytics Overview");

    // 2. Data Sheets
    const sheetsToProcess = [
        { name: "Daily Tasks", data: currentReportData.structuredDailyTasks || currentReportData.dailyTasks },
        { name: "Projects", data: currentReportData.structuredProjects || currentReportData.projects },
        { name: "Active Tasks", data: currentReportData.structuredActiveTasks || currentReportData.activeTasks }
    ];

    sheetsToProcess.forEach(sheetObj => {
        const formattedData = formatForExcel(sheetObj.data);
        if (formattedData.length > 0) {
            const ws = XLSX.utils.json_to_sheet(formattedData);
            applyGlobalStyles(ws);
            XLSX.utils.book_append_sheet(wb, ws, sheetObj.name);
        }
    });

    // 3. Outlook
    if (currentShowPrivate && currentOutlookData) {
        const outlookRows = [];
        if(currentOutlookData.meetings) outlookRows.push({ Type: "MEETINGS", Content: stripHtml(currentOutlookData.meetings) });
        if(currentOutlookData.emails) outlookRows.push({ Type: "EMAILS", Content: stripHtml(currentOutlookData.emails) });
        
        if (outlookRows.length > 0) {
            const wsOutlook = XLSX.utils.json_to_sheet(outlookRows);
            applyGlobalStyles(wsOutlook);
            XLSX.utils.book_append_sheet(wb, wsOutlook, "Outlook Data");
        }
    }

    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Briefing_${targetId}_${dateStr}.xlsx`);
}
