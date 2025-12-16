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

// Toggle comment visibility
window.toggleComments = function(id) {
    const el = document.getElementById(id);
    if (el) {
        el.classList.toggle('open');
        // Save Name Preference
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
    const textVal = container.querySelector('.comment-input-text').value.trim();

    if (!nameVal || !textVal) {
        alert("Please enter both your Name and a Comment.");
        return;
    }

    // Save name for next time
    localStorage.setItem('commenterName', nameVal);

    // Fetch current data
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

        docRef.update({
            [listKey]: listData
        }).then(() => {
            renderList(containerId, listData, currentShowPrivate);
            const newContainer = document.getElementById('comments-' + uniqueId);
            if (newContainer) {
                newContainer.classList.add('open');
                newContainer.querySelector('.comment-input-text').value = '';
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
        let name, notes = '', status, priority = '', milestone = '', tester = '', startDate = '', endDate = '', lastUpdated = '', goal = '', attachment = '', itComments = '', publicComments = [];
        
        if (typeof item === 'object' && item !== null && item.name) {
            name = item.name;
            notes = item.notes || '';
            status = item.status;
            priority = item.priority; 
            milestone = item.milestone;
            tester = item.tester;
            startDate = item.startDate;
            endDate = item.endDate;
            lastUpdated = item.lastUpdated;
            goal = item.goal;
            attachment = item.attachment;
            itComments = item.itComments;
            publicComments = item.publicComments || [];

        } else if (typeof item === 'string') {
            name = item;
        } else {
            return;
        }

        const safeName = linkify(name);
        const safeNotes = linkify(notes);
        const uniqueId = `${listType}-${index}`;

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
        
        // Goal and Attachment HTML
        let goalHtml = '';
        if (goal) {
            goalHtml = `<div class="item-goal">🎯 <strong>Goal:</strong> ${linkify(goal)}</div>`;
        }
        
        let attachmentHtml = '';
        if (attachment) {
            attachmentHtml = `<div class="item-attachment">🔗 <a href="${attachment}" target="_blank">View Resource</a></div>`;
        }

        // IT Comments HTML (Only if showPrivate is true)
        let itCommentsHtml = '';
        if (itComments && showPrivate) {
            itCommentsHtml = `<div class="item-it-comment">🔒 <strong>IT Only:</strong> ${linkify(itComments)}</div>`;
        }

        // --- COMMENTS HTML ---
        const commentCount = publicComments.length;
        const commentLabel = commentCount > 0 ? `💬 View/Add Comments (${commentCount})` : `💬 Add Question/Comment`;
        
        let commentsListHtml = '';
        publicComments.forEach(c => {
            let timeStr = '';
            if(c.timestamp) {
                const d = new Date(c.timestamp);
                timeStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            }
            commentsListHtml += `
                <div class="comment-bubble">
                    <div class="comment-header">
                        <span class="comment-author">${linkify(c.author)}</span>
                        <span>${timeStr}</span>
                    </div>
                    <div>${linkify(c.text)}</div>
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
                        <input type="text" class="comment-input-text" placeholder="Type a question or comment...">
                        <button class="btn-post" onclick="postComment('${listType}', ${index}, '${uniqueId}')">Post</button>
                    </div>
                </div>
            </div>
        `;

        // --- RENDER ITEM ---
        if (typeof item === 'object') {
            let color;
            switch (status) {
                case 'On Track': color = 'green'; break;
                case 'Delayed': color = 'red'; break;
                case 'On-Hold': color = 'grey'; break;
                case 'Completed': color = '#800080'; break;
                case 'Follow-Up': color = '#fd7e14'; break;
                case 'Maintenance': color = '#6610f2'; break;
                case 'Stable': color = '#007bff'; break;
                default: color = 'grey';
            }

            let pColor = '#777';
            if (priority === 'High') pColor = '#d9534f'; // Red
            if (priority === 'Medium') pColor = '#f0ad4e'; // Orange
            if (priority === 'Low') pColor = '#5cb85c'; // Green

            const priorityBadge = priority ? `<span style="font-size:0.8em; color:${pColor}; border:1px solid ${pColor}; padding:0 4px; border-radius:4px; margin-left:5px;">${priority}</span>` : '';

            const statusBadge = status ? `<span style="font-size:0.8em; color:${color}; border:1px solid ${color}; padding:0 4px; border-radius:4px; margin-left:5px;">${status}</span>` : '';
            const milestoneHtml = milestone ? `<small style="color:#999; font-size:0.8em; display:block;">🏁 Next Milestone: ${linkify(milestone)}</small>` : '';

            html += `<li style="margin-bottom:15px;">
                <div style="margin-bottom:2px;">
                    <strong>${safeName}</strong>${statusBadge}${priorityBadge}
                </div>
                <small style="color:#666; display:block; margin-bottom:2px;">${safeNotes}</small>
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
    // Hide Lists
    document.getElementById('container-tasks').classList.add('hidden');
    document.getElementById('container-projects').classList.add('hidden');
    document.getElementById('container-active').classList.add('hidden');
    document.getElementById('container-meetings').classList.add('hidden');
    document.getElementById('container-emails').classList.add('hidden');
    
    // Show Analytics
    document.getElementById('container-analytics').classList.remove('hidden');
    
    renderPublicAnalytics();
}

function closeAnalytics() {
    // Hide Analytics
    document.getElementById('container-analytics').classList.add('hidden');
    
    // Show Lists
    document.getElementById('container-tasks').classList.remove('hidden');
    document.getElementById('container-projects').classList.remove('hidden');
    document.getElementById('container-active').classList.remove('hidden');
    
    // Restore Outlook sections ONLY if we are in private mode
    if (currentShowPrivate) {
        document.getElementById('container-meetings').classList.remove('hidden');
        document.getElementById('container-emails').classList.remove('hidden');
    }
}

function renderPublicAnalytics() {
    if (!currentReportData) return;

    // Register DataLabels Plugin (Safe check)
    if (typeof ChartDataLabels !== 'undefined') {
        Chart.register(ChartDataLabels);
    }

    // 1. Consolidate Data (Safe Public Data Only)
    const projects = currentReportData.structuredProjects || currentReportData.projects || [];
    const active = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
    const daily = currentReportData.structuredDailyTasks || currentReportData.dailyTasks || [];
    
    const allItems = [...projects, ...active];

    // 2. Calculate Counts
    const statusCounts = { 'On Track': 0, 'Delayed': 0, 'Completed': 0, 'On-Hold': 0, 'Other': 0 };
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

    // --- CALCULATE TOTALS AND UPDATE HEADERS ---
    
    // Project Status Total
    const totalStatus = Object.values(statusCounts).reduce((a, b) => a + b, 0);
    const statusHeader = document.getElementById('chart-header-status');
    if(statusHeader) statusHeader.textContent = `Project Status (Total: ${totalStatus})`;

    // Workload Total
    const totalWorkload = daily.length + projects.length + active.length;
    const workloadHeader = document.getElementById('chart-header-workload');
    if(workloadHeader) workloadHeader.textContent = `Workload Distribution (Total: ${totalWorkload})`;

    // Priority Total (Optional, but good for consistency)
    const totalPriority = Object.values(prioCounts).reduce((a, b) => a + b, 0);
    const priorityHeader = document.getElementById('chart-header-priority');
    if(priorityHeader) priorityHeader.textContent = `Priority Breakdown (Total: ${totalPriority})`;


    // 3. Destroy Old Charts
    if (statusChart) statusChart.destroy();
    if (priorityChart) priorityChart.destroy();
    if (workloadChart) workloadChart.destroy();

    // 4. Render Charts
    // Status Chart
    statusChart = new Chart(document.getElementById('chartStatus'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(statusCounts),
            datasets: [{
                data: Object.values(statusCounts),
                backgroundColor: ['#28a745', '#dc3545', '#17a2b8', '#6c757d', '#e2e6ea']
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

    // Priority Chart
    const pKeys = Object.keys(prioCounts);
    const pValues = Object.values(prioCounts);
    
    // Create Custom Labels: "High (9)", "Medium (3)", "Low (13)"
    const pLabels = pKeys.map((key, i) => `${key} (${pValues[i]})`);

    priorityChart = new Chart(document.getElementById('chartPriority'), {
        type: 'bar',
        data: {
            labels: pLabels, // Use the new labels with counts
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
                // DISABLE DataLabels for Bar Chart (since we moved count to the Axis)
                datalabels: {
                    display: false
                }
            }
        }
    });

    // Workload Chart
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
function exportReportToExcel() {
    if (!currentReportData) {
        alert("No data loaded to export.");
        return;
    }

    // --- HELPER 1: FORMAT DATA & NEWLINES ---
    const formatForExcel = (list) => {
        if (!Array.isArray(list)) return [];
        return list.map(item => {
            let commentsStr = "";
            // Use \r\n for line breaks which Excel recognizes inside cells
            if (item.publicComments && item.publicComments.length > 0) {
                commentsStr = [...item.publicComments].reverse().map(c => {
                    let timeStr = "N/A";
                    if (c.timestamp) {
                        const d = new Date(c.timestamp);
                        timeStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
                    }
                    return `[${timeStr}]: [${c.author}]: ${c.text}`;
                }).join("\r\n");
            }

            // Define the base row order (Attachment removed from middle)
            let row = {
                Name: item.name,
                Status: item.status || "",
                Priority: item.priority || "",
                Goal: item.goal || "",
                Milestone: item.milestone || "",
                Start: item.startDate || "",
                End: item.endDate || "",
                Updated: item.lastUpdated || "",
                Tester: item.tester || "",
                Notes: item.notes || "",
                Public_Comments: commentsStr 
            };

            // Add Private Comments next (if applicable)
            if (currentShowPrivate) {
                row.IT_Private_Comments = item.itComments || "";
            }

            // Add Attachment LAST
            row.Attachment = item.attachment || "";

            return row;
        });
    };

    // --- HELPER 2: APPLY STYLES (WRAP TEXT) ---
    const applyColumnStyles = (ws, targetHeader) => {
        if (!ws['!ref']) return;
        const range = XLSX.utils.decode_range(ws['!ref']);
        
        // 1. Find the Column Index for the Header
        let colIndex = -1;
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const address = XLSX.utils.encode_cell({ r: range.s.r, c: C });
            if (ws[address] && ws[address].v === targetHeader) {
                colIndex = C;
                break;
            }
        }
        if (colIndex === -1) return; // Header not found

        // 2. Set Column Width (approx 60 chars)
        if (!ws['!cols']) ws['!cols'] = [];
        // Fill empty slots to prevent sparse array errors
        for (let i = 0; i <= range.e.c; i++) { if (!ws['!cols'][i]) ws['!cols'][i] = { wch: 10 }; }
        ws['!cols'][colIndex] = { wch: 60 }; 

        // 3. Iterate Rows and Apply Wrap Style
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            const address = XLSX.utils.encode_cell({ r: R, c: colIndex });
            
            // If cell doesn't exist (empty), create it so styling applies
            if (!ws[address]) ws[address] = { t: 's', v: '' };

            // Apply Style Object
            if (!ws[address].s) ws[address].s = {};
            
            ws[address].s.alignment = { 
                wrapText: true, 
                vertical: 'top', 
                horizontal: 'left' 
            };
        }
    };

    // --- MAIN EXPORT LOGIC ---
    const wb = XLSX.utils.book_new();

    // 1. Generate Analytics
    const generateAnalyticsSheet = () => {
        const projects = currentReportData.structuredProjects || currentReportData.projects || [];
        const active = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
        const daily = currentReportData.structuredDailyTasks || currentReportData.dailyTasks || [];
        const allItems = [...projects, ...active];

        const statusCounts = {};
        const prioCounts = {};
        
        allItems.forEach(i => {
            const s = i.status || 'No Status';
            statusCounts[s] = (statusCounts[s] || 0) + 1;
            const p = i.priority || 'No Priority';
            prioCounts[p] = (prioCounts[p] || 0) + 1;
        });

        const rows = [
            { Category: "WORKLOAD", Metric: "Daily Tasks", Count: daily.length },
            { Category: "WORKLOAD", Metric: "Active Projects", Count: projects.length },
            { Category: "WORKLOAD", Metric: "Active Tasks", Count: active.length },
            { Category: "", Metric: "", Count: "" }
        ];

        Object.keys(statusCounts).forEach(k => rows.push({ Category: "STATUS BREAKDOWN", Metric: k, Count: statusCounts[k] }));
        rows.push({ Category: "", Metric: "", Count: "" });
        Object.keys(prioCounts).forEach(k => rows.push({ Category: "PRIORITY BREAKDOWN", Metric: k, Count: prioCounts[k] }));

        // --- NEW: FOOTER ROWS ---
        // 1. Spacing (Two empty rows)
        rows.push({ Category: "", Metric: "", Count: "" });
        rows.push({ Category: "", Metric: "", Count: "" });
        
        // 2. Report ID (Label in Col A, ID in Col B)
        rows.push({ Category: "Report for:", Metric: targetId || "Unknown", Count: "" });
        
        // 3. Current Date/Time (Label in Col A, Date in Col B)
        rows.push({ Category: "Date Created:", Metric: new Date().toLocaleString(), Count: "" });

        return rows;
    };

    const analyticsData = generateAnalyticsSheet();
    if(analyticsData.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analyticsData), "Analytics Overview");

    // 2. Generate Data Sheets
    const sheetsToProcess = [
        { name: "Daily Tasks", data: currentReportData.structuredDailyTasks || currentReportData.dailyTasks },
        { name: "Projects", data: currentReportData.structuredProjects || currentReportData.projects },
        { name: "Active Tasks", data: currentReportData.structuredActiveTasks || currentReportData.activeTasks }
    ];

    sheetsToProcess.forEach(sheetObj => {
        const formattedData = formatForExcel(sheetObj.data);
        if (formattedData.length > 0) {
            const ws = XLSX.utils.json_to_sheet(formattedData);
            
            // Apply Styling
            applyColumnStyles(ws, "Public_Comments");
            if(currentShowPrivate) applyColumnStyles(ws, "IT_Private_Comments");

            XLSX.utils.book_append_sheet(wb, ws, sheetObj.name);
        }
    });

    // 3. Outlook Data
    if (currentShowPrivate && currentOutlookData) {
        const outlookRows = [];
        if(currentOutlookData.meetings) {
             const cleanMeetings = currentOutlookData.meetings.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>?/gm, '');
             outlookRows.push({ Type: "MEETINGS", Content: cleanMeetings });
        }
        if(currentOutlookData.emails) {
             const cleanEmails = currentOutlookData.emails.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>?/gm, '');
             outlookRows.push({ Type: "EMAILS", Content: cleanEmails });
        }
        if (outlookRows.length > 0) {
            const wsOutlook = XLSX.utils.json_to_sheet(outlookRows);
            applyColumnStyles(wsOutlook, "Content"); 
            XLSX.utils.book_append_sheet(wb, wsOutlook, "Outlook Data");
        }
    }

    const dateStr = new Date().toISOString().split('T')[0];
    const mode = currentShowPrivate ? "Private_Briefing" : "Public_Report";
    XLSX.writeFile(wb, `${mode}_${targetId}_${dateStr}.xlsx`);
}