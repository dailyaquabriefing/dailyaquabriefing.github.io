// --- CONFIGURATIONS ---
const firebaseConfig = {
    apiKey: "AIzaSyCtFf85MUkNSsSsT6Nv8M_09Fphm2DcQOU",
    authDomain: "dailybriefing-fe7df.firebaseapp.com",
    projectId: "dailybriefing-fe7df",
    storageBucket: "dailybriefing-fe7df.firebasestorage.app",
    messagingSenderId: "",
    appId: ""
};
firebase.initializeApp(firebaseConfig);

// --- GLOBAL VARIABLES ---
const db = firebase.firestore();
let pendingData = null;
let targetId = null;
let currentShowPrivate = false;
let currentReportData = null;
let currentOutlookData = null;
let commentEditors = {};

// CHART INSTANCES
let statusChart = null;
let teamChart = null;
let workloadChart = null;

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

function stripHtml(html) {
   if (!html) return "";
   let tmp = document.createElement("DIV");
   tmp.innerHTML = html;
   return tmp.textContent || tmp.innerText || "";
}

// Helper to format data specifically for Excel Rows
function formatForExcel(dataArray) {
    if (!Array.isArray(dataArray)) return [];
    return dataArray.map(item => {
        // Process Latest 5 Comments
        let comments = "";
        if (item.publicComments && Array.isArray(item.publicComments)) {
            comments = item.publicComments
                .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
                .slice(0, 5)
                .map(c => `${c.author}: ${stripHtml(c.text)}`)
                .join("\r\n");
        }

        // Process Latest 5 Daily Checks
        let checkHistory = "";
        if (item.dailyChecks && Array.isArray(item.dailyChecks)) {
            checkHistory = item.dailyChecks
                .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
                .slice(0, 5)
                .map(c => `[${c.status}] ${stripHtml(c.note || 'No Note')}`)
                .join("\r\n");
        }

        return {
            "Name": item.name || "",
            "Status": item.status || "N/A",
            "Priority": item.priority || "N/A",
            "Goal": stripHtml(item.goal || ""),
            "Next Milestone": item.milestone || "",
            "Notes": stripHtml(item.notes || ""),
            "Daily_Check_Status": item.dailyChecks?.[0]?.status || "N/A",
            "Daily_Check_History": checkHistory,
            "Public_Comments": comments,
            "Attachment": item.attachments ? item.attachments.map(a => a.name).join(", ") : (item.attachment ? "Resource" : ""),
            "Collaborators": item.collaborators || "",
            "Last_Updated": item.lastUpdated || ""
        };
    });
}

// --- COMMENT LOGIC ---

window.toggleComments = function(id) {
    const el = document.getElementById(id);
    if (el) {
        el.classList.toggle('open');
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

window.postComment = function(listType, itemIndex, uniqueId) {
    const container = document.getElementById('comments-' + uniqueId);
    const nameVal = container.querySelector('.comment-input-name').value.trim();
    
    let textVal = "";
    if (commentEditors[uniqueId]) {
        const editorContent = commentEditors[uniqueId].root.innerHTML;
        const textOnly = commentEditors[uniqueId].getText().trim();
        if(textOnly.length > 0) textVal = editorContent;
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
        let listKey = (listType === 'daily') ? 'structuredDailyTasks' : (listType === 'project' ? 'structuredProjects' : 'structuredActiveTasks');
        let listData = data[listKey] || [];

        if (!listData[itemIndex]) return;

        listData[itemIndex].publicComments = listData[itemIndex].publicComments || [];
        listData[itemIndex].publicComments.push({
            author: nameVal,
            text: textVal,
            timestamp: new Date().toISOString()
        });
        listData[itemIndex].lastUpdated = new Date().toLocaleString();

        docRef.update({ [listKey]: listData, lastUpdated: new Date().toISOString() }).then(() => {
            delete commentEditors[uniqueId];
            renderList(listType === 'daily' ? 'content-tasks' : (listType === 'project' ? 'content-projects' : 'content-active'), listData, currentShowPrivate);
        });
    });
};

// --- LIST RENDERING ---

const renderList = (id, items, showPrivate = false) => {
    const el = document.getElementById(id);
    const headerEl = document.getElementById('header-' + id.replace('content-', ''));
    let listType = id.includes('tasks') ? 'daily' : (id.includes('projects') ? 'project' : 'active');

    if (headerEl) {
        const baseTitle = headerEl.id.includes('tasks') ? 'Daily Tasks' : (headerEl.id.includes('projects') ? 'Active Projects' : 'Active Tasks');
        headerEl.textContent = `${baseTitle} (${items ? items.length : 0})`;
    }

    if (!Array.isArray(items) || !items.length) {
        el.innerHTML = "<em>No items.</em>";
        return;
    }
    
    let html = '<ol style="padding-left:20px;">';
    items.forEach((item, index) => {
        const name = typeof item === 'object' ? item.name : item;
        const uniqueId = `${listType}-${index}`;
        const status = item.status || 'Other';
        
        let color;
        switch (status) {
            case 'On Track': color = 'green'; break;
            case 'Testing': color = '#ff9f43'; break;
            case 'Delayed': color = 'red'; break;
            case 'Completed': color = '#800080'; break;
            default: color = 'grey';
        }

        // Simplifed render for brevity in this snippet, following your logic
        html += `<li style="margin-bottom:15px;">
            <strong>${linkify(name)}</strong> 
            <span style="font-size:0.8em; color:${color}; border:1px solid ${color}; padding:0 4px; border-radius:4px; margin-left:5px;">${status}</span>
            <div style="color:#666; font-size:0.9em;">${linkify(item.notes || "")}</div>
            <div class="comments-section">
                <button class="comment-toggle" onclick="toggleComments('comments-${uniqueId}')">💬 Comments (${item.publicComments ? item.publicComments.length : 0})</button>
                <div id="comments-${uniqueId}" class="comments-container">
                    <div class="comment-form">
                        <input type="text" class="comment-input-name" placeholder="Your Name">
                        <div id="editor-container-${uniqueId}"></div>
                        <button class="btn-post" onclick="postComment('${listType}', ${index}, '${uniqueId}')">Post</button>
                    </div>
                </div>
            </div>
        </li>`;
    });
    el.innerHTML = html + '</ol>';
};

// --- ANALYTICS ---

function renderPublicAnalytics() {
    if (!currentReportData) return;
    if (typeof ChartDataLabels !== 'undefined') Chart.register(ChartDataLabels);

    const projects = currentReportData.structuredProjects || currentReportData.projects || [];
    const active = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
    const daily = currentReportData.structuredDailyTasks || currentReportData.dailyTasks || [];
    const allItems = [...projects, ...active];

    const statusCounts = { 'On Track': 0, 'Testing': 0, 'Delayed': 0, 'Completed': 0, 'Other': 0 };
    allItems.forEach(item => {
        const s = item.status || 'Other';
        statusCounts[statusCounts.hasOwnProperty(s) ? s : 'Other']++;
    });

    const healthyCount = (statusCounts['On Track'] || 0) + (statusCounts['Completed'] || 0);
    const healthScore = allItems.length > 0 ? Math.round((healthyCount / allItems.length) * 100) : 0;
    
    const healthEl = document.getElementById('health-score-display');
    if (healthEl) {
        healthEl.textContent = `${healthScore}%`;
        healthEl.style.color = healthScore > 70 ? '#28a745' : (healthScore > 40 ? '#ff9f43' : '#dc3545');
    }

    if (statusChart) statusChart.destroy();
    statusChart = new Chart(document.getElementById('chartStatus'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(statusCounts),
            datasets: [{ data: Object.values(statusCounts), backgroundColor: ['#28a745', '#ff9f43', '#dc3545', '#800080', '#6c757d'] }]
        }
    });

    if (workloadChart) workloadChart.destroy();
    workloadChart = new Chart(document.getElementById('chartWorkload'), {
        type: 'radar',
        data: {
            labels: ['Daily Tasks', 'Projects', 'Active Tasks'],
            datasets: [{ label: 'Volume', data: [daily.length, projects.length, active.length], backgroundColor: 'rgba(0, 120, 212, 0.2)', borderColor: '#0078D4' }]
        }
    });
}

function toggleAnalyticsView() {
    ['container-tasks', 'container-projects', 'container-active', 'container-meetings', 'container-emails'].forEach(id => {
        const el = document.getElementById(id);
        if(el) el.classList.add('hidden');
    });
    document.getElementById('container-analytics').classList.remove('hidden');
    renderPublicAnalytics();
}

function closeAnalytics() {
    document.getElementById('container-analytics').classList.add('hidden');
    ['container-tasks', 'container-projects', 'container-active'].forEach(id => document.getElementById(id).classList.remove('hidden'));
}

// --- EXCEL EXPORT & STYLING ---

const applyGlobalStyles = (ws) => {
    if (!ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const COLUMN_WIDTH_CHARS = 45;
    const DEFAULT_ROW_HEIGHT = 20; 
    const LINE_HEIGHT_PTS = 15;

    const statusColors = {
        'On Track': 'C6EFCE', 'Requirement Gathering': 'DBEAFE', 'Development': 'E0E7FF',
        'Testing': 'FFEB9C', 'Delayed': 'FFC7CE', 'Future': 'F3F4F6',
        'On-Hold': 'E5E7EB', 'Completed': 'E9D5FF', 'Follow-Up': 'FFEDD5',
        'Maintenance': 'F5F3FF', 'Stable': 'DBEAFE'
    };

    let statusColIdx = -1;
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const header = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
        if (header && header.v === "Status") { statusColIdx = C; break; }
    }

    for (let R = range.s.r; R <= range.e.r; ++R) {
        let rowColor = 'FFFFFF'; 
        if (R > 0 && statusColIdx !== -1) {
            const statusCell = ws[XLSX.utils.encode_cell({ r: R, c: statusColIdx })];
            if (statusCell && statusColors[statusCell.v]) rowColor = statusColors[statusCell.v];
        }

        for (let C = range.s.c; C <= range.e.c; ++C) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[addr]) continue;
            if (!ws[addr].s) ws[addr].s = {};
            
            ws[addr].s.border = {
                top: { style: "thin", color: { rgb: "D1D5DB" } },
                bottom: { style: "thin", color: { rgb: "D1D5DB" } },
                left: { style: "thin", color: { rgb: "D1D5DB" } },
                right: { style: "thin", color: { rgb: "D1D5DB" } }
            };
            ws[addr].s.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
            
            if (R > 0) ws[addr].s.fill = { patternType: "solid", fgColor: { rgb: rowColor } };
            if (R === 0) {
                ws[addr].s.font = { bold: true };
                ws[addr].s.fill = { patternType: "solid", fgColor: { rgb: "F2F2F2" } };
            }

            const headerVal = ws[XLSX.utils.encode_cell({ r: 0, c: C })]?.v || "";
            if (["Notes", "Public_Comments", "Daily_Check_History", "Attachment"].includes(headerVal)) {
                ws['!cols'][C] = { wch: COLUMN_WIDTH_CHARS };
            }

            if (R > 0) {
                const lineCount = (String(ws[addr].v).match(/\r\n/g) || []).length + 1;
                const targetLines = Math.min(lineCount, 5); 
                const h = Math.max(DEFAULT_ROW_HEIGHT, targetLines * LINE_HEIGHT_PTS);
                if (!ws['!rows'][R] || h > ws['!rows'][R].hpt) ws['!rows'][R] = { hpt: h };
            }
        }
    }
};

function exportReportToExcel() {
    if (!currentReportData) { alert("No data loaded."); return; }
    const wb = XLSX.utils.book_new();

    // 1. Analytics
    const getAnalyticsData = () => {
        const projects = currentReportData.structuredProjects || currentReportData.projects || [];
        const active = currentReportData.structuredActiveTasks || currentReportData.activeTasks || [];
        const healthyCount = [...projects, ...active].filter(i => ['On Track', 'Completed'].includes(i.status)).length;
        const healthScore = Math.round((healthyCount / ([...projects, ...active].length || 1)) * 100);

        return [
            { Category: "KPI", Metric: "Project Health Index", Value: healthScore + "%" },
            { Category: "SYSTEM", Metric: "User ID", Value: targetId },
            { Category: "SYSTEM", Metric: "Export Date", Value: new Date().toLocaleString() }
        ];
    };
    
    const wsAnalytics = XLSX.utils.json_to_sheet(getAnalyticsData());
    applyGlobalStyles(wsAnalytics);
    XLSX.utils.book_append_sheet(wb, wsAnalytics, "Analytics");

    // 2. Data Sheets
    const sheets = [
        { name: "Daily Tasks", data: currentReportData.structuredDailyTasks || currentReportData.dailyTasks || [] },
        { name: "Projects", data: currentReportData.structuredProjects || currentReportData.projects || [] },
        { name: "Active Tasks", data: currentReportData.structuredActiveTasks || currentReportData.activeTasks || [] }
    ];

    sheets.forEach(sheetObj => {
        if (sheetObj.data.length > 0) {
            const ws = XLSX.utils.json_to_sheet(formatForExcel(sheetObj.data));
            applyGlobalStyles(ws);
            XLSX.utils.book_append_sheet(wb, ws, sheetObj.name);
        }
    });

    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Briefing_${targetId}_${dateStr}.xlsx`);
}

// --- INITIALIZATION ---

function viewReport() {
    const val = document.getElementById('userid-input').value.toLowerCase().trim();
    if (val) window.location.href = "?report=" + encodeURIComponent(val);
}

document.addEventListener('DOMContentLoaded', () => {
    const urlParams = new URLSearchParams(window.location.search);
    const id = urlParams.get('report') || urlParams.get('daily');
    if (id) {
        targetId = id.toLowerCase();
        db.collection('briefings').doc(targetId).onSnapshot(doc => {
            if (doc.exists) {
                currentReportData = doc.data();
                renderReport(currentReportData, !!urlParams.get('daily'));
            }
        });
    }
});

function renderReport(data, isDaily) {
    document.getElementById('loading-overlay').classList.add('hidden');
    document.getElementById('default-message').classList.add('hidden');
    document.getElementById('nav-links').classList.remove('hidden');
    document.getElementById('report-body').classList.remove('hidden');
    
    renderList('content-tasks', data.structuredDailyTasks || data.dailyTasks);
    renderList('content-projects', data.structuredProjects || data.projects);
    renderList('content-active', data.structuredActiveTasks || data.activeTasks);
}