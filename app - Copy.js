// --- CONFIGURATIONS ---

const firebaseConfig = {
    apiKey: "AIzaSyCtFf85MUkNSsSsT6Nv8M_09Fphm2DcQOU", 
    authDomain: "dailybriefing-fe7df.firebaseapp.com",
    projectId: "dailybriefing-fe7df",
    storageBucket: "dailybriefing-fe7df.appspot.com",
    messagingSenderId: "",
    appId: ""
};
firebase.initializeApp(firebaseConfig);

// --- GLOBAL VARIABLES ---
const db = firebase.firestore();
let pendingData = null; 
let targetId = null; 

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

const renderList = (id, items) => {
    const el = document.getElementById(id);
    const headerEl = document.getElementById('header-' + id.replace('content-', ''));

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
    
    items.forEach(item => {
        let name, notes = '', status, milestone = '', tester = '', startDate = '', endDate = '', lastUpdated = '';
        
        if (typeof item === 'object' && item !== null && item.name) {
            name = item.name;
            notes = item.notes || '';
            status = item.status;
            milestone = item.milestone;
            tester = item.tester;
            startDate = item.startDate;
            endDate = item.endDate;
            lastUpdated = item.lastUpdated;
        } else if (typeof item === 'string') {
            name = item;
        } else {
            return; 
        }

        // Determine if it's a "Complex" item (Projects/Active Tasks) OR if it has rich data
        const isComplexItem = (status || tester || startDate || endDate);

        // Build Extra Meta Data HTML
        let metaHtml = '';
        let dateParts = [];
        
        // Helper to format date YYYY-MM-DD to cleaner format if desired, or keep as is
        const fmtDate = (d) => d; 

        if (startDate) dateParts.push(`Start: ${fmtDate(startDate)}`);
        if (lastUpdated) dateParts.push(`Updated: ${fmtDate(lastUpdated)}`);
        if (endDate) dateParts.push(`End: ${fmtDate(endDate)}`);
        
        if (dateParts.length > 0) {
            // Using Unicode escape for Calendar Emoji
            metaHtml += `<div style="font-size:0.8em; color:#777; margin-top:2px;">\uD83D\uDCC5 ${dateParts.join(' <span style="color:#ccc;">|</span> ')}</div>`;
        }

        if (tester) {
            // Using Unicode escape for User Icon (Bust in Silhouette)
            metaHtml += `<div style="margin-top:2px;"><span style="font-size:0.75em; background:#eef; color:#336; padding:1px 6px; border-radius:4px; border:1px solid #dde;">\uD83D\uDC64 Tester: ${tester}</span></div>`;
        }


        if (!isComplexItem && !status) {
            // Simple Task (Minimal Display)
            html += `<li style="margin-bottom:5px;">
                <strong>${name}</strong>
                ${metaHtml}
            </li>`;
        } else {
            // Complex Item Display
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

            // Status Badge
            const statusBadge = status ? `<span style="font-size:0.8em; color:${color}; border:1px solid ${color}; padding:0 4px; border-radius:4px; margin-left:5px;">${status}</span>` : '';

            // Using Unicode escape for Chequered Flag
            const milestoneHtml = milestone ? `<small style="color:#999; font-size:0.8em; display:block;">\uD83C\uDFC1 Next Milestone: ${milestone}</small>` : '';

            html += `<li style="margin-bottom:12px;">
                <div style="margin-bottom:2px;">
                    <strong>${name}</strong>${statusBadge}
                </div>
                <small style="color:#666; display:block; margin-bottom:2px;">${notes}</small>
                ${milestoneHtml}
                ${metaHtml}
            </li>`;
        }
    });
    
    el.innerHTML = linkify(html + '</ol>');
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
            headerMeetings.textContent = `4. Meetings (${data.meetings_count || 0})`;
            headerEmails.textContent   = `5. Emails (${data.emails_count || 0})`;
            meetingsContent.innerHTML = linkify(data.meetings || "") || "<i>No meetings found.</i>";
            emailsContent.innerHTML   = linkify(data.emails || "")   || "<i>No unread emails.</i>";
        } else {
            headerMeetings.textContent = `4. This Week's Meetings (0)`;
            headerEmails.textContent   = `5. Unread Emails (0)`;
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

    renderList('content-tasks', dailyTasksData || []);
    renderList('content-projects', projectsData || []);
    renderList('content-active', activeData || []);
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

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('userid-input').addEventListener("keypress", function(event) {
        if (event.key === "Enter") viewReport();
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