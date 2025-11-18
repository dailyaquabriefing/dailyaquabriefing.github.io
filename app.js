/**
 * Converts URLs and emails in an HTML string to clickable HTML links.
 */
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

// --- Helper: Generate HTML from Arrays (For New Admin Data) ---
function generateHtmlFromStructuredData(dataArray, type) {
    if (!dataArray || dataArray.length === 0) return "";

    if (type === 'simpleList') {
        // For Daily Tasks
        let html = "";
        dataArray.forEach((task, i) => {
            html += `${i + 1}. ${task}\n`; 
        });
        return html.replace(/\n/g, "<br>");
    }

    if (type === 'projectList') {
        // For Projects and Active Tasks
        let html = "";
        dataArray.forEach((item, i) => {
            html += `<b>${i + 1}. ${type === 'projects' ? 'Project/Task Name' : 'Task Description'}:</b>&nbsp;${item.name}<br>`;
            html += `<b>Status|Priority:</b>&nbsp;(${item.status} | ${item.priority})<br>`;
            html += `<b>Next Milestone:</b>&nbsp;${item.milestone}<br>`;
            html += `<b>Notes:</b>&nbsp;${item.notes}<br>`;
            html += `---<br>`;
        });
        return html;
    }
    return "";
}

// --- Main Logic ---

const db = firebase.firestore();

const els = {
    title: document.getElementById("report-title"),
    subtitle: document.getElementById("report-subtitle"),
    lastUpdated: document.getElementById("last-updated"),
    navLinks: document.getElementById("nav-links"),
    linkWeekly: document.getElementById("link-weekly"),
    linkDaily: document.getElementById("link-daily"),
    defaultMessage: document.getElementById("default-message"),
    reportBody: document.getElementById("report-body"),
    
    headerTasks: document.getElementById("header-tasks"),
    headerProjects: document.getElementById("header-projects"),
    headerActive: document.getElementById("header-active"),
    headerMeetings: document.getElementById("header-meetings"),
    headerEmails: document.getElementById("header-emails"),

    tasks: document.getElementById("content-tasks"),
    projects: document.getElementById("content-projects"),
    active: document.getElementById("content-active"),
    meetings: document.getElementById("content-meetings"),
    emails: document.getElementById("content-emails"),
    
    contMeetings: document.getElementById("container-meetings"),
    contEmails: document.getElementById("container-emails")
};

const urlParams = new URLSearchParams(window.location.search);
let reportId = null; 
let isDaily = false;

if (urlParams.has('daily')) {
    isDaily = true;
    reportId = urlParams.get('daily');
} else if (urlParams.has('report')) {
    isDaily = false;
    reportId = urlParams.get('report');
}

if (reportId) {
    els.defaultMessage.classList.add("hidden");
    els.reportBody.classList.remove("hidden");
    els.navLinks.classList.remove("hidden");
    els.linkWeekly.href = `?report=${reportId}`;
    els.linkDaily.href = `?daily=${reportId}`;
    setupView(isDaily);
    listenToFirestore(reportId, isDaily);
} else {
    els.subtitle.style.display = 'none';
    els.lastUpdated.style.display = 'none';
}

function setupView(isDaily) {
    if (isDaily) {
        els.contMeetings.classList.remove("hidden");
        els.contEmails.classList.remove("hidden");
    } else {
        els.contMeetings.classList.add("hidden");
        els.contEmails.classList.add("hidden");
    }
}

function listenToFirestore(reportId, isDaily) {
    // --- LISTENER 1: Main User Data (Tasks, Projects) ---
    // Reads from 'briefings/jdoe'
    db.collection("briefings").doc(reportId)
        .onSnapshot((doc) => {
            if (doc.exists) {
                const data = doc.data();
                
                let tasksHtml = "";
                let projectsHtml = "";
                let activeTasksHtml = "";
                
                // 1. Handle Daily Tasks
                if (data.structuredDailyTasks) {
                    tasksHtml = generateHtmlFromStructuredData(data.structuredDailyTasks, 'simpleList');
                } else {
                    tasksHtml = data.tasks || "";
                }

                // 2. Handle Projects
                if (data.structuredProjects) {
                     projectsHtml = generateHtmlFromStructuredData(data.structuredProjects, 'projectList');
                } else {
                     projectsHtml = data.projects || "";
                }

                // 3. Handle Active Tasks
                if (data.structuredActiveTasks) {
                     activeTasksHtml = generateHtmlFromStructuredData(data.structuredActiveTasks, 'projectList');
                } else {
                     activeTasksHtml = data.activeTasks || "";
                }

                // Title & Header Updates
                let updateTime = "Unknown";
if (data.lastUpdated) {
    // Handle both Firestore Timestamps AND String dates
    const dateObj = data.lastUpdated.toDate ? data.lastUpdated.toDate() : new Date(data.lastUpdated);
    
    // Format to readable text
    if (!isNaN(dateObj)) {
        updateTime = dateObj.toLocaleString('en-US', { 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric', 
            hour: '2-digit', 
            minute: '2-digit' 
        });
    } else {
        updateTime = data.lastUpdated; // Fallback if date is invalid
    }
}

                if (isDaily) {
                    els.title.textContent = "Daily Briefing";
                    els.subtitle.style.display = 'block';
                    els.subtitle.textContent = `Daily briefing for ${reportId}`;
                } else {
                    els.title.textContent = "Weekly Report";
                    els.subtitle.style.display = 'block';
                    els.subtitle.textContent = `Weekly Report for ${reportId}`;
                }
                els.lastUpdated.textContent = `Last generated: ${updateTime}`;

                els.headerTasks.textContent    = `1. Daily Tasks (${data.tasks_count || 0})`;
                els.headerProjects.textContent = `2. Active Projects (${data.projects_count || 0})`;
                els.headerActive.textContent   = `3. Active Tasks (${data.activeTasks_count || 0})`;

                els.tasks.innerHTML    = linkify(tasksHtml)       || "<i>No tasks found.</i>";
                els.projects.innerHTML = linkify(projectsHtml)    || "<i>No active projects found.</i>";
                els.active.innerHTML   = linkify(activeTasksHtml) || "<i>No active tasks found.</i>";
            
            } else {
                els.title.textContent = "Report Not Found";
                els.subtitle.textContent = `No data found for ID: ${reportId}`;
                els.lastUpdated.textContent = "Please ensure you have set up your Dashboard and Sync Agent.";
            }
        }, (error) => {
            console.error("Error main:", error);
        });

    // --- LISTENER 2: Outlook Data (Emails, Meetings) ---
    // Reads from 'briefings/jdoe_outlook' (The safe sandbox file)
    if (isDaily) {
        db.collection("briefings").doc(reportId + "_outlook")
            .onSnapshot((doc) => {
                if (doc.exists) {
                    const data = doc.data();
                    els.headerMeetings.textContent = `4. This Week's Meetings (${data.meetings_count || 0})`;
                    els.headerEmails.textContent   = `5. Unread Emails (Last 24h) (${data.emails_count || 0})`;
                    
                    els.meetings.innerHTML = linkify(data.meetings) || "<i>No meetings found.</i>";
                    els.emails.innerHTML   = linkify(data.emails)   || "<i>No unread emails.</i>";
                } else {
                    // If the outlook doc doesn't exist yet, just show empty
                    els.meetings.innerHTML = "<i>Waiting for Outlook Sync...</i>";
                    els.emails.innerHTML = "<i>Waiting for Outlook Sync...</i>";
                }
            }, (error) => {
                 console.error("Error outlook:", error);
            });
    }
}