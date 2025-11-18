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
            html += `${i + 1}. ${task}\n`; // Using newline for linkify to catch or simple formatting
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
    // Note: reportId is now likely a Firebase User UID if created via Admin page
    db.collection("briefings").doc(reportId)
        .onSnapshot((doc) => {
            if (doc.exists) {
                const data = doc.data();
                
                // --- HYBRID HANDLING ---
                // Check if data comes from new Admin Tool (Structured) or Old AHK (Strings)
                
                let tasksHtml = "";
                let projectsHtml = "";
                let activeTasksHtml = "";
                
                // 1. Handle Daily Tasks
                if (data.structuredDailyTasks) {
                    // Data from Admin Tool
                    tasksHtml = generateHtmlFromStructuredData(data.structuredDailyTasks, 'simpleList');
                } else {
                    // Data from Legacy AHK
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
                
                // --- Outlook Data (Emails/Meetings) ---
                // These will likely still come as strings from the AHK script for now
                // until you update the AHK script to push structured data too.
                // The current AHK sends strings, so we keep that logic.
                
                const dateStr = data.dateString || "Unknown Date";
                // Use lastUpdated timestamp if available, else string
                const updateTime = data.lastUpdated ? 
                    (data.lastUpdated.toDate ? data.lastUpdated.toDate().toLocaleString() : data.lastUpdated) 
                    : "Unknown";

                if (isDaily) {
                    els.title.textContent = "Daily Briefing";
                    els.subtitle.style.display = 'block';
                    els.subtitle.textContent = `Daily briefing for ${reportId}`; // Removed date from title to avoid clutter
                } else {
                    els.title.textContent = "Weekly Report";
                    els.subtitle.style.display = 'block';
                    els.subtitle.textContent = `Weekly Report for ${reportId}`;
                }
                els.lastUpdated.textContent = `Last generated: ${updateTime}`;

                els.headerTasks.textContent    = `1. Daily Tasks (${data.tasks_count || 0})`;
                els.headerProjects.textContent = `2. Active Projects (${data.projects_count || 0})`;
                els.headerActive.textContent   = `3. Active Tasks (${data.activeTasks_count || 0})`;

                // Inject Content (Linkify handles both the new HTML generated above and old AHK strings)
                els.tasks.innerHTML    = linkify(tasksHtml)       || "<i>No tasks found.</i>";
                els.projects.innerHTML = linkify(projectsHtml)    || "<i>No active projects found.</i>";
                els.active.innerHTML   = linkify(activeTasksHtml) || "<i>No active tasks found.</i>";
                
                if (isDaily) {
                    els.headerMeetings.textContent = `4. This Week's Meetings (${data.meetings_count || 0})`;
                    els.headerEmails.textContent   = `5. Unread Emails (Last 24h) (${data.emails_count || 0})`;
                    els.meetings.innerHTML = linkify(data.meetings) || "<i>No meetings found.</i>";
                    els.emails.innerHTML   = linkify(data.emails)   || "<i>No unread emails.</i>";
                }
            
            } else {
                els.title.textContent = "Report Not Found";
                els.subtitle.textContent = `No data found for ID: ${reportId}`;
                els.lastUpdated.textContent = "If this is a new user ID, please log into the Admin Dashboard to initialize your data.";
            }
        }, (error) => {
            console.error("Error:", error);
            els.subtitle.textContent = "Error loading report. Check console.";
        });
}