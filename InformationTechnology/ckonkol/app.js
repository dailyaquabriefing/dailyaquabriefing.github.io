const db = firebase.firestore();

// --- Elements ---
const els = {
    // Header
    title: document.getElementById("report-title"),
    subtitle: document.getElementById("report-subtitle"),
    lastUpdated: document.getElementById("last-updated"),
    
    // Navigation
    navLinks: document.getElementById("nav-links"),
    linkWeekly: document.getElementById("link-weekly"),
    linkDaily: document.getElementById("link-daily"),

    // Placeholders
    defaultMessage: document.getElementById("default-message"),
    reportBody: document.getElementById("report-body"),
    
    // Content Sections
    tasks: document.getElementById("content-tasks"),
    projects: document.getElementById("content-projects"),
    active: document.getElementById("content-active"),
    meetings: document.getElementById("content-meetings"),
    emails: document.getElementById("content-emails"),
    
    // Containers (for hiding)
    contMeetings: document.getElementById("container-meetings"),
    contEmails: document.getElementById("container-emails")
};

// --- 1. Get URL Parameters ---
const urlParams = new URLSearchParams(window.location.search);

let reportId = null;
let isDaily = false;

// --- 2. Determine Mode and Report ID ---
if (urlParams.has('daily')) {
    // --- DAILY MODE ---
    // URL is index.html?daily={reportid}
    isDaily = true;
    reportId = urlParams.get('daily');
} else if (urlParams.has('report')) {
    // --- WEEKLY MODE ---
    // URL is index.html?report={reportid}
    isDaily = false;
    reportId = urlParams.get('report');
}

// --- 3. Run Logic ONLY if a valid ID was found ---
if (reportId) {
    // A valid report was requested. Hide the default message and show the report.
    els.defaultMessage.classList.add("hidden");
    els.reportBody.classList.remove("hidden");
    
    // Show and update the nav links
    els.navLinks.classList.remove("hidden");
    els.linkWeekly.href = `?report=${reportId}`;
    els.linkDaily.href = `?daily=${reportId}`;

    // Setup the view (hide/show meetings)
    setupView(isDaily);
    
    // Start listening to data
    listenToFirestore(reportId, isDaily);
} else {
    // No valid parameters found. Do nothing.
    // The page will just show the default title and message.
    els.subtitle.style.display = 'none';
    els.lastUpdated.style.display = 'none';
}


// --- 4. Setup View Function ---
function setupView(isDaily) {
    if (isDaily) {
        // DAILY MODE: Show everything
        els.contMeetings.classList.remove("hidden");
        els.contEmails.classList.remove("hidden");
    } else {
        // WEEKLY MODE: Hide Meetings & Emails
        els.contMeetings.classList.add("hidden");
        els.contEmails.classList.add("hidden");
    }
}

// --- 5. Listen to Firestore Function ---
function listenToFirestore(reportId, isDaily) {
    db.collection("briefings").doc(reportId)
        .onSnapshot((doc) => {
            if (doc.exists) {
                const data = doc.data();
                const dateStr = data.dateString || "Unknown Date";
                const updateTime = new Date(data.lastUpdated).toLocaleString();

                // --- Update Headers based on Request ---
                if (isDaily) {
                    els.title.textContent = "Daily Briefing";
                    els.subtitle.style.display = 'block';
                    els.subtitle.textContent = `Daily briefing for ${reportId} for ${dateStr}`;
                } else {
                    els.title.textContent = "Weekly Report";
                    els.subtitle.style.display = 'block';
                    els.subtitle.textContent = `Weekly Report for ${reportId} for ${dateStr}`;
                }
                els.lastUpdated.textContent = `Last generated: ${updateTime}`;

                // --- Inject Content ---
                els.tasks.innerHTML    = data.tasks       || "<i>No tasks found.</i>";
                els.projects.innerHTML = data.projects    || "<i>No active projects found.</i>";
                els.active.innerHTML   = data.activeTasks || "<i>No active tasks found.</i>";
                
                if (isDaily) {
                    els.meetings.innerHTML = data.meetings || "<i>No meetings found.</i>";
                    els.emails.innerHTML   = data.emails   || "<i>No unread emails.</i>";
                }
            
            } else {
                els.title.textContent = "Report Not Found";
                els.subtitle.textContent = `No data found for ID: ${reportId}`;
                els.lastUpdated.textContent = "Please run the MorningReport.ahk script.";
            }
        }, (error) => {
            console.error("Error:", error);
            els.subtitle.textContent = "Error loading report. Check console.";
        });
}