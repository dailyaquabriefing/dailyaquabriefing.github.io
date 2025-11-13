const db = firebase.firestore();

// --- Elements ---
const els = {
    title: document.getElementById("report-title"),
    subtitle: document.getElementById("report-subtitle"),
    lastUpdated: document.getElementById("last-updated"),
    
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
const reportId = urlParams.get('report') || 'ckonkol';

// --- 2. Determine Mode (The Fix) ---
// If the user specified a 'mode' in the URL, use it.
// IF NOT: Check if a 'report' ID exists. 
//    - If yes (e.g. ?report=bob), default to 'daily'.
//    - If no (just index.html), default to 'weekly'.
let defaultMode = urlParams.has('report') ? 'daily' : 'weekly';
const mode = urlParams.get('mode') || defaultMode;
const isDaily = (mode.toLowerCase() === 'daily');

// --- 3. Setup View Logic ---
if (isDaily) {
    // DAILY MODE: Show everything
    els.contMeetings.classList.remove("hidden");
    els.contEmails.classList.remove("hidden");
} else {
    // WEEKLY MODE: Hide Meetings & Emails
    els.contMeetings.classList.add("hidden");
    els.contEmails.classList.add("hidden");
}

// --- 4. Listen to Firestore ---
db.collection("briefings").doc(reportId)
    .onSnapshot((doc) => {
        if (doc.exists) {
            const data = doc.data();
            const dateStr = data.dateString || "Unknown Date";
            const updateTime = new Date(data.lastUpdated).toLocaleString();

            // --- Update Headers based on Request ---
            if (isDaily) {
                // EXACT FORMAT REQUESTED: Daily briefing for {reportid} for {Date}
                els.title.textContent = `Daily briefing for ${reportId} for ${dateStr}`;
                els.subtitle.style.display = 'none'; // Hide subtitle since title has all info
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