// Manually initialize if not using Firebase Hosting
const firebaseConfig = {
  apiKey: "AIzaSyCtFf85MUkNSsSsT6Nv8M_09Fphm2DcQOU",
  authDomain: "dailybriefing-fe7df.firebaseapp.com",
  projectId: "dailybriefing-fe7df",
  storageBucket: "dailybriefing-fe7df.firebasestorage.app",
  messagingSenderId: "676708745644",
  appId: "1:676708745644:web:b1848c5edff6f0289eba09",
  measurementId: "G-QWFFCRMTVQ"
};
// Initialize Firebase
firebase.initializeApp(firebaseConfig);
// Initialize Firebase
const db = firebase.firestore();

// --- Elements ---
const els = {
    title: document.getElementById("report-title"),
    subtitle: document.getElementById("report-subtitle"),
    lastUpdated: document.getElementById("last-updated"),
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

// User ID: Default to 'ckonkol' if not specified
const reportId = urlParams.get('report') || 'ckonkol';

// Mode: Default to 'weekly' if not specified
const mode = urlParams.get('mode') || 'weekly';

// --- 2. Setup View Logic (Daily vs Weekly) ---
const isDaily = (mode === 'daily');

if (isDaily) {
    // DAILY MODE
    els.contMeetings.classList.remove("hidden");
    els.contEmails.classList.remove("hidden");
} else {
    // WEEKLY MODE (Default)
    els.contMeetings.classList.add("hidden");
    els.contEmails.classList.add("hidden");
}

// --- 3. Listen to Firestore ---
db.collection("briefings").doc(reportId)
    .onSnapshot((doc) => {
        if (doc.exists) {
            const data = doc.data();
            const dateStr = data.dateString || "Unknown Date";
            const updateTime = new Date(data.lastUpdated).toLocaleString();

            // --- Update Headers based on Mode ---
            if (isDaily) {
                els.title.textContent = "Morning Briefing";
                els.subtitle.textContent = `Daily briefing for ${reportId} for ${dateStr}`;
            } else {
                els.title.textContent = "Weekly Report";
                els.subtitle.textContent = `Weekly Report for ${reportId} for ${dateStr}`;
            }
            els.lastUpdated.textContent = `Generated: ${updateTime}`;

            // --- Inject Content ---
            els.tasks.innerHTML = data.tasks || "No tasks.";
            els.projects.innerHTML = data.projects || "No projects.";
            els.active.innerHTML = data.activeTasks || "No active tasks.";
            
            // Only populate these if we are showing them (optional optimization)
            if (isDaily) {
                els.meetings.innerHTML = data.meetings || "No meetings.";
                els.emails.innerHTML = data.emails || "No unread emails.";
            }
        
        } else {
            els.subtitle.textContent = `No data found for ID: ${reportId}`;
            els.lastUpdated.textContent = "Offline";
        }
    }, (error) => {
        console.error("Error:", error);
        els.subtitle.textContent = "Error loading report.";
    });