// --- CONFIGURATION ---
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
let targetId = null; // Stored user ID after detection

// --- HELPER FUNCTIONS ---

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

/**
 * Helper to render structured/array content lists.
 */
const renderList = (id, items) => {
    const el = document.getElementById(id);
    // Check if items is a non-empty array
    if (!Array.isArray(items) || !items.length) { 
        el.innerHTML = "<em>No items.</em>"; 
        return; 
    }
    
    let html = '<ul style="padding-left:20px;">';
    items.forEach(item => {
        if (typeof item === 'string') {
            html += `<li>${item}</li>`;
        } else {
            // Determine Color based on Status (including new Completed status)
            let color;
            switch (item.status) {
                case 'On Track': color = 'green'; break;
                case 'Delayed': color = 'red'; break;
                case 'On-Hold': color = 'grey'; break;
                case 'Completed': color = '#800080'; break; // Purple for Completed
                default: color = 'grey';
            }

            // For objects (projects/active tasks)
            html += `<li style="margin-bottom:10px;">
                <strong>${item.name}</strong> 
                <span style="font-size:0.8em; color:${color}; border:1px solid ${color}; padding:0 4px; border-radius:4px;">${item.status}</span>
                <br><small style="color:#666">${item.notes || ''}</small>`;
            
            // Add Next Milestone on a new line with different style (small font, color #999)
            if (item.milestone) {
                html += `<br><small style="color:#999; font-size:0.8em;">Next Milestone: ${item.milestone}</small>`;
            }
            
            html += `</li>`;
        }
    });
    el.innerHTML = linkify(html + '</ul>'); // Apply linkify to the final HTML
};


// --- OUTLOOK DATA LOADER ---
function loadOutlookData(reportId) {
    const outlookDocId = reportId + "_outlook";
    
    // Listen for changes on the dedicated Outlook document
    db.collection("briefings").doc(outlookDocId).onSnapshot((doc) => {
        const headerMeetings = document.getElementById('header-meetings');
        const headerEmails = document.getElementById('header-emails');
        const meetingsContent = document.getElementById('content-meetings');
        const emailsContent = document.getElementById('content-emails');

        if (doc.exists) {
            const data = doc.data();

            // Update Headers with Counts
            headerMeetings.textContent = `4. This Week's Meetings (${data.meetings_count || 0})`;
            headerEmails.textContent   = `5. Unread Emails (Last 24h) (${data.emails_count || 0})`;
            
            // Update Content using original app.js logic (Linkify data)
            meetingsContent.innerHTML = linkify(data.meetings || "") || "<i>No meetings found.</i>";
            emailsContent.innerHTML   = linkify(data.emails || "")   || "<i>No unread emails.</i>";
        } else {
            // Show placeholders if the outlook doc doesn't exist yet
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
    // Hide overlays
    document.getElementById('default-message').classList.add('hidden');
    document.getElementById('lock-screen').classList.add('hidden');
    document.getElementById('loading-overlay').classList.add('hidden'); // HIDE LOADING OVERLAY
    
    // Show Report elements
    document.getElementById('nav-links').classList.remove('hidden');
    document.getElementById('report-body').classList.remove('hidden');
    
    // Header Info
    document.getElementById('report-subtitle').textContent = "Report: " + data.reportId;
    
    // Handle Timestamp formatting
    let updateTime = "Unknown";
    if (data.lastUpdated) {
        const dateObj = data.lastUpdated.toDate ? data.lastUpdated.toDate() : new Date(data.lastUpdated);
        if (!isNaN(dateObj)) {
            updateTime = dateObj.toLocaleString('en-US', { 
                year: 'numeric', 
                month: 'long', 
                day: 'numeric', 
                hour: '2-digit', 
                minute: '2-digit' 
            });
        } else {
            updateTime = data.lastUpdated;
        }
    }
    document.getElementById('last-updated').textContent = `Last generated: ${updateTime}`;


    // Handle Visibility of Daily Sections (Meetings & Emails)
    if (isDailyMode) {
        document.getElementById('container-meetings').classList.remove('hidden');
        document.getElementById('container-emails').classList.remove('hidden');
    } else {
        document.getElementById('container-meetings').classList.add('hidden');
        document.getElementById('container-emails').classList.add('hidden');
    }

    // --- DATA FALLBACK LOGIC ---
    const dailyTasksData = data.structuredDailyTasks || data.dailyTasks;
    const projectsData = data.structuredProjects || data.projects;
    const activeData = data.structuredActiveTasks || data.activeTasks;

    // Render Main Sections
    renderList('content-tasks', dailyTasksData || []);
    renderList('content-projects', projectsData || []);
    renderList('content-active', activeData || []);
}


// --- CORE APPLICATION LOGIC (Initialization) ---

/**
 * Function called from the main input button
 */
function viewReport() {
    const input = document.getElementById('userid-input');
    const val = input.value.trim();
    if (val) {
        window.location.href = "?daily=" + encodeURIComponent(val);
    } else {
        alert("Please enter a User ID.");
        input.focus();
    }
}

/**
 * Function called when user enters correct passcode
 */
function attemptUnlock() {
    const entered = document.getElementById('unlock-pass').value;
    if (entered === pendingData.passcode) {
        document.getElementById('lock-screen').classList.add('hidden');
        renderReport(pendingData, true); // Unlock into Daily Mode
        loadOutlookData(targetId); // Start loading Outlook data after unlock
    } else {
        document.getElementById('unlock-error').style.display = 'block';
    }
}

// --- INITIAL PAGE LOAD AND ROUTING ---

document.addEventListener('DOMContentLoaded', () => {
    // Event listeners attached to DOM elements now that the file is loaded
    document.getElementById('userid-input').addEventListener("keypress", function(event) {
        if (event.key === "Enter") viewReport();
    });
    
    const urlParams = new URLSearchParams(window.location.search);
    const dailyId = urlParams.get('daily');
    const reportId = urlParams.get('report');

    targetId = dailyId || reportId;
    const isDailyMode = !!dailyId; 

    if (targetId) {
        // Setup Links
        document.getElementById('link-weekly').href = "?report=" + encodeURIComponent(targetId);
        document.getElementById('link-daily').href = "?daily=" + encodeURIComponent(targetId);

        // Fetch Data
        db.collection('briefings').doc(targetId).get().then(doc => {
            if (doc.exists) {
                const data = doc.data();
                
                if (isDailyMode) {
                    // --- DAILY BRIEFING MODE (Protected) ---
                    document.getElementById('report-title').textContent = "Daily Briefing";
                    
                    if (data.passcode && data.passcode.trim() !== "") {
                        pendingData = data; 
                        document.getElementById('loading-overlay').classList.add('hidden');
                        document.getElementById('lock-screen').classList.remove('hidden');
                        document.getElementById('report-subtitle').textContent = "Protected";
                    } else {
                        // If passcode is EMPTY, render the report immediately
                        renderReport(data, true); 
                        loadOutlookData(targetId); // Load Outlook data immediately
                    }
                } else {
                    // --- WEEKLY REPORT MODE (Public) ---
                    document.getElementById('report-title').textContent = "Weekly Report";
                    renderReport(data, false); // false = isWeeklyMode
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
        // --- LANDING PAGE ---
        document.getElementById('report-subtitle').textContent = "Please enter ID below";
        // ONLY HIDE the loading overlay if we are on the landing page
        document.getElementById('loading-overlay').classList.add('hidden'); 
        document.getElementById('default-message').classList.remove('hidden');
    }
});