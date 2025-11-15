/**
 * Escapes HTML special characters in a string.
 * @param {string} str The string to escape.
 * @returns {string} The escaped string.
 */
function escapeHTML(str) {
    if (!str) return "";
    return str.replace(/[&<>"']/g, function(match) {
        return {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[match];
    });
}

/**
 * Converts URLs and emails in a plain text string to clickable HTML links.
 * This function also escapes the text to prevent XSS.
 * @param {string} plainText The raw plain text to convert.
 * @returns {string} HTML string with links.
 */
function linkify(plainText) {
    if (!plainText) return "";

    // 1. Escape the entire string to make it safe for .innerHTML
    let escapedText = escapeHTML(plainText);

    // 2. Define patterns
    // URL Pattern: Finds http, https, and www links.
    // It specifically looks for &amp; which is the escaped form of &
    const urlPattern = /(\b(https?:\/\/[-\w+&@#\/%?=~_|!:,.;&amp;]*[-\w+&@#\/%=~_|])|(\bwww\.[-\w+&@#\/%?=~_|!:,.;&amp;]*[-\w+&@#\/%=~_|]))/gi;
    
    // Email Pattern
    const emailPattern = /(\b[\w.-]+@[\w.-]+\.\w{2,4}\b)/gi;

    // 3. Replace URLs
    escapedText = escapedText.replace(urlPattern, function(match) {
        let url = match; // This is the escaped URL (e.g., "example.com?a=1&amp;b=2")
        let href = url;

        // Create the unescaped href attribute
        if (url.startsWith('www.')) {
            href = 'http://' + url.replace(/&amp;/g, '&');
        } else {
            href = url.replace(/&amp;/g, '&');
        }
        
        // The link text 'url' is already escaped. The href must be unescaped.
        return '<a href="' + href.replace(/"/g, '&quot;') + '" target="_blank">' + url + '</a>';
    });

    // 4. Replace Emails
    escapedText = escapedText.replace(emailPattern, '<a href="mailto:$1">$1</a>');

    return escapedText;
}


// --- Your existing code begins below ---

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
    
    // Section Headers (NEW)
    headerTasks: document.getElementById("header-tasks"),
    headerProjects: document.getElementById("header-projects"),
    headerActive: document.getElementById("header-active"),
    headerMeetings: document.getElementById("header-meetings"),
    headerEmails: document.getElementById("header-emails"),

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

let reportId = null; // Default to null. No data will load.
let isDaily = false;

// --- 2. Determine Mode and Report ID ---
if (urlParams.has('daily')) {
    // --- DAILY MODE ---
    isDaily = true;
    reportId = urlParams.get('daily');
} else if (urlParams.has('report')) {
    // --- WEEKLY MODE ---
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
    // --- NO REPORT ID FOUND ---
    // Do nothing. The page will just show the default title and placeholder.
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

                // --- Update Section Headers with Counts (NEW) ---
                els.headerTasks.textContent    = `1. Daily Tasks (${data.tasks_count || 0})`;
                els.headerProjects.textContent = `2. Active Projects (${data.projects_count || 0})`;
                els.headerActive.textContent   = `3. Active Tasks (${data.activeTasks_count || 0})`;

                // --- Inject Content ---
                // UPDATED: Now calls linkify()
                els.tasks.innerHTML    = linkify(data.tasks)       || "<i>No tasks found.</i>";
                els.projects.innerHTML = linkify(data.projects)    || "<i>No active projects found.</i>";
                els.active.innerHTML   = linkify(data.activeTasks) || "<i>No active tasks found.</i>";
                
                if (isDaily) {
                    // --- Update Headers for Daily Sections (NEW) ---
                    els.headerMeetings.textContent = `4. This Week's Meetings (${data.meetings_count || 0})`;
                    els.headerEmails.textContent   = `5. Unread Emails (Last 24h) (${data.emails_count || 0})`;
                    
                    // --- Inject Daily Content ---
                    // UPDATED: Now calls linkify()
                    els.meetings.innerHTML = linkify(data.meetings) || "<i>No meetings found.</i>";
                    els.emails.innerHTML   = linkify(data.emails)   || "<i>No unread emails.</i>";
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
