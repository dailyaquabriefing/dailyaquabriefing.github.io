// --- CONFIGURATION ---
const firebaseConfig = {
    apiKey: "AIzaSyCtFf85MUkNSsSsT6Nv8M_09Fphm2DcQOU", 
    authDomain: "dailybriefing-fe7df.firebaseapp.com",
    projectId: "dailybriefing-fe7df",
    storageBucket: "dailybriefing-fe7df.appspot.com",
    messagingSenderId: "",
    appId: ""
};

if (!firebaseConfig.apiKey) {
    alert("CRITICAL ERROR: API Key is missing in app.js!\nPlease edit the file and paste your API Key from index.html.");
}

firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// --- STATE ---
let currentReportId = null; // The User ID (e.g., jdoe)
let currentData = null;
let currentMode = 'default'; // 'default', 'daily', or 'report'

// --- INIT & URL PARSING ---
document.addEventListener('DOMContentLoaded', () => {
    // Hide loading overlay once Firebase is ready
    document.getElementById('loading-overlay').style.display = 'none';

    // 1. Get query parameters
    const params = new URLSearchParams(window.location.search);
    
    if (params.has('daily')) {
        currentReportId = params.get('daily').trim().toLowerCase();
        currentMode = 'daily';
        
    } else if (params.has('report')) {
        currentReportId = params.get('report').trim().toLowerCase();
        currentMode = 'report';
        
    } else {
        // Default View: Show input box
        document.getElementById('default-message').style.display = 'block';
        document.getElementById('report-body').classList.add('hidden');
        document.getElementById('nav-links').classList.add('hidden');
        return;
    }

    // 2. Report ID found in URL: Initialize
    if (currentReportId) {
        document.getElementById('report-subtitle').textContent = `Report for: ${currentReportId}`;
        
        // Hide default screen
        document.getElementById('default-message').style.display = 'none';
        
        // Setup navigation links
        document.getElementById('nav-links').classList.remove('hidden');
        document.getElementById('link-daily').href = `?daily=${currentReportId}`;
        document.getElementById('link-weekly').href = `?report=${currentReportId}`;

        loadData();
    }
});


// --- AUTH & DATA LOADING ---
function viewReport() {
    const input = document.getElementById('userid-input');
    const id = input.value.trim().toLowerCase();
    if (id) {
        // Redirect to the URL with the Report ID
        window.location.href = `?daily=${id}`;
    }
}

function loadData() {
    // Listen for real-time updates on the briefing document
    db.collection('briefings').doc(currentReportId).onSnapshot(doc => {
        if (doc.exists) {
            currentData = doc.data();
            
            // SECURITY CHECK: If in 'daily' mode and passcode is set
            if (currentMode === 'daily' && currentData.passcode) {
                document.getElementById('lock-screen').classList.remove('hidden');
                document.getElementById('report-body').classList.add('hidden');
            } else {
                // No passcode needed or in 'report' mode
                document.getElementById('lock-screen').classList.add('hidden');
                document.getElementById('report-body').classList.remove('hidden');
                renderData(currentData);
            }
            
        } else {
            // Document not found - show error/default view
            alert("Error: Report ID not found. Please check the ID or complete the setup process.");
            window.location.search = ''; // Go back to default view
        }
    }, error => {
        console.error("Error loading data:", error);
        alert("A critical error occurred while fetching data.");
    });
}

function attemptUnlock() {
    const enteredPass = document.getElementById('unlock-pass').value.trim();
    const storedPass = currentData.passcode;
    const errorEl = document.getElementById('unlock-error');
    
    if (enteredPass === storedPass) {
        document.getElementById('lock-screen').classList.add('hidden');
        document.getElementById('report-body').classList.remove('hidden');
        errorEl.style.display = 'none';
        renderData(currentData);
    } else {
        errorEl.style.display = 'block';
    }
}


// --- RENDERING ---
function renderData(data) {
    // 1. Render Daily Tasks (Always Visible)
    renderList('content-tasks', data.structuredDailyTasks);
    
    // 2. Render Projects (Always Visible)
    renderItems('content-projects', data.structuredProjects, 'No active projects found.');
    
    // 3. Render Active Tasks (Always Visible)
    renderItems('content-active', data.structuredActiveTasks, 'No active tasks found.');
    
    // 4. Render Meetings (Daily Mode Only)
    renderOutlookContent('container-meetings', 'content-meetings', currentMode === 'daily', data.rawMeetings, 'No meetings scheduled for this week.');

    // 5. Render Emails (Daily Mode Only)
    renderOutlookContent('container-emails', 'content-emails', currentMode === 'daily', data.rawEmails, 'No unread emails from the last 24 hours.');

    // 6. Update Last Updated/Generated Timestamp
    let updateEl = document.getElementById('report-generated'); // UPDATED ID
    
    // Prioritize agentLastRun (from AHK) over lastUpdated (from Admin Page)
    let timeSource = data.agentLastRun || data.lastUpdated; 
    
    // Change the label to "Report generated"
    if (timeSource) {
        updateEl.textContent = 'Report generated: ' + formatTime(timeSource); // UPDATED TEXT
    } else {
        updateEl.textContent = 'Report generated: Never'; // UPDATED TEXT
    }
}

function renderList(containerId, dataArray) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';
    if (!dataArray || dataArray.length === 0) {
        container.innerHTML = '<p style="color:#999;">No daily tasks entered.</p>';
        return;
    }
    const ul = document.createElement('ul');
    ul.style.listStyle = 'decimal';
    ul.style.paddingLeft = '20px';
    dataArray.forEach(item => {
        const li = document.createElement('li');
        li.textContent = item;
        ul.appendChild(li);
    });
    container.appendChild(ul);
}

function renderItems(containerId, dataArray, emptyMessage) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';
    if (!dataArray || dataArray.length === 0) {
        container.innerHTML = `<p style="color:#999; text-align:center;">${emptyMessage}</p>`;
        return;
    }
    
    dataArray.forEach(item => {
        const statusClass = 'st-' + item.status.toLowerCase().replace(' ', '-');
        const itemDiv = document.createElement('div');
        itemDiv.style.borderBottom = '1px solid #eee';
        itemDiv.style.padding = '10px 0';

        itemDiv.innerHTML = `
            <div style="font-weight:bold; font-size:1.05em; color:#0078D4;">${item.name}</div>
            <div style="font-size:0.9em; margin-top:5px; display:flex; align-items:center; gap:10px;">
                <span style="padding: 4px 8px; border-radius: 12px; font-size: 0.8rem; font-weight: bold; color: white;" class="badge ${statusClass}">${item.status}</span>
                <span style="color:#666;">Priority: ${item.priority}</span>
            </div>
            <div style="font-size:0.85em; color:#555; margin-top:8px;">
                <strong>Next Milestone:</strong> ${item.milestone || 'N/A'}
            </div>
            <div style="font-size:0.85em; color:#555; margin-top:4px;">
                <strong>Notes:</strong> ${item.notes || 'No notes.'}
            </div>
        `;
        container.appendChild(itemDiv);
    });
}

function renderOutlookContent(containerId, contentId, isVisible, rawContent, emptyMessage) {
    const container = document.getElementById(containerId);
    const content = document.getElementById(contentId);
    
    if (isVisible) {
        container.classList.remove('hidden');
        if (rawContent && rawContent.trim()) {
            // Format raw content (e.g., from AHK) as pre-formatted text
            const pre = document.createElement('pre');
            pre.textContent = rawContent.trim();
            pre.style.whiteSpace = 'pre-wrap';
            pre.style.wordWrap = 'break-word';
            pre.style.backgroundColor = '#f4f4f4';
            pre.style.padding = '10px';
            pre.style.borderRadius = '4px';
            content.innerHTML = '';
            content.appendChild(pre);
        } else {
            content.innerHTML = `<p style="color:#999; text-align:center;">${emptyMessage}</p>`;
        }
    } else {
        container.classList.add('hidden');
    }
}


// --- UTILITY ---
function formatTime(isoString) {
    try {
        const date = new Date(isoString);
        // Date options for display: e.g., November 18, 2025
        const dateOptions = { year: 'numeric', month: 'long', day: 'numeric' };
        // Time options for display: e.g., 2:19 PM
        const timeOptions = { hour: '2-digit', minute: '2-digit', hour12: true };
        
        // Use local time zone for display
        const formattedDate = date.toLocaleDateString(undefined, dateOptions);
        const formattedTime = date.toLocaleTimeString(undefined, timeOptions);
        
        return `${formattedDate} at ${formattedTime}`;
    } catch (e) {
        console.error("Error formatting date:", e);
        return 'Invalid Date';
    }
}