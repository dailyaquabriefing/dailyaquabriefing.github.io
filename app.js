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
let currentShowPrivate = false; // <--- ADD THIS LINE

// NEW GLOBALS FOR EXPORT
let currentReportData = null;
let currentOutlookData = null;

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

// Toggle comment visibility
window.toggleComments = function(id) {
    const el = document.getElementById(id);
    if (el) {
        el.classList.toggle('open');
        // Save Name Preference
        const nameInput = el.querySelector('.comment-input-name');
        if(nameInput && !nameInput.value) {
            nameInput.value = localStorage.getItem('commenterName') || '';
        }
    }
};

// Post a new comment
window.postComment = function(listType, itemIndex, uniqueId) {
    const container = document.getElementById('comments-' + uniqueId);
    const nameVal = container.querySelector('.comment-input-name').value.trim();
    const textVal = container.querySelector('.comment-input-text').value.trim();

    if (!nameVal || !textVal) {
        alert("Please enter both your Name and a Comment.");
        return;
    }

    // Save name for next time
    localStorage.setItem('commenterName', nameVal);

    // Fetch current data 
    const docRef = db.collection('briefings').doc(targetId); 

    docRef.get().then(doc => {
        if (!doc.exists) return;
        const data = doc.data();
        
        // Determine which array to update and which Container to re-render
        let listKey = '';
        let listData = [];
        let containerId = '';
        
        if (listType === 'daily') { 
            listKey = 'structuredDailyTasks'; 
            listData = data.structuredDailyTasks || data.dailyTasks; 
            containerId = 'content-tasks';
        } else if (listType === 'project') { 
            listKey = 'structuredProjects'; 
            listData = data.structuredProjects || data.projects; 
            containerId = 'content-projects';
        } else if (listType === 'active') { 
            listKey = 'structuredActiveTasks'; 
            listData = data.structuredActiveTasks || data.activeTasks; 
            containerId = 'content-active';
        }

        if (!listData[itemIndex]) return;

        // Create Comment Object
        const newComment = {
            author: nameVal,
            text: textVal,
            timestamp: new Date().toISOString()
        };

        // Initialize array if it doesn't exist
        if (!listData[itemIndex].publicComments) {
            listData[itemIndex].publicComments = [];
        }

        // Add to local memory immediately
        listData[itemIndex].publicComments.push(newComment);

        // Save back to Firestore
        docRef.update({
            [listKey]: listData
        }).then(() => {
            // 1. Re-render the list immediately using local data
            // We use the global 'currentShowPrivate' variable we set in Step 2
            renderList(containerId, listData, currentShowPrivate);

            // 2. The re-render will close the box, so we must force it open again
            const newContainer = document.getElementById('comments-' + uniqueId);
            if (newContainer) {
                newContainer.classList.add('open');
                
                // 3. Clear the text box (but keep the name)
                newContainer.querySelector('.comment-input-text').value = '';
                newContainer.querySelector('.comment-input-name').value = nameVal;
            }
        });
    });
};

const renderList = (id, items, showPrivate = false) => {
    const el = document.getElementById(id);
    const headerEl = document.getElementById('header-' + id.replace('content-', ''));
    
    // Determine List Type for saving logic
    let listType = 'daily';
    if (id === 'content-projects') listType = 'project';
    if (id === 'content-active') listType = 'active';

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
    
    items.forEach((item, index) => {
        let name, notes = '', status, milestone = '', tester = '', startDate = '', endDate = '', lastUpdated = '', goal = '', attachment = '', itComments = '', publicComments = [];
        
        if (typeof item === 'object' && item !== null && item.name) {
            name = item.name;
            notes = item.notes || '';
            status = item.status;
            milestone = item.milestone;
            tester = item.tester;
            startDate = item.startDate;
            endDate = item.endDate;
            lastUpdated = item.lastUpdated;
            goal = item.goal;
            attachment = item.attachment;
            itComments = item.itComments;
            publicComments = item.publicComments || []; // Get Comments

        } else if (typeof item === 'string') {
            name = item;
        } else {
            return; 
        }

        const safeName = linkify(name);
        const safeNotes = linkify(notes);
        const uniqueId = `${listType}-${index}`;

        // Determine if it's a "Complex" item
        const isComplexItem = (status || tester || startDate || endDate || goal || attachment || itComments);

        // Build Extra Meta Data HTML
        let metaHtml = '';
        let dateParts = [];
        const fmtDate = (d) => d; 

        if (startDate) dateParts.push(`Start: ${fmtDate(startDate)}`);
        if (lastUpdated) dateParts.push(`Updated: ${fmtDate(lastUpdated)}`);
        if (endDate) dateParts.push(`End: ${fmtDate(endDate)}`);
        
        if (dateParts.length > 0) {
            metaHtml += `<div style="font-size:0.8em; color:#777; margin-top:2px;">\uD83D\uDCC5 ${dateParts.join(' <span style="color:#ccc;">|</span> ')}</div>`;
        }

        if (tester) {
            metaHtml += `<div style="margin-top:2px;"><span style="font-size:0.75em; background:#eef; color:#336; padding:1px 6px; border-radius:4px; border:1px solid #dde;">\uD83D\uDC64 Tester: ${tester}</span></div>`;
        }
        
        // Build Goal and Attachment HTML
        let goalHtml = '';
        if (goal) {
            goalHtml = `<div class="item-goal">🎯 <strong>Goal:</strong> ${linkify(goal)}</div>`;
        }
        
        let attachmentHtml = '';
                if (attachment) {
                    // Changed text to "View Resource"
                    attachmentHtml = `<div class="item-attachment">🔗 <a href="${attachment}" target="_blank">View Resource</a></div>`;
                }

        // IT Comments HTML (Only if showPrivate is true)
        let itCommentsHtml = '';
        if (itComments && showPrivate) {
            itCommentsHtml = `<div class="item-it-comment">🔒 <strong>IT Only:</strong> ${linkify(itComments)}</div>`;
        }

        // --- BUILD COMMENTS HTML ---
        const commentCount = publicComments.length;
        const commentLabel = commentCount > 0 ? `💬 View/Add Comments (${commentCount})` : `💬 Add Question/Comment`;
        
        let commentsListHtml = '';
        publicComments.forEach(c => {
            let timeStr = '';
            if(c.timestamp) {
                const d = new Date(c.timestamp);
                timeStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            }
            commentsListHtml += `
                <div class="comment-bubble">
                    <div class="comment-header">
                        <span class="comment-author">${linkify(c.author)}</span>
                        <span>${timeStr}</span>
                    </div>
                    <div>${linkify(c.text)}</div>
                </div>
            `;
        });

        const commentsSectionHtml = `
            <div class="comments-section">
                <button class="comment-toggle" onclick="toggleComments('comments-${uniqueId}')">${commentLabel}</button>
                <div id="comments-${uniqueId}" class="comments-container">
                    ${commentsListHtml}
                    <div class="comment-form">
                        <input type="text" class="comment-input-name" placeholder="Your Name" maxlength="20">
                        <input type="text" class="comment-input-text" placeholder="Type a question or comment...">
                        <button class="btn-post" onclick="postComment('${listType}', ${index}, '${uniqueId}')">Post</button>
                    </div>
                </div>
            </div>
        `;

        // --- RENDER ITEM ---
        if (typeof item === 'object') {
            // Complex Item or Object-based Daily Task
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

            const statusBadge = status ? `<span style="font-size:0.8em; color:${color}; border:1px solid ${color}; padding:0 4px; border-radius:4px; margin-left:5px;">${status}</span>` : '';
            const milestoneHtml = milestone ? `<small style="color:#999; font-size:0.8em; display:block;">\uD83C\uDFC1 Next Milestone: ${linkify(milestone)}</small>` : '';

            html += `<li style="margin-bottom:15px;">
                <div style="margin-bottom:2px;">
                    <strong>${safeName}</strong>${statusBadge}
                </div>
                <small style="color:#666; display:block; margin-bottom:2px;">${safeNotes}</small>
                ${goalHtml}
                ${attachmentHtml}
                ${itCommentsHtml}
                ${milestoneHtml}
                ${metaHtml}
                ${commentsSectionHtml}
            </li>`;
        } else {
             // String item (Legacy Daily Task) - minimal rendering, no comments
             html += `<li style="margin-bottom:5px;"><strong>${safeName}</strong></li>`;
        }
    });
    
    el.innerHTML = html + '</ol>';
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
            currentOutlookData = data; // NEW: Capture for export
            headerMeetings.textContent = `4. Meetings (${data.meetings_count || 0})`;
            headerEmails.textContent   = `5. Emails (${data.emails_count || 0})`;
            // Keep linkify here as these are raw text blobs from Outlook
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
    
    // NEW: Capture data for export
        currentReportData = data;
    
    // Check if the report is protected (has a passcode).
    const hasPasscode = (data.passcode && data.passcode.trim() !== "");

    // STRICT VISIBILITY RULE:
    // 1. Must be in Daily Mode (Secure Tab).
    // 2. Must have a Passcode configured.
    // If we are active in renderReport in Daily Mode, it implies the passcode was entered successfully.
    // If we are in Weekly Mode (isDailyMode=false), showPrivate forces to false.
    const showPrivate = isDailyMode && hasPasscode;
    currentShowPrivate = showPrivate; // <--- ADD THIS LINE

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

    // Pass the restricted visibility flag
    renderList('content-tasks', dailyTasksData || [], showPrivate);
    renderList('content-projects', projectsData || [], showPrivate);
    renderList('content-active', activeData || [], showPrivate);
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

function cancelUnlock() {
    // Redirects to the base URL, effectively clearing the query parameters 
    // and exiting the "Protected" mode back to the default welcome screen.
    window.location.href = window.location.pathname;
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
    // 1. Enter Key for Main User ID Input
    document.getElementById('userid-input').addEventListener("keypress", function(event) {
        if (event.key === "Enter") viewReport();
    });

    // 2. NEW: Enter Key for Passcode Input
    document.getElementById('unlock-pass').addEventListener("keypress", function(event) {
        if (event.key === "Enter") attemptUnlock();
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
                    // Check logic for protection
                    if (data.passcode && data.passcode.trim() !== "") {
                        pendingData = data; 
                        document.getElementById('loading-overlay').classList.add('hidden');
                        document.getElementById('lock-screen').classList.remove('hidden');
                        document.getElementById('report-subtitle').textContent = "Protected";
                        
                        // 3. NEW: Auto-focus the passcode field
                        setTimeout(() => {
                            document.getElementById('unlock-pass').focus();
                        }, 100);

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
// --- EXPORT FUNCTION ---
function exportReportToExcel() {
    if (!currentReportData) {
        alert("No data loaded to export.");
        return;
    }

    // Helper: Clean data based on privacy settings
    const formatForExcel = (list) => {
        if (!Array.isArray(list)) return [];
        return list.map(item => {
            // Flatten comments
            let commentsStr = "";
            if (item.publicComments && item.publicComments.length > 0) {
                commentsStr = item.publicComments.map(c => `[${c.author}]: ${c.text}`).join(" | ");
            }

            let row = {
                Name: item.name,
                Status: item.status || "",
                Priority: item.priority || "",
                Goal: item.goal || "",
                Milestone: item.milestone || "",
                Start: item.startDate || "",
                End: item.endDate || "",
                Updated: item.lastUpdated || "",
                Tester: item.tester || "",
                Notes: item.notes || "",
                Attachment: item.attachment || "",
                Public_Comments: commentsStr
            };

            // SECURITY CHECK: Only export IT comments if we are currently in Private Mode
            if (currentShowPrivate) {
                row.IT_Private_Comments = item.itComments || "";
            }

            return row;
        });
    };

    const wb = XLSX.utils.book_new();

    // 1. Task Sheets
    const daily = formatForExcel(currentReportData.structuredDailyTasks || currentReportData.dailyTasks);
    const projects = formatForExcel(currentReportData.structuredProjects || currentReportData.projects);
    const active = formatForExcel(currentReportData.structuredActiveTasks || currentReportData.activeTasks);

    if(daily.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(daily), "Daily Tasks");
    if(projects.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(projects), "Projects");
    if(active.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(active), "Active Tasks");

    // 2. Outlook Sheet (Only if Private Mode is Active AND data exists)
    if (currentShowPrivate && currentOutlookData) {
        const outlookRows = [];
        
        // Add Meetings
        if(currentOutlookData.meetings) {
            // Outlook data often comes as a big HTML blob or text block.
            // We'll just dump the text content into a cell.
            outlookRows.push({ Type: "MEETINGS", Content: currentOutlookData.meetings.replace(/<[^>]*>?/gm, '') });
        }
        
        // Add Emails
        if(currentOutlookData.emails) {
            outlookRows.push({ Type: "EMAILS", Content: currentOutlookData.emails.replace(/<[^>]*>?/gm, '') });
        }

        if (outlookRows.length > 0) {
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(outlookRows), "Outlook Data");
        }
    }

    // 3. Download
    const dateStr = new Date().toISOString().split('T')[0];
    const mode = currentShowPrivate ? "Private_Briefing" : "Public_Report";
    XLSX.writeFile(wb, `${mode}_${targetId}_${dateStr}.xlsx`);
}
