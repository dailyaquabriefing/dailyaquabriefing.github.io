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

// --- DOM Elements ---
const reportContainer = document.getElementById("report-container");
const lastUpdatedEl = document.getElementById("last-updated");
const briefingTitle = document.getElementById("briefing-title"); // Assumes you add an <h1 id="briefing-title">

// --- THIS IS THE CHANGE ---
// 1. Get the report ID from the URL
const urlParams = new URLSearchParams(window.location.search);
const reportId = urlParams.get('report') || 'ckonkol'; // Get 'report' from URL, or use a default

// 2. Update the page title
if (briefingTitle) {
    briefingTitle.textContent = `${reportId}'s Morning Briefing`;
}

// 3. Listen to the correct document in the 'briefings' collection
db.collection("briefings").doc(reportId)
    .onSnapshot((doc) => {
        if (doc.exists) {
            // ... (rest of the code is the same) ...
            const data = doc.data();
            const reportHtml = data.html;
            // ... etc ...
            reportContainer.innerHTML = reportHtml;
            lastUpdatedEl.textContent = "Report generated at: " + new Date(data.lastUpdated).toLocaleString();
        
        } else {
            reportContainer.innerHTML = `No briefing found for report ID: ${reportId}.`;
            lastUpdatedEl.textContent = "Offline";
        }
    }, (error) => {
        console.error("Error fetching briefing:", error);
        reportContainer.innerHTML = "Error loading report. Check console.";
    });