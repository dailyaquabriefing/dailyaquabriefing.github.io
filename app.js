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

// --- Realtime Listener ---
// This is the core of the website. It listens for changes
// to the 'latest' document in the 'briefing' collection.
db.collection("briefing").doc("latest")
    .onSnapshot((doc) => {
        if (doc.exists) {
            // 1. Get the data
            const data = doc.data();
            const reportHtml = data.html;
            const updateTime = new Date(data.lastUpdated);

            // 2. Put the HTML from AHK directly into the page
            reportContainer.innerHTML = reportHtml;

            // 3. Update the timestamp
            lastUpdatedEl.textContent = "Report generated at: " + updateTime.toLocaleString();
        
        } else {
            // Show a message if the AHK script hasn't run yet
            reportContainer.innerHTML = "No briefing has been generated yet. Please run the MorningReport.ahk script.";
            lastUpdatedEl.textContent = "Offline";
        }
    }, (error) => {
        console.error("Error fetching briefing:", error);
        reportContainer.innerHTML = "Error loading report. Check console.";
    });
