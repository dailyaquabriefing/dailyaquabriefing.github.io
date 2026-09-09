// Shared site navigation bar for Daily Aqua Briefing.
// Include on any page with: <script src="nav.js" defer></script>
// Remembers the last used report ID (from ?report=, ?daily=, or the admin
// dashboard) so Briefing / Status Report links stay personalized.
(function () {
    'use strict';

    // --- Figure out the report ID ---
    var params = new URLSearchParams(location.search);
    var reportId = (params.get('report') || params.get('daily') || '').toLowerCase().trim();
    try {
        if (reportId) {
            localStorage.setItem('dab_reportId', reportId);
        } else {
            reportId = localStorage.getItem('dab_reportId') || '';
        }
    } catch (e) { /* localStorage blocked; links fall back to generic pages */ }

    var q = reportId ? '?report=' + encodeURIComponent(reportId) : '';
    var page = (location.pathname.split('/').pop() || 'index.html').toLowerCase();
    if (page === '') page = 'index.html';

    var links = [
        { href: 'index.html' + q, label: '🏠 Briefing', match: 'index.html' },
        { href: 'admin.html', label: 'Dashboard', match: 'admin.html' },
        { href: 'weekly-report.html' + q, label: 'Status Report', match: 'weekly-report.html' },
        { href: 'howtouse.html', label: 'Help', match: 'howtouse.html' },
        { href: 'users.html', label: 'Users (IT)', match: 'users.html' }
    ];

    // --- Styles (namespaced so they never clash with page CSS) ---
    var style = document.createElement('style');
    style.textContent =
        '#dab-nav{background:#005a9e;max-width:1100px;margin:0 auto 20px auto;padding:8px 14px;' +
        'border-radius:8px;display:flex;align-items:center;gap:4px;flex-wrap:wrap;box-sizing:border-box;' +
        "font-family:'Segoe UI',system-ui,sans-serif;box-shadow:0 2px 8px rgba(0,0,0,0.12);}" +
        '#dab-nav .dab-brand{color:#fff;font-weight:700;font-size:0.95rem;text-decoration:none;margin-right:auto;padding:6px 8px;}' +
        '#dab-nav a.dab-link{color:rgba(255,255,255,0.92);text-decoration:none;font-size:0.88rem;padding:6px 10px;border-radius:4px;white-space:nowrap;}' +
        '#dab-nav a.dab-link:hover{background:rgba(255,255,255,0.18);}' +
        '#dab-nav a.dab-link.dab-active{background:#fff;color:#005a9e;font-weight:600;}' +
        '@media (max-width:700px){#dab-nav{justify-content:center;}#dab-nav .dab-brand{margin-right:0;width:100%;text-align:center;}}' +
        '@media print{#dab-nav{display:none;}}';
    document.head.appendChild(style);

    // --- Build the bar ---
    var nav = document.createElement('nav');
    nav.id = 'dab-nav';

    var brand = document.createElement('a');
    brand.className = 'dab-brand';
    brand.href = 'index.html';
    brand.textContent = 'Aqua Daily Briefing';
    nav.appendChild(brand);

    links.forEach(function (l) {
        var a = document.createElement('a');
        a.className = 'dab-link' + (page === l.match ? ' dab-active' : '');
        a.href = l.href;
        a.textContent = l.label;
        nav.appendChild(a);
    });

    document.body.insertBefore(nav, document.body.firstChild);
})();
