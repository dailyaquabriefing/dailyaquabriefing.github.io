// Shared site navigation bar for Daily Aqua Briefing.
// Include on any page with: <script src="nav.js" defer></script>
// Remembers the last used report ID (from ?report=, ?daily=, or the admin
// dashboard). The remembered ID personalizes the Briefing / Status Report
// links ONLY while signed in — signed-out visitors get the plain pages.
// An ID in the current page URL still carries across (active context).
(function () {
    'use strict';

    // --- Figure out the report ID ---
    var params = new URLSearchParams(location.search);
    var urlId = (params.get('report') || params.get('daily') || '').toLowerCase().trim();
    try {
        if (urlId) localStorage.setItem('dab_reportId', urlId);
    } catch (e) { /* localStorage blocked; links fall back to generic pages */ }

    function linkQ(id) { return id ? '?report=' + encodeURIComponent(id) : ''; }

    var page = (location.pathname.split('/').pop() || 'index.html').toLowerCase();
    if (page === '') page = 'index.html';

    var links = [
        { href: 'index.html' + linkQ(urlId), label: '🏠 Briefing', match: 'index.html', idBase: 'index.html' },
        { href: 'admin.html', label: 'Dashboard', match: 'admin.html' },
        { href: 'weekly-report.html' + linkQ(urlId), label: 'Status Report', match: 'weekly-report.html', idBase: 'weekly-report.html' },
        // Team Hub lives on the intranet dashboards server (Windows sign-in),
        // fronted by the managerhub subdomain. Opens in a new tab so the
        // briefing page stays put.
        { href: 'https://managerhub.aquatools.app/', label: 'Team Hub', match: 'manager-hub.html', newTab: true },
        { href: 'howtouse.html', label: 'Help', match: 'howtouse.html' },
        { href: 'users.html', label: 'Admin', match: 'users.html' }
    ];

    // --- Styles (namespaced so they never clash with page CSS) ---
    var style = document.createElement('style');
    style.textContent =
        // Fixed 1100px width (like the Admin page) on every page: the bar
        // centers itself on the viewport, escaping width-constrained bodies
        // (Briefing 900px, Help 800px) via left:50% + translateX(-50%).
        '#dab-nav{background:#005a9e;width:min(1100px,calc(100vw - 40px));position:relative;left:50%;' +
        'transform:translateX(-50%);margin:0 0 20px 0;padding:8px 14px;' +
        'border-radius:8px;display:flex;align-items:center;gap:4px;flex-wrap:wrap;box-sizing:border-box;' +
        "font-family:'Segoe UI',system-ui,sans-serif;box-shadow:0 2px 8px rgba(0,0,0,0.12);}" +
        '#dab-nav .dab-brand{color:#fff;font-weight:700;font-size:0.95rem;text-decoration:none;margin-right:auto;padding:6px 8px;display:inline-flex;align-items:center;gap:8px;}' +
        '#dab-nav .dab-logo{height:24px;width:auto;background:#fff;border-radius:4px;padding:2px 5px;box-sizing:content-box;}' +
        '#dab-nav a.dab-link{color:rgba(255,255,255,0.92);text-decoration:none;font-size:0.88rem;padding:6px 10px;border-radius:4px;white-space:nowrap;cursor:pointer;}' +
        '#dab-nav a.dab-link:hover{background:rgba(255,255,255,0.18);}' +
        '#dab-nav a.dab-link.dab-active{background:#fff;color:#005a9e;font-weight:600;}' +
        '#dab-nav .dab-user{display:inline-flex;align-items:center;gap:4px;margin-left:6px;padding-left:10px;border-left:1px solid rgba(255,255,255,0.35);}' +
        '#dab-nav .dab-email{color:rgba(255,255,255,0.85);font-size:0.8rem;white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis;}' +
        '@media (max-width:700px){#dab-nav{justify-content:center;}#dab-nav .dab-brand{margin-right:0;width:100%;text-align:center;}}' +
        '@media print{#dab-nav{display:none;}}';
    document.head.appendChild(style);

    // --- Build the bar ---
    var nav = document.createElement('nav');
    nav.id = 'dab-nav';

    var brand = document.createElement('a');
    brand.className = 'dab-brand';
    brand.href = 'index.html';
    var logo = document.createElement('img');
    logo.src = 'aqua-aerobics-logo.png';
    logo.alt = 'Aqua-Aerobic Systems';
    logo.className = 'dab-logo';
    brand.appendChild(logo);
    brand.appendChild(document.createTextNode('Aqua Daily Briefing'));
    nav.appendChild(brand);

    var idAnchors = [];
    links.forEach(function (l) {
        var a = document.createElement('a');
        a.className = 'dab-link' + (page === l.match ? ' dab-active' : '');
        a.href = l.href;
        a.textContent = l.label;
        if (l.newTab) { a.target = '_blank'; a.rel = 'noopener'; }
        if (l.idBase) idAnchors.push({ a: a, base: l.idBase });
        nav.appendChild(a);
    });

    document.body.insertBefore(nav, document.body.firstChild);

    // --- Signed-in user + Logout ---
    // Shown only on pages that load Firebase Auth (nav.js is deferred, so the
    // page's Firebase scripts have already run) and only while signed in.
    // This is the single place username/logout live — pages don't add their own.
    function applyLinkId(id) {
        idAnchors.forEach(function (x) { x.a.href = x.base + linkQ(id); });
    }

    try {
        if (window.firebase && firebase.apps && firebase.apps.length && firebase.auth) {
            firebase.auth().onAuthStateChanged(function (user) {
                // Briefing / Status Report links: while signed in they point
                // at the account's OWN report — the linked network ID from
                // user_prefs (falling back to the remembered ID until that
                // read resolves, or on pages without Firestore). Signed out
                // falls back to the URL's ID (if any) or the plain pages.
                var id = urlId;
                if (user && !id) {
                    try { id = localStorage.getItem('dab_reportId') || ''; } catch (e) {}
                }
                applyLinkId(id);
                if (user) {
                    try {
                        if (firebase.firestore) {
                            firebase.firestore().collection('user_prefs').doc(user.uid).get().then(function (doc) {
                                var own = doc && doc.exists ? String(doc.data().reportId || '').toLowerCase() : '';
                                if (own) {
                                    applyLinkId(own);
                                    try { localStorage.setItem('dab_reportId', own); } catch (e) {}
                                }
                            }).catch(function () { /* keep the fallback links */ });
                        }
                    } catch (e) { /* no Firestore on this page */ }
                }

                var existing = document.getElementById('dab-user');
                if (existing) existing.remove();
                if (!user) return;

                var wrap = document.createElement('span');
                wrap.id = 'dab-user';
                wrap.className = 'dab-user';

                // Show just the user ID (email with the @domain stripped);
                // the full email stays available as the tooltip.
                var email = document.createElement('span');
                email.className = 'dab-email';
                email.textContent = (user.email || '').split('@')[0];
                email.title = user.email || '';
                wrap.appendChild(email);

                var out = document.createElement('a');
                out.className = 'dab-link';
                out.textContent = 'Logout';
                out.onclick = function (e) { e.preventDefault(); firebase.auth().signOut(); };
                wrap.appendChild(out);

                nav.appendChild(wrap);
            });
        }
    } catch (e) { /* page has no Firebase Auth; nav stays anonymous */ }
})();
