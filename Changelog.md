# Changelog

All notable changes to the Daily Aqua Briefing app are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.22.9] - 2026-09-16

### Changed
- Projects panel now also lists **Future** projects, so pipeline items can be shown/hidden on the report (they appear in the "Up next (Future)" line and the Future status chip). Completed and Maintenance remain excluded from the picker.

## [1.22.8] - 2026-09-16

### Changed
- Status of Projects chips now count only the projects **actually in view** — after status/priority/focus filters and Projects-panel hiding — instead of the full portfolio. Combined with the existing section-checkbox gating, the band now exactly mirrors the report below it (web and PDF).

## [1.22.7] - 2026-09-16

### Changed
- Save PDF: **projects no longer split across pages.** Each project (band + roadblock note + milestone rows) renders as its own block with page-break avoidance — if it doesn't fit in the space left on a page, the whole project moves to the next page. (A project taller than a full page still flows, as it must.)
- PDF tables also gained a proper bottom margin so rows can no longer run into the page footer, and single rows never split mid-row anywhere in the document.

## [1.22.6] - 2026-09-16

### Removed
- Save PDF: the "Sections hidden from this report: …" line at the top of the report — hidden sections are simply absent now, with no callout. (The on-screen report never showed this line; the "Filtered view" and staleness notices are unchanged.)

## [1.22.5] - 2026-09-16

### Changed
- Save PDF: projects now genuinely pop. Milestone status letters render as small colored text instead of full-cell color blocks (the status column was a solid stripe drowning out project rows), row striping lightened to near-white (it was almost the same blue as the project bands), project bands are taller (more padding, 10.5pt name), and a **blue rule is drawn across the top of every project row** in Key Projects and On Hold.

## [1.22.4] - 2026-09-16

### Changed
- Status Report: **project rows now pop out as blue bands** so management can scan projects at a glance — steel-blue background, blue rule above each project, larger bold project name in brand blue, bold Assigned To. Milestone sub-rows stay light so the hierarchy reads instantly. Applied to Key Projects and On Hold on the web report and in the Save PDF output (roadblock note rows get a soft amber fill in the PDF too).

## [1.22.3] - 2026-09-11

### Fixed
- Projects panel ordering appeared to do nothing unless the Order dropdown was manually set to "Custom (Projects panel)". Using ▲/▼ in the panel now **switches to Custom order automatically** and rebuilds the report immediately, and the panel header explains it.

## [1.22.2] - 2026-09-11

### Changed
- Projects panel now lists only **in-flight projects** (active statuses and On-Hold). Completed, Maintenance, and Future projects no longer clutter the picker — they weren't part of the curated Key Projects / On Hold ordering anyway.

## [1.22.1] - 2026-09-11

### Added
- Projects panel: each row now shows the project's **Priority** as a colored chip (High red, Medium amber, Low green) next to the owner/status — easier to decide what to show and how to order for the committee.

## [1.22.0] - 2026-09-11

### Added
- Status Report: **📂 Projects… panel** — lists every loaded project (with owner and status). Untick to hide a project from the report; ▲/▼ set a custom order. A legend notes how many projects are hidden. Works per-user and merged (selections are keyed by owner + project name).
- New **"Order: Custom (Projects panel)"** option in the Order dropdown — Key Projects and On Hold follow the panel's arrangement.
- **Saving:** panel selections and the order choice persist locally, and **💾 Save team now stores them with the team** (alongside members, merge setting, and sections) — opening the team restores the fully curated view. Older teams keep working; they just don't carry these until re-saved.

## [1.21.1] - 2026-09-11

### Fixed
- Status Report: **all active projects now get full rows** — the old top-10 "Key Projects" cut is removed. In merged multi-user view the cut was hiding most of the team's projects, relegating them to a names-only "Also active" line. The Key Projects section title now shows the count (e.g. "Key Projects (14)"), web and PDF.

## [1.21.0] - 2026-09-11

### Added
- Status Report: priority filter gains a **"High + Medium"** option, so steering-committee views can show both without re-entering projects at a different priority.
- Status Report: new **Order** dropdown — "Needs attention first" (the existing attention/priority/date sort) or **"My dashboard order"**, which renders Key Projects, On Hold, and Active Tasks in the exact order arranged with the Up/Down buttons on the Dashboard. The choice is remembered between visits.

### Changed
- Dashboard no longer auto-sorts projects/active tasks by status every time an item is saved — manual Up/Down ordering now sticks (this is what makes "My dashboard order" usable). The "↕ Sort by Status" button still sorts on demand.
- The "Filtered view" banner shows combined priorities as "Priority: High + Medium".

## [1.20.0] - 2026-09-11

### Added
- **Explicit 🚧 Roadblock flag on milestones.** Each milestone row in the admin edit modal now has a Roadblock checkbox with a reason field (mutually exclusive with In Progress / On-Hold). No more relying on magic keywords ("waiting", "pending"…) in notes to get a roadblock onto the report.
- Everywhere it flows: the Status Report's "Roadblock:" note row and attention-first sorting now use flagged milestones first (keyword detection remains as fallback); milestone rows show a red "🚧 Roadblock: reason" and a Y code (web and PDF); the viewer shows a 🚧 icon with the reason; the Excel export writes a `[Roadblock]` state with its reason, and the report's `.xlsx` fallback parses it back.

### Changed
- Milestone legend now reads "Y = roadblock or on hold (reason shown) or past its date".

## [1.19.3] - 2026-09-11

### Changed
- `commit.txt` workflow notes corrected: this repo pushes to **GitHub** (was mislabeled "Gitea", which is used by other internal projects).

## [1.19.2] - 2026-09-11

### Added
- Status Report: the yellow "Filtered view — …" banner now carries its own **Clear ×** button that resets all filters (status, priority, focus, completed window) exactly like the toolbar's Clear — no need to open ⚙ Options to unfilter.

### Changed
- UI buttons inside the report are excluded from "Copy as email text", the downloadable web report, and print/PDF output.

## [1.19.1] - 2026-09-11

### Changed
- Status Report: the **Status of Projects chip band now follows the section checkboxes**. Unchecking Recently Completed hides the Completed chip, Maintenance hides the Maintenance chip, On Hold hides the On-Hold chip, and unchecking Key Projects hides the active/Future/unknown status chips. Applies to the web report and the PDF; if every relevant section is off, the whole band auto-hides.

## [1.19.0] - 2026-09-11

### Added
- Saved teams now store the **"Merge into one report" setting and the section checkboxes** along with the member list. Opening a team (chip or `?team=` link) restores all three, so a team like "Steering Committee" always opens merged with the same sections.
- With a team open, **💾 Save team prefills the team's name**, so tweaking members or sections and re-saving updates the team in place instead of creating a duplicate. Connecting fresh from the connect box clears the prefill.
- Teams saved before this release still work — they just don't carry merge/section settings until re-saved.

## [1.18.3] - 2026-09-11

### Removed
- Reverted 1.18.0's "load all users by default": opening `weekly-report.html` with no parameters shows the connect screen again (last-used ID prefilled). Personal links (`?report=<id>`), team links (`?team=<name>`), the Add-user dropdown, and the saved-team chips are unchanged.

## [1.18.2] - 2026-09-11

### Added
- Status Report: **⤢ Full screen** button — presents the report itself full screen (browser Fullscreen API; Esc exits, scrolling works). Enabled once a report is loaded.

### Changed
- Merged team header now reads "Team: Chuck Konkol, Carmen Shields · Information Technology" — the department name(s) replace the "· Aqua-Aerobic Systems, Inc." suffix on that row (single-user headers keep the company name). Web and PDF.

### Removed
- The "Included in this report (N): …" banner from 1.18.1 (web and PDF) — redundant now that the Team row lists everyone.

## [1.18.1] - 2026-09-11

### Added
- Status Report: with two or more users loaded, the report now opens with an **"Included in this report (N): Chuck Konkol (Information Technology) · Carmen Shields (…)"** banner listing every selected user with their department — on the web report, the downloadable web page, the copied email text, and page 1 of the PDF.

### Changed
- Merged team view's header now uses display names ("Team: Chuck Konkol, Carmen Shields") instead of user IDs.

## [1.18.0] - 2026-09-11

### Added
- Status Report: **⚙ Options toggle** — filters, user chips (with Add user / Merge / Save team), and section checkboxes are now in a collapsible area that starts collapsed, leaving a single tidy top row: title, Copy as email text, Download web report, Save PDF, Full width, Options.
- Status Report: **⛶ Full width toggle** — stretches the report across the whole screen instead of the centered 860px column; choice is remembered between visits.

### Changed
- Status Report now loads **all users by default**: opening `weekly-report.html` with no `?report=`/`?team=` parameter pulls everyone from the user directory. Personal links from the nav bar (`?report=<id>`) and team links (`?team=<name>`) still load just that user or team.

## [1.17.2] - 2026-09-11

### Added
- Home page welcome box now shows **saved-team chips** ("Or open a team status report: 👥 IT (3) …") read from the shared `briefings/_teams` doc. Each chip opens `weekly-report.html?team=<name>` — same teams saved from the Status Report toolbar. The row hides itself when no teams exist.

## [1.17.1] - 2026-09-11

### Added
- Users (IT) page: **Edit Profile** button (admin only) — set any user's Display Name and Department on their behalf. The new admin-only `setProfile` function action (deployed) updates their `user_prefs`, their briefing document (report header), and the shared directory in one call; if the user hasn't linked a Network ID yet, the values are saved to their account and applied once they link.
- The users list now returns each user's display name so the Edit Profile prompts prefill with current values.

### Changed
- Dashboard Settings save now writes the display name to `user_prefs` as well as the briefing doc, keeping the Users page prefill complete.

## [1.17.0] - 2026-09-11

### Added
- **Profile (Display Name + Department)** on the Dashboard Settings tab. The Status Report header (web and PDF) now shows "Chuck Konkol · Information Technology · Aqua-Aerobic Systems, Inc." from these fields instead of the hardcoded IT wording and user-ID fallbacks. The viewer subtitle shows them too.
- **Company branding:** the Aqua-Aerobic logo now appears in the shared nav bar, the Status Report web header, and the PDF header.
- **Department-neutral statuses:** Planning, In Progress, and Review added everywhere (admin dropdown/badges/sorting, viewer, analytics, Excel export colors, Status Report letter codes: Planning→N/S, In Progress→G, Review→T).
- **Staleness warnings:** report sections (and PDF) show a red "Last updated N days ago" notice when a briefing is more than 7 days old.
- **Saved teams:** load a group of users, click "💾 Save team", and it becomes a one-click chip on the connect screen (shared with everyone). `weekly-report.html?team=<name>` opens a team directly. Stored in the `briefings/_teams` doc.
- **User directory (`briefings/_directory`):** maintained by the Dashboard on login/settings-save and cleaned up on user deletion. The Status Report's "Add user…" dropdown now reads it (fast), is **grouped by department**, and shows display names; falls back to a collection scan if empty.
- **Department rollup in merged reports:** with two or more departments loaded, the Completion % box adds a per-department line (e.g. "Engineering 55% (5/9) · IT 67% (24/36)"), web and PDF.
- **Delegated admin roles:** admins can mark a user as **department lead** on the Users page. Leads can list users, create accounts, and reset passwords; only admins can disable, delete, or change roles. Users page shows Department and Role columns and hides admin-only buttons from leads.
- **Unsaved-changes guard:** the Dashboard edit modal warns before discarding edits on Cancel.

### Changed
- Empty sections now auto-hide on the Status Report (Completion %/Status band when a user has no projects; Daily Workload when there are no daily tasks) — new users' sections look intentional in team reports.
- Help page documents all of the above.

### Removed
- **`updateBriefing` Cloud Function deleted** — it was an unauthenticated endpoint left over from the retired desktop agent that could overwrite any briefing. All writes now go through the authenticated app or the admin-gated `manageUsers` function. (Its structured-restore capability was used one last time to seed the `_teams`/`_directory` storage docs.)

## [1.16.2] - 2026-09-10

### Fixed
- **Incident:** deleting a test user that was linked to reportId `ckonkol` deleted the real ckonkol briefing (the delete removes whatever briefing the deleted account is linked to). Recovery: the full document was captured from Firestore's ~1-hour version retention via a point-in-time read (snapshot from 2 minutes before the wipe, zero data loss) and saved to `C:\Users\CKonkol\Projects\Dailybriefing\backups\`; data restored to Firestore.
- `manageUsers` delete now has two server-side guards (deployed): briefing data is **kept** if any other account is still linked to the same report ID, and the calling admin's own briefing can never be deleted. The function reports `dataDeleted`/`dataNote` so the Users page shows what actually happened.
- `updateBriefing` now passes through `structuredDailyTasks`/`structuredProjects`/`structuredActiveTasks`/`passcode`, enabling full-document restores from a snapshot or export.

### Changed
- Users (IT) page: delete confirm and result messages reflect the guards ("briefing data kept — another account is still linked…").
- Admin dashboard: linking your Network ID now warns (with item count) if a briefing for that ID already exists with data — preventing a test account from silently claiming a real user's briefing, which is how the incident started.

### Security/DR notes
- Recommend enabling Firestore **Point-in-Time Recovery** (7-day window) and scheduled backups; today's recovery only worked because it was caught within the ~1-hour version retention.
- `updateBriefing` remains an unauthenticated endpoint (pre-existing, from the retired desktop agent) — worth locking down or removing in the future.

## [1.16.1] - 2026-09-10

### Fixed
- User deletion now **actually** removes briefing data. The 1.16.0 approach deleted from the browser, but the Firestore security rules deny client-side deletes on `briefings/*` (verified with a direct REST test — the login was removed while the data delete was silently rejected). The cleanup moved into the `manageUsers` Cloud Function, where the Admin SDK bypasses rules: delete now removes the Auth login, the `user_prefs` mapping, `briefings/<networkid>`, and `briefings/<networkid>_outlook` in one call. Function updated and deployed to Firebase (project dailybriefing-fe7df) on 2026-09-10; source lives in `C:\Users\CKonkol\Projects\Dailybriefing\functions\index.js`.
- Users (IT) page simplified accordingly: no more client-side Firestore delete (and the Firestore SDK include it needed was removed); the success message reports whether briefing data was removed based on the function's response.

## [1.16.0] - 2026-09-10

### Changed
- Users (IT) page: **deleting a user now also deletes their briefing data.** Previously Delete only removed the Firebase Auth login — the `briefings/<networkid>` document was untouched by design (the old confirm even said "their briefing data stays in the database"), so deleted users kept a working status report and still appeared in the Status Report's "Add user…" dropdown. Delete now removes the briefing document and the legacy `<id>_outlook` sync document after the login is deleted, and the confirm dialog states exactly what will be removed. If the user never linked a Network ID, only the login is deleted (nothing else exists).
- If the data cleanup fails after the login was already deleted, the page reports it clearly instead of pretending everything was removed.

## [1.15.1] - 2026-09-10

### Fixed
- New users can now open "View My Briefing" and their Status Report right after first login, before adding any items. Cause: the `briefings/<id>` document was only created when the first item was saved, so the viewer showed "ID not found" and the status report errored. The Dashboard now writes an empty briefing document automatically the first time it loads for an ID with no data.
- Friendlier messages when a briefing genuinely doesn't exist: the viewer says 'No briefing found for "<id>". New user? Sign in to the Dashboard once to set it up.' and the status report's connect error explains the same (welcome email steps 1-3).

## [1.15.0] - 2026-09-10

### Added
- Admin dashboard Settings tab: **Change Password** section. Users enter their current password and the new one twice; the page re-authenticates against Firebase Auth and updates the password, with clear inline errors (wrong current password, too short, mismatch, too many attempts). Intended for replacing the default password from IT after first login — IT can still reset passwords from the Users page.

### Changed
- Help page: Settings Tab section documents Change Password; Step 1 recommends changing the default password right after first sign-in.

## [1.14.4] - 2026-09-10

### Changed
- Users (IT) page: welcome email's "Here is link to your status report:" line now adds "(click on only after completing steps 1-3 above)".

## [1.14.3] - 2026-09-10

### Changed
- Users (IT) page: welcome email reworded — the example link now reads "Here is link to what a working status report looks like:" (admin's report, defaults to ckonkol), and a new closing block "Here is link to your status report:" gives the new user their own `weekly-report.html?report=<userid>` link.

## [1.14.2] - 2026-09-10

### Fixed
- Save PDF: milestone rows no longer show garbage characters (`'`, `%æ`) with stretched letter-spacing and clipped names. Cause: the ✓/◦ markers aren't in the PDF font's character set (jsPDF built-in Helvetica), which also broke its text-width math. Milestones now use a safe bullet (•); done/held state still shows via the C/G/Y/N-S code chip.
- Long milestone names and hold reasons now wrap inside their cell instead of bleeding into neighboring columns (explicit line-break overflow on all PDF tables).

### Added
- Save PDF shows what was selected: a "Sections hidden from this report: …" line appears under the header whenever any section toggles are off (alongside the existing "Filtered view" line).

### Changed
- More professional PDF layout: blue rule under the report header, and a proper footer on every page — rule, company line, "Generated from…" line, and **Page X of Y** — replacing the single footer block that only appeared after the last table.

## [1.14.1] - 2026-09-09

### Changed
- Help page: added a "Minimal usage suggestion" callout — keep **Active Projects** up to date at a minimum, since they feed the Weekly Status Report; **Daily Tasks** and **Active Tasks** are optional. Same note added to the Status Report Builder section.
- Users (IT) page: the copy-paste welcome email now confirms the new user's actual Aqua Network ID (derived from their email), adds the "Minimal Usage Suggestion: Use Active Projects for the Status Report" step, and ends with a "Here is what mine looks like" example link to the sending admin's live Status Report (falls back to ckonkol).

## [1.14.0] - 2026-09-09

### Added
- Weekly report: **show/hide toggles for every section**. A new "Sections" row in the toolbar has a checkbox per section — Completion %, Status of Projects, Key Projects, On Hold, Active Tasks, Daily Workload, Maintenance, Recently Completed. Unchecking removes the section from the on-screen report, the PDF, the downloadable web report, and the copied email text alike; choices are saved and restored on the next visit.
- Section toggles work per user section and in merged team view, and combine with the status/priority/focus filters (a hidden section stays hidden regardless of filters).
- The shared status-code legend shows whenever Key Projects or Active Tasks is visible; the milestone legend follows Key Projects.

## [1.13.0] - 2026-09-09

### Fixed
- Weekly report: **Active Tasks now appear on the Status Report.** Previously the report loaded them but only showed tasks whose status was Completed (in Recently Completed) — in-flight tasks were never rendered.

### Added
- New **Active Tasks (N)** section between On Hold and Recurring Daily Workload, in both the web report and the PDF: Task | Status | Assigned To | Start | Est/Act Complete, with the same letter codes, stage labels, goal line, and roadblock notes as Key Projects. Sorted attention-first (Delayed/blocked, then priority, then nearest date). Status/priority/focus filters apply, and merged view shows owner tags per task.

## [1.12.1] - 2026-09-09

### Changed
- Weekly report: the toolbar's "Add user" textbox is now an **"Add user…" dropdown** listing everyone with a briefing (pulled from Firestore, `_outlook` sync docs excluded, alphabetical). Users already on the report are hidden from the list, and picking a name adds them immediately — no typing or + Add click.
- The main connect box now offers the same names as autocomplete suggestions while typing.
- If the user directory can't be read, the old textbox and + Add button reappear as a fallback.

## [1.12.0] - 2026-09-09

### Added
- Weekly report: **"Merge into one report"** checkbox in the toolbar (appears when two or more users are loaded). Checked, it combines every loaded user into a single report — one overall completion %, one Status of Projects band, and one Key Projects table with the attention-first sort applied across the whole team; On Hold, Recurring Daily Workload, Maintenance, and Recently Completed merge the same way. Unchecked (default), each user still renders as their own section.
- In merged view, every row carries a small **owner tag** with the User ID it came from (web tables and legends as a blue chip, PDF as `[userid]` after the name), so ownership stays clear even when a project's Assigned To is N/A or someone else.
- The merged report header reads "Team: <user1>, <user2>…", uses the newest lastUpdated date across the team, and the merge choice is remembered between visits. Filters, PDF, web download, and copy-as-email-text all follow whichever view is active.

## [1.11.0] - 2026-09-09

### Added
- Weekly report can now show **one or more user reports** on a single page, in a chosen order:
  - The connect box accepts multiple comma/space-separated User IDs (e.g. `jdoe, msmith`), and `weekly-report.html?report=jdoe,msmith` auto-connects the whole list.
  - New toolbar user bar: each loaded user shows as a numbered chip with ▲/▼ to reorder and × to remove, plus an **+ Add** input to pull in another user without losing the ones already loaded.
  - Each user renders as a full report section (their own header, completion %, status chips, tables) separated by a divider; printing puts each user on a new page.
  - "Save PDF" renders every loaded user, one per page set, with the filename joining the IDs (`WeeklyReport_jdoe+msmith_<date>.pdf`); the web download and "Copy as email text" also cover all loaded users.
  - Status/priority/focus filters apply across all loaded users; the status filter dropdown offers the union of everyone's statuses.
  - Dropping multiple `.xlsx` exports now combines them (matched by user ID) instead of replacing the loaded report, so a multi-user report can also be built fully offline.
  - The user list is remembered (localStorage), so returning to the page reconnects the same team in the same order.

### Changed
- Internal refactor: single-user `DATA`/`MODEL` globals replaced by ordered `REPORTS`/`MODELS` arrays; each report tracks its own live/export source for the header and footer lines.
- Help page: Weekly Status Report Builder section documents multi-user reports and ordering.

## [1.10.0] - 2026-09-09

### Added
- Milestones can now be rearranged: each milestone row in the admin edit modal has ▲/▼ buttons (disabled at the ends of the list). Order matters downstream — the viewer and weekly report show milestones in list order, and the "next open milestone" summary is the first unfinished one, so moving a milestone up makes it the next one reported.

### Changed
- Help page milestone bullet mentions the new reorder buttons.

## [1.9.0] - 2026-09-09

### Removed
- Setup wizard (`setup.html`) deleted — the desktop Outlook sync agent is no longer used, so the install walkthrough is obsolete. Removed the Setup link from the shared nav bar (`nav.js`), the "First Time User? Setup Here" buttons on the home page and admin login, and the setup/agent references on the Users (IT) page (including step 4 of the copy-paste welcome email, which now points to the help page).

### Changed
- Help page (`howtouse.html`) refreshed to match the current app:
  - "Two-Part System" (dashboard + Outlook agent) replaced with a "Finding Your Way Around" section describing the shared navigation bar.
  - Step 1 now covers first-time login directly on the Admin Dashboard (sign in with IT-provided credentials, link your Aqua Network ID when prompted) — no software install.
  - Milestones bullet updated for the multi-milestone list (Base/Est-Act dates, In Progress / On-Hold flags with reason).
  - Excel export section mentions the new Assigned To column and Milestones list; Outlook data references removed.
  - New "Weekly Status Report Builder" subsection under Viewing & Sharing Reports: live data connection, status/priority/focus filters, PDF download, standalone web page, copy-as-email-text, and the `.xlsx` offline fallback.
  - Meetings/Unread Emails removed from the Daily Briefing description (agent-fed data).

## [1.8.1] - 2026-09-09

### Added
- **Mike Hevey** added to the Assigned To dropdown (listed before N/A). Existing items assigned to him via "Other" now reopen with his name selected in the dropdown.

## [1.8.0] - 2026-09-09

### Added
- **Assigned To** field on Projects and Active Tasks (requested by James for the weekly report):
  - Admin edit modal: new dropdown after Status/Priority with Steve Williams, Chuck Konkol, Dan Burns, Paul Stanek, Matt Markley, Mark Milligan, Nick Schroeder, Carmen Shields, N/A, and **Other (Enter Name)** — picking Other reveals a text box for a custom name. New items default to **N/A**; choosing Other with a blank name also saves as N/A.
  - Weekly report: new **Assigned To** column in the Key Projects, On Hold, and Recently Completed tables. Works in both live Firestore mode and the `.xlsx` drop fallback (reads the export's new `Assigned_To` column). Unassigned projects show **N/A**. Carries through to the jsPDF "Print / Save PDF" export, the downloadable web report, and "Copy as email text".
  - Viewer (`app.js`): expanded item details show a green "👤 Assigned To" badge; the admin list preview shows 👤 with the assignee next to Status and Priority.
  - Excel export: new `Assigned_To` column (defaults to N/A).
- Shared site navigation bar (`nav.js`) added to every page (Briefing, Dashboard, Status Report, Setup, Help, Users (IT)). It remembers the last used report ID (from `?report=`/`?daily=` or the admin dashboard) so the Briefing and Status Report links stay personalized, highlights the current page, and hides itself when printing.

### Changed
- Merged the separately uploaded weekly-report rework (status/priority/focus filters with tappable status chips, direct PDF download via jsPDF, "Recently Completed" window) with the Assigned To column integrated into all of its render paths, including the PDF tables.
- Home page: the old inline nav links (Weekly Status Report / Admin Dashboard / How to Use) were removed in favor of the shared nav bar; Analytics and Export to Excel stay in the report toolbar.
- Items saved before this release have no assignee stored and show **N/A** on the weekly report until edited.
- How-to page documents the new Assigned To field.

## [1.7.1] - 2026-09-01

### Changed
- Weekly report: "Print / Save PDF" now suggests the filename `Weekly Report — <userid> - <date>.pdf` (date = the selected week's Monday). The page title is set during printing — which browsers use as the default PDF name — and restored afterward.

## [1.7.0] - 2026-09-01

### Added
- Milestones now have explicit **In Progress** and **On-Hold** checkboxes in the admin (mutually exclusive), plus an **On-Hold Reason** field that appears when On-Hold is checked.
- Weekly report milestone codes follow the checkboxes: C = done or both dates entered (explicit flags override), Y = on hold (reason shown in italics) or past its date, G = in progress, N/S = neither flag checked.
- Viewer milestone table shows the new states: ✅ done, ⏸️ on hold with the reason under the name, ⚠️ overdue, ⏳ in progress, ⚪ not started.
- Excel export Milestones column writes `[Done|On-Hold|In Progress|Not Started]` state tokens and a `Reason:` segment for held milestones; the weekly report's `.xlsx` fallback parses them (legacy `[Open]` rows still work).

### Changed
- Milestones saved before this release have no flags and show N/S until marked In Progress or completed.

## [1.6.0] - 2026-09-01

### Changed
- Weekly report: Key Projects now use an attention-first sort — Delayed or roadblocked projects float to the top, then priority (High → Low), then nearest Est/Act complete date, with most-recently-updated as the tiebreaker. This also changes which projects make the top-10 cut.
- Admin dashboard: new status ranking for auto-sort and "Sort by Status" — Delayed first (was second-to-last), in-flight work next (Requirement Gathering → On Track → Development → Testing → Training → Follow-Up), then Future and On-Hold, with Completed and Maintenance at the bottom. The public viewer follows the saved order automatically.
- Analytics charts/status table and the Excel export's status breakdown use the same new order (colors unchanged), as does the weekly report's "Status of Projects" chip band.

## [1.5.2] - 2026-09-01

### Changed
- Weekly report milestone codes: **C** now also applies when Base and Est/Act dates are both entered (not just the done checkbox); **N/S** shows when a milestone has no activity (no dates, not done).
- The ✔ marker and the "Act" date label follow the same complete rule, so a milestone with both dates reads as finished throughout the row.
- The **Y** (late) check now falls back to the Base date when no Est/Act date is entered, and the red highlight lands on whichever date is overdue.
- Milestone legend updated to spell out all four codes (C / G / Y / N/S).

## [1.5.1] - 2026-09-01

### Changed
- Weekly report: restored the steering committee's C/G/T/Y/N-S letter-code badges in the Key Projects, On Hold, and Completed This Week tables; the project's actual briefing status now appears as a small label under each code.
- Every milestone row now carries its own code badge: C = complete, G = open and on track, Y = open and past its estimated date.
- Status code mapping extended to all briefing statuses: On Track / Development / Training / Follow-Up → G (Y if a roadblock is detected), Testing → T, Delayed and On-Hold → Y, Future and Requirement Gathering → N/S, Completed → C.
- Legends restored/updated to explain the letter codes for both projects and milestones. The "Status of Projects" chip summary band is unchanged.

## [1.5.0] - 2026-09-01

### Added
- Weekly report: new **Status of Projects** section under the completion % — colored chips with the project count per status, using the briefing app's status order and colors.
- "Up next (Future)" line under the Key Projects table lists pipeline projects.

### Changed
- Status columns (Key Projects, On Hold, Completed This Week) now show each project's actual briefing status as a colored chip instead of template letter codes (C/G/T/Y/N-S); the letter-code legend was replaced accordingly.
- Overall completion % caption now states that maintenance and future projects are excluded from the denominator.

### Fixed
- Projects with status On Track, Training, Follow-Up, or Delayed no longer disappear from the report — Key Projects now includes all in-flight statuses (previously only Development, Testing, and Requirement Gathering were shown, hiding projects in the app's default "On Track" status).

## [1.4.0] - 2026-09-01

### Added
- Home page nav bar now includes a **📋 Weekly Status Report** link (between Analytics and Admin Dashboard). It opens `weekly-report.html?report=<userid>` in a new tab with the current briefing's user ID, so the status report auto-connects and builds immediately.

## [1.3.0] - 2026-09-01

### Added
- Weekly report: project milestones now appear as indented sub-rows under each project in the Key Projects table — ✔/◦ marker, `C` badge when complete, Base date, and Est (open) or Act (done) date. Overdue estimates show in red.
- Milestones render in both data modes: live Firestore data passes the `milestones` array through; the `.xlsx` fallback parses the export's `Milestones` column back into the same list.
- A milestone legend line appears under the status legend, only when a listed project has milestones.
- Milestone rows carry through to print/PDF, the downloadable web report, and "Copy as email text".

## [1.2.0] - 2026-09-01

### Added
- Weekly report (`weekly-report.html`) now connects directly to Firestore — enter a briefing User ID and hit **Connect**; no Excel export/upload needed.
- Auto-connect via URL: `weekly-report.html?report=<id>` builds the report on page load. The last-used ID is remembered and prefilled.
- Live data maps the new multi-milestone list into the report's `Milestone` field (next open milestone + its Est/Act or Base date), falling back to the legacy single field.

### Changed
- Drop zone replaced by a connect card; the `.xlsx` drop/browse remains as an offline fallback box below it.
- Report header shows "Live data as of <date>" (briefing `lastUpdated`) in live mode; footer states whether the report came from live data or an export.

## [1.1.0] - 2026-09-01

### Added
- Multiple milestones per project and active task. Each milestone has a description, a **Base Complete** date, an **Est/Act Complete** date, and a done checkbox.
- Admin edit modal: new **🏁 Milestones** section replaces the single "Next Milestone" text box. Rows are added, edited inline, marked complete, and removed.
- Admin project list preview shows milestone progress (e.g., "2/5 complete — Next: Go-Live (2026-10-01)").
- Viewer (`app.js`): item details render a milestone table (Milestone | Base Complete | Est/Act Complete) with a completion count. Done milestones show ✅; open milestones past their Est/Act date show ⚠️ with the date in red.
- Excel export: new `Milestones` column listing every milestone as `[Done|Open] Name — Base: date | Est/Act: date`, one per line.

### Changed
- On save, empty milestone rows are dropped and the legacy `milestone` field is set to the next open milestone (name + date) so the weekly report and older data readers keep working.
- Opening an item that only has the old single milestone text automatically converts it into the first entry in the new milestone list.

## [1.0.0] - Baseline

- Existing Daily Aqua Briefing app: Firebase-backed daily briefing and weekly report viewer (`index.html`/`app.js`), admin dashboard (`admin.html`) with daily tasks, projects, active tasks, attachments, daily operational checks, public/private comments, analytics charts, Excel export, and weekly steering-committee report (`weekly-report.html`).
