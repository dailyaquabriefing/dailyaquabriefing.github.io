# Changelog

All notable changes to the Daily Aqua Briefing app are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
