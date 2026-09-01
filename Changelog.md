# Changelog

All notable changes to the Daily Aqua Briefing app are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
