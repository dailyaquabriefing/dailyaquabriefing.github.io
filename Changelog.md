# Changelog

All notable changes to the Daily Aqua Briefing app are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
