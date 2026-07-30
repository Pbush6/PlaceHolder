# Calendar Month Grid + Agenda Design

**Date:** 2026-07-30  
**Status:** Approved direction (mockup A2)  
**Mockup:** `Cursor Output/PurviewTeamsPstToHtmlApp/Mockups/calendar-views-2026-07-30/A2-month-grid-agenda.html`

## Goal

Replace the Calendar report’s flat card list with a real monthly calendar that places appointments on the days they start, plus a scrollable chronological agenda on the left and a detail panel on the right. Selecting either an agenda row or a calendar chip shows the same appointment details and keeps both views in sync.

## Layout

Three columns on desktop:

1. **Agenda (≈240px, sticky, scrollable)** — chronological list for the visible month, grouped by date, showing time, title, and location/folder.
2. **Month grid (flexible, widest)** — Sunday–Saturday month board with colored chips per appointment; “+N more” when a day overflows.
3. **Detail panel (≈320px, sticky)** — full appointment fields for the selected item.

Narrow viewports stack: agenda, then month, then detail.

## Interaction

- Click agenda row → select matching chip, flash/scroll the day cell into view, open detail.
- Click calendar chip (or “+N more”) → select matching agenda row, open detail.
- One shared selection state; both surfaces highlight the active item.
- Month prev/next/today navigation; default month is the month containing the earliest record after filters (or today if empty).
- Toolbar filters remain: free-text search, from/to date, folder, type, all-day, recurring. Filtering updates agenda list, month chips, and visible count. Clearing filters restores the full selected set.

## Data / placement rules

- Place each appointment on its **start date** (local calendar date from existing `StartTime` / filter start fields).
- All-day items use the existing exclusive-midnight end-date normalization for filter ranges; they still appear on the start day on the grid.
- Recurring series remain **one record** (no occurrence expansion), matching current extraction.
- Preserve comprehensive detail fields already exported (subject, type, times, all-day, location, organizer, attendees, recurrence, categories, sensitivity, attachments, folder, notes, timestamps, message class, Entry ID).

## Non-goals

- No week/timeline view in this change.
- No Outlook live sync; static offline HTML only.
- No change to Contacts, Teams, or Email reports beyond shared CSS reuse if already shared safely.
- No plugin framework.

## Compatibility

- Keep typed `*_Calendar.html` / `*_Calendar.log` paths, exported counts, launcher open-in-browser behavior, HTML encoding, and offline self-contained CSS/JS.
- Version bump and packaging follow the project’s next release process after implementation verification.
