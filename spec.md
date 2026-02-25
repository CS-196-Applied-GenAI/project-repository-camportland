# Camp Kesem Event Revenue Tracking System Spec

## Overview
A Google Sheets-based system embedded with Google Apps Script to track offline revenue contributions for Camp Kesem events. The system simplifies the current manual spreadsheet, automates revenue calculations, and generates member contribution charts ready for PDF export.

## Roles and Permissions

| Role | Permissions |
|------|-------------|
| Super Admin | Full control: create/edit/delete events, add/edit members, view all data, generate PDFs |
| Admin | View all members and event summaries; cannot edit data |
| Viewer | View only own contributions |

Role enforcement is based on Google account login, with dynamic filtering applied in Apps Script.

## Sheet Structure

1. **Members**
   - Columns: `First Name`, `Last Name`, `Kesem Name`, `Member ID`
   - Only Super Admin can edit; new members added via form

2. **Events**
   - Columns: `Event Name`, `Date`, `Total Revenue (USD)`, `Raw Attendance List`
   - Attendance list accepts repeated names to indicate multiple shifts
   - Super Admin only can add/edit/delete events via form

3. **Processed Data**
   - Condensed attendance list: each member with multiplicity of shifts
   - Calculates **Revenue per Shift = Total Revenue ÷ Total Shifts**
   - Calculates **Revenue per Person = Revenue per Shift × Number of Shifts**

4. **Member Chart**
   - Columns: `Member Name`, `Events`, `Total Revenue`
   - `Events` column lists events chronologically: `Event Name x#Shifts (MM/DD)`
   - Chart automatically updates when events or member list are changed
   - PDF-ready formatting: bold headers, alternating light row colors, club logo at top

## Attendance Input
- Input via **form**: Event Name, Date, Total Revenue, raw attendance list (names can repeat)
- Script condenses duplicates and calculates multiplicity automatically
- Weighted revenue calculation based on shifts
- Unmatched names flagged for manual member addition

## Member Management
- **Add New Member** via separate form
- Members have: `First Name`, `Last Name`, `Kesem Name`
- Attendance input matches names to member list; unmatched names flagged

## Automation & Calculations
- Automatic recalculation triggered when:
  - New event is added/edited/deleted
  - Member list is updated
- Totals and charts update live
- Revenue in USD, formatted with dollar sign, commas, two decimals

## PDF Export
- Member chart is printable/exportable as PDF
- Features:
  - Bold headers
  - Alternating light row colors
  - Club logo at the top
- PDF can be manually emailed or downloaded

## UI / UX
- Embedded in Google Sheet
- Sidebars / dialogs for adding events and members
- Buttons for: `Add Event`, `Add Member`, `Generate PDF`
- Frozen headers for scrolling
- Color-coded sections for clarity
- Filtered views based on role

## Security
- Super Admin only can edit sheets directly; tabs protected
- Apps Script enforces role-based access and filtered views
- Viewer and Admin roles cannot modify underlying data

## Event & Member Handling
- Event attendance stored as raw list with repeated names
- System condenses duplicates to calculate shifts per person
- Events in member chart sorted chronologically by date
- Editing or deleting events automatically updates totals
- Only warnings for unmatched names; no other validations required

## Scalability
- Designed for ~60 members, ~10 events
- Sheet-based system sufficient for current scale
- Focused strictly on shift-based revenue for now

## Optional Features / Future Considerations
- Could add multiple contribution types (donations, merchandise) in future
- Potential for web-app deployment or more advanced dashboards if needed

---
This spec covers all requirements for implementation in Google Sheets with Apps Script, including role-based access, attendance and revenue tracking, dynamic member charts, and PDF export.
