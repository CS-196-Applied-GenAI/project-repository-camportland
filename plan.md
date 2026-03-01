# Camp Kesem Event Revenue Tracking System
## Development Blueprint (Iterative + Testable)

---

# 1. Architecture Overview

## Platform
- Google Sheets (data store + UI surface)
- Google Apps Script (logic, roles, automation, UI dialogs)

## Core Sheets
1. Members
2. Events
3. ProcessedData
4. MemberChart

## Core Modules (Apps Script)
- RoleService
- MemberService
- EventService
- ProcessingService
- ChartService
- PdfService
- UiService

---

# 2. High-Level Development Phases

PHASE 1 – Foundation
- Create sheet structure
- Configure script project
- Add menu and UI shell
- Implement role detection

PHASE 2 – Member Management
- Add member form
- Persist to Members sheet
- Add validation
- Add unmatched-name flagging logic

PHASE 3 – Event Entry
- Add event form
- Store raw attendance list
- Validate revenue + date
- Prevent unauthorized edits

PHASE 4 – Revenue Processing Engine
- Condense duplicate names
- Calculate:
    - Total shifts
    - Revenue per shift
    - Revenue per person
- Populate ProcessedData sheet

PHASE 5 – Member Chart Builder
- Aggregate revenue per member
- Generate chronological event string
- Format for PDF
- Add alternating row colors

PHASE 6 – Automation & Triggers
- Recalculate on:
    - Event add/edit/delete
    - Member add
- Maintain sorting
- Prevent infinite loops

PHASE 7 – Role-Based Access Enforcement
- Sheet protections
- Filtered views
- Sidebar restrictions

PHASE 8 – PDF Export
- Format sheet
- Insert logo
- Export as PDF blob
- Download or email

PHASE 9 – Hardening
- Edge case testing
- Performance testing
- Cleanup and refactor

---

# 3. Iterative Build Strategy

We will build vertical slices that are:
- Safe
- Testable
- Independently verifiable
- Small enough to debug easily

---

# 4. Iteration Breakdown (Round 1)

## Iteration 1: Sheet Skeleton
- Create all 4 sheets
- Add headers
- Freeze header rows
- Add basic formatting

## Iteration 2: Basic Member Entry
- Add sidebar
- Add member form
- Save to Members sheet
- Confirm write works

## Iteration 3: Basic Event Entry
- Add event form
- Store event row
- Store raw attendance string

## Iteration 4: Attendance Condensing Logic
- Parse attendance string
- Trim whitespace
- Count duplicates
- Output condensed result in ProcessedData

## Iteration 5: Revenue Calculation
- Calculate total shifts
- Calculate revenue per shift
- Multiply per member
- Validate decimals

## Iteration 6: Member Aggregation
- Aggregate revenue across events
- Build chronological event string
- Populate MemberChart

## Iteration 7: Automation Trigger
- Recalculate on event submit
- Recalculate on member submit

## Iteration 8: Role Enforcement
- Detect email
- Assign role
- Restrict editing

## Iteration 9: PDF Export
- Apply formatting
- Export
- Validate layout

---

# 5. Further Breakdown to Safe Implementation Steps

Below is a final right-sized implementation roadmap.

Each step should:
- Be testable in isolation
- Avoid touching too many modules
- Have a clear verification method

---

# FINAL STEP PLAN

---

## STEP 1 – Create Sheets & Headers

Create sheets:
- Members
- Events
- ProcessedData
- MemberChart

Add headers only.
No logic yet.

TEST:
- All sheets exist
- Headers correct
- Frozen row applied

---

## STEP 2 – Add Custom Menu

Add:
Menu: "Kesem Revenue System"
Options:
- Add Member
- Add Event
- Generate PDF

TEST:
- Menu loads on open
- Click handlers run (even if empty)

---

## STEP 3 – Member Sidebar UI (No Logic Yet)

Build sidebar HTML form:
- First Name
- Last Name
- Kesem Name
- Submit button

TEST:
- Sidebar opens
- Submit triggers stub function

---

## STEP 4 – Persist Member to Sheet

Implement:
- Validate non-empty fields
- Generate Member ID (UUID)
- Append row

TEST:
- Member row added correctly
- ID unique
- No overwrite

---

## STEP 5 – Event Sidebar UI

Build event form:
- Event Name
- Date
- Total Revenue
- Raw Attendance List (textarea)

TEST:
- Form loads
- Submit triggers stub

---

## STEP 6 – Persist Event

Append to Events sheet:
- Store raw attendance as string
- Validate revenue numeric
- Validate date

TEST:
- Row saved correctly
- Currency formatted

---

## STEP 7 – Parse Attendance

Implement function:
parseAttendance(rawString)

Returns:
Map:
{
  "Name A": 2,
  "Name B": 1
}

TEST:
- Duplicate counting works
- Whitespace trimmed
- Empty entries ignored

---

## STEP 8 – Match Names to Members

Implement:
- Case-insensitive matching
- Flag unmatched names

Store unmatched in log column.

TEST:
- Known member matches
- Unknown member flagged

---

## STEP 9 – Revenue Per Shift Calculation

Implement:
totalShifts = sum(multiplicity)
revenuePerShift = totalRevenue / totalShifts

TEST:
- Decimal precision correct (2 places)
- Division by zero handled

---

## STEP 10 – Revenue Per Person

For each matched member:
revenuePerPerson = revenuePerShift × shifts

Write to ProcessedData sheet.

TEST:
- Math accurate
- Correct member mapping

---

## STEP 11 – Aggregate Member Totals

Build in-memory structure:
{
  memberId: {
    totalRevenue,
    events: [...]
  }
}

TEST:
- Multi-event accumulation works

---

## STEP 12 – Chronological Event Sorting

Sort events by date before building string:
"Event Name x2 (MM/DD)"

TEST:
- Correct chronological order

---

## STEP 13 – Populate MemberChart

Clear sheet.
Write:
Member Name | Events | Total Revenue

TEST:
- Totals match processed data
- Currency formatting correct

---

## STEP 14 – Alternating Row Formatting

Apply:
- Bold header
- Alternating background colors

TEST:
- Visual inspection

---

## STEP 15 – Automation Hook

After:
- addMember
- addEvent

Call:
recalculateAll()

TEST:
- Changes immediately reflected

---

## STEP 16 – Sheet Protections

Protect:
- ProcessedData
- MemberChart

TEST:
- Only Super Admin can edit

---

## STEP 17 – Role Detection

Detect via:
Session.getActiveUser().getEmail()

Assign:
Super Admin / Admin / Viewer

TEST:
- Role changes behavior

---

## STEP 18 – Role-Based Filtering

If Viewer:
- Show only own row in MemberChart

TEST:
- Viewer cannot see others

---

## STEP 19 – Insert Logo

Add image at top of MemberChart.

TEST:
- Displays correctly
- Doesn’t break formatting

---

## STEP 20 – PDF Export

Export MemberChart as PDF:
- Landscape
- Fit to width
- Proper margins

TEST:
- Layout correct
- Dollar formatting preserved

---

# 6. Safety & Testing Strategy

For every step:
- Only modify one sheet at a time
- Use Logger.log for verification
- Avoid editing historical rows during early iterations
- Use small test dataset (3 members, 2 events)

---

# 7. Deployment Strategy

Version 1:
- Super Admin only
- Manual testing

Version 2:
- Enable Admin role

Version 3:
- Enable Viewer filtering

---

# 8. Done Criteria

System is complete when:

✔ Event add recalculates automatically  
✔ Duplicate names correctly counted  
✔ Revenue allocation correct  
✔ Member chart sorted chronologically  
✔ Viewer only sees own data  
✔ PDF export is clean and formatted  
✔ No manual formula editing required  

---

END OF PLAN
