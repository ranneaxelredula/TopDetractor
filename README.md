# TopDetractor

# Weekly Failure Analysis (`frmFailureAnalysisWeekly`)

A Microsoft Access form that turns one week of entries from the `Failure_analysis` table into a weekly failure analysis report. You pick an ISO year and work week, check that week's events in Access, then export a formatted Excel workbook (or PDF) with charts and a list of scheduled and unscheduled maintenance.

The form and its code are in the Access front end `trial-IBAD3_V3.3.1.accdb` in this repository. `Failure_analysis` and the shared review table are linked tables in the shared back end `IBAD3_V3_backend.accdb` on the IBAD-3 network share.

## The form

| Control | What it does |
|---|---|
| **ISO year** (`txtYear`) | ISO year of the week to report. Defaults to the current ISO year. |
| **Work week** (`txtWeek`) | ISO work week (1-52, or 53 in some years). Defaults to the current week. |
| **Review weekly events** | Opens a list of the week's events in Access, to check them and edit report notes. |
| **Export Excel report** | Builds the weekly report as an `.xlsx` workbook and opens it in Excel. |
| **Export PDF report** | Builds the same report and saves only the report sheet as a PDF. |

A week runs Monday to Sunday (ISO 8601). An event belongs to the week in which its `When_problem_was_resolved` date falls.

The form has no code module of its own. Each button's **On Click** property calls a public function in the standard module `modFailureAnalysisReport`:

| Button | On Click | Calls |
|---|---|---|
| Review weekly events | `=FA_ReviewClick()` | `OpenFailureWeekReview(year, week)` |
| Export Excel report | `=FA_ExportClick()` | `GenerateFailureReport(year, week)` |
| Export PDF report | `=FA_ExportPDFClick()` | `GenerateFailureReport(year, week, True)` |

Without the form, run `OpenFailureWeekReviewPrompt` or `GenerateFailureReportPrompt` from the Immediate window (Ctrl+G). Each asks for the year and week in input boxes. `GenerateFailureReportPrompt` exports to Excel only.

### Scheduled vs. Unscheduled

Events are classified automatically from the `Type of Maintenance` field, so nothing has to be classified by hand. The rules are checked in this order:

| Condition | Shown on the report as |
|---|---|
| Equipment is `Etching Profile` | **Not listed**: left off the event list, but still counted in the charts |
| Type of Maintenance is `Preventive Maintenance` | **Scheduled** |
| Any other type (`Corrective Maintenance`, `Parameter Adjustment`) or blank | **Unscheduled** |

### Shared review table

Before doing anything else, the Review button and both export buttons sync the selected week into `tblFailureReportItemsShared`. This table sits in the shared Access back end next to `Failure_analysis`, so every copy of the front end sees the same report notes. In `trial-IBAD3_V3.3.1.accdb` it is already linked.

- Each event in the week gets one row, keyed by event ID (`FailureID`) and tagged with the ISO year and work week.
- New rows get a pre-filled report note: `Chamber - Equipment: Error - Solution`.
- If an event's resolved date has moved to another week, its year/week tag is updated.
- If the table doesn't exist yet, it is created in the back end and linked into the front end. If the front end has a *local* table with this name, the sync stops with an error, because the notes would not be shared.
- If an older local `tblFailureReportItems` table is found, its rows that are missing from the shared table are copied across. The local table is kept as a backup.
- The table also has `IncludeInNotes` and `ScheduleClass` columns. `ScheduleClass` is left over from the manual-classification version. Neither column is used by the review list or the report.

## Review weekly events

1. Syncs the week into the shared review table (see above).
2. Rebuilds the saved query `qryFailureWeekReview`. It joins `Failure_analysis` to the review table for the selected week, sorted by resolved time.
3. Creates the datasheet form `frmFailureReviewRows` the first time it's needed, or rebuilds it if it still has an older layout.
4. Opens the list in datasheet view:

| Column | Editable? |
|---|---|
| Event ID, Resolved, Chamber, Equipment, Error, Solution | No |
| Type of Maintenance | No. Change it in the Type of Maintenance list on the `Failure_analysis` form. |
| On report as (Scheduled / Unscheduled / Not listed) | No. It's calculated from the rules above. |
| Report note | **Yes**. Saved straight to the shared back-end table. |

Rows can't be added or deleted here. Use the list to spot events with a blank or wrong Type of Maintenance, and to edit the wording that will appear on the report. Close the list before exporting so the last edit is saved.

## Export Excel report

1. Checks that the work week exists in that ISO year, then syncs the week into the shared review table.
2. Reads the week's `Failure_analysis` events, sorted by resolved time and then ID, and builds a new workbook through Excel automation. Excel stays hidden while the workbook is built.
3. Saves the workbook as `Failure_Analysis_<year>_WW<nn>.xlsx` in the same folder as the front-end `.accdb`. If a file with that name already exists, a timestamp (`_yyyymmdd_hhnnss`) is added to the new file's name, so nothing is overwritten.
4. Opens the workbook in Excel and shows a message with the number of events and the file path.

The workbook has three sheets.

**Weekly Report**: the report itself, ready to print
- Title `Failure Analysis Report (WWnn)`, with the ISO year and the week's date range.
- Four column charts: events by chamber, plus an equipment breakdown for each of the three chambers with the most events. The charts count every event in the week, including events marked Not listed.
- An event count line (`Events: n`).
- From row 58, the event list, with one line per listed event:
  - a ☑/☐ box under **Scheduled**
  - a ☑/☐ box under **Unscheduled**
  - the report note under **Problems encountered / Downtime contributors**

  The list is built from formulas that read the Event Details sheet, using a hidden helper column AK. If Event Details is edited in Excel, the list updates.
- Page setup: landscape, fit to one page wide.

**Event Details**: one row per event
- Columns: ID, Resolved, Tape, Chamber, Equipment, Error, Reason, Solution, Type of Maintenance, Employee, On report as, Report note.
- *On report as* is an Excel formula that applies the rules above.
- A blank chamber or equipment shows as `Unspecified`.
- The header row is styled and frozen, with AutoFilter turned on.

**Chart Data**: the counts behind the four charts.

Edits made in the Excel file stay in that file. To change a maintenance type or report note for good, change it in Access and export again.

### Export PDF report

Works the same way as the Excel export, with three differences:
- The event list is written as fixed values instead of formulas.
- Only the Weekly Report sheet is exported, to `Failure_Analysis_<year>_WW<nn>.pdf`.
- Excel closes without saving the workbook.

## Requirements

- Microsoft Access, full version. Access Runtime won't work, because the review list is created with `CreateForm`.
- Microsoft Excel 2010 or later on the same PC.
- DAO reference: *Microsoft Office xx.x Access database engine Object Library*.
- `Failure_analysis` linked to an Access back end. ODBC back ends aren't supported for the shared review table.
- Write access to the back-end folder, so the shared table can be created on first run.
- Write access to the front-end folder, where the output files are saved.

## Setup

`trial-IBAD3_V3.3.1.accdb` is already set up. It contains:
- the module `modFailureAnalysisReport`
- the form `frmFailureAnalysisWeekly`
- the review list `frmFailureReviewRows`, in the current layout with Type of Maintenance and On report as
- the query `qryFailureWeekReview`
- links to `Failure_analysis` and `tblFailureReportItemsShared` in the shared back end

Open `frmFailureAnalysisWeekly` to use it.

The database also contains `frmFailureAnalysis`, an older copy of the form. Its buttons call the same functions, but its hint text still describes the manual classification and its work week box defaults to a fixed week (42). Use `frmFailureAnalysisWeekly` instead.

To add the report to another front end:
- **New front end:** in `trial-IBAD3_V3.3.1.accdb`, export `modFailureAnalysisReport` from the VBA editor (right-click the module > Export File). Import the `.bas` file into the other front end, then run `InstallFailureAnalysisForm` from the Immediate window (Ctrl+G). It creates the form and opens it.
- **Upgrading from the manual-classification version:** replace that front end's module with the exported one, then run `UpgradeFailureReportForms` once.
