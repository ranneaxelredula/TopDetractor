# TopDetractor

# Weekly Failure Analysis (`frmFailureAnalysisWeekly`)

A Microsoft Access form that turns one week of entries from the `Failure_analysis` table into a weekly failure analysis report. You pick an ISO year and work week, check that week's events in Access, then export a formatted Excel workbook (or PDF) with charts and a list of scheduled and unscheduled maintenance.

The code is the standard module `modFailureAnalysisReport`. Its latest version is `modFailureAnalysisReport.bas` in this repository. The Access front end `trial-IBAD3_V3.3.1.accdb` still has the previous version, without the chamber dropdowns, until the `.bas` file is imported (see [Setup](#setup)). `Failure_analysis` and the shared review table are linked tables in the shared back end `IBAD3_V3_backend.accdb` on the IBAD-3 network share.

The exported report follows the layout of the **Current Chart** sheet in `IBAD Top Detractor Updated.xlsx`.

## The form

| Control | What it does |
|---|---|
| **ISO year** (`txtYear`) | ISO year of the week to report. Defaults to the current ISO year. |
| **Work week** (`txtWeek`) | ISO work week (1-52, or 53 in some years). Defaults to the current week. |
| **Top-left chart** (`txtChart1`) | Always *All chambers*. |
| **Top-right chart**, **Bottom-left chart**, **Bottom-right chart** (`cboChart2`, `cboChart3`, `cboChart4`) | The chamber shown in each of the other three charts. See [Choosing the chambers](#choosing-the-chambers-for-the-charts). |
| **Review weekly events** | Opens a list of the week's events in Access, to check them and edit report notes. |
| **Export Excel report** | Builds the weekly report as an `.xlsx` workbook and opens it in Excel. |
| **Export PDF report** | Builds the same report and saves only the report sheet as a PDF. |

A week runs Monday to Sunday (ISO 8601). An event belongs to the week in which its `When_problem_was_resolved` date falls.

The form has no code module of its own. Its event properties call public functions in the standard module `modFailureAnalysisReport`:

| Event | Property value | Calls |
|---|---|---|
| Review weekly events: On Click | `=FA_ReviewClick()` | `OpenFailureWeekReview(year, week)` |
| Export Excel report: On Click | `=FA_ExportClick()` | `GenerateFailureReport(year, week, False, chambers)` |
| Export PDF report: On Click | `=FA_ExportPDFClick()` | `GenerateFailureReport(year, week, True, chambers)` |
| Form: On Load | `=FA_WeeklyFormLoad([Form])` | Fills the chamber dropdowns for the current week |
| ISO year, Work week: After Update | `=FA_WeekChanged([Form])` | Fills the chamber dropdowns again for the new week |
| Chamber dropdowns: After Update | `=FA_ChamberChanged([Form])` | Updates the note under the dropdowns |

### Choosing the chambers for the charts

The report has four charts. The top-left one always shows all chambers. The three dropdowns pick the chamber for the other three.

- Each dropdown lists every chamber (`Location_of_problem`) with events in the selected week, busiest first, with its event count, for example `PM4 (18 events)`. A blank chamber is listed as `Unspecified`. The last entry, *(none - leave this chart out)*, leaves that chart empty.
- When the form opens, or the year or week changes, the dropdowns are set to the three chambers with the most events. Chambers with the same count are taken in alphabetical order.
- Any listed chamber can be picked, so a chamber with fewer events can be shown when its event matters more.
- A note under the dropdowns appears when:
  - a chamber left out has the same number of events as one that is shown, for example *PM1, PM5 and PM6 are tied at 3 events, so PM6 isn't in a chart*
  - the same chamber is picked for two charts. The export stops with a message until one of them is changed.
- Chambers with no events in the week aren't listed, because their chart would be empty.

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
4. Opens the workbook in Excel and shows a message with the number of events and the file path. If a chamber picked on the form has no events that week, the message says its chart was left out.

The workbook has three sheets.

**Weekly Report**: the report itself, ready to print, laid out like the Current Chart sheet
- Title `Failure Analysis Report WW43(10/19-10/25)`: the work week and its Monday-Sunday dates, centred across the page.
- Four pivot charts, built from the Event Details sheet:

  | Position | Cells | Title | Shows |
  |---|---|---|---|
  | Top-left | C4:V31 | Error Count per Chamber | Events per chamber, all chambers |
  | Top-right | W4:AP31 | *Chamber* Top Detractor | Events per equipment for the chamber picked in the top-right dropdown |
  | Bottom-left | C32:V59 | *Chamber* Top Detractor | Same, bottom-left dropdown |
  | Bottom-right | W32:AP59 | *Chamber* Top Detractor | Same, bottom-right dropdown |

  - Every title has the week (`WW43(10/19-10/25)`) as its second line and is centred above its chart.
  - Bars are sorted with the most events first. Styling matches Current Chart: blue bars, values in green boxes, 32 pt titles.
  - The charts count every event in the week, including events marked Not listed.
  - Each chamber chart keeps its **Chamber** filter button, so another chamber can be picked in Excel. Its title follows the filter.
- An event count line in row 60 (`Events: n`).
- Headers in row 61, with **Scheduled:** centred over column E and **Unscheduled:** centred over column H, as on Current Chart.
- From row 63, the event list, with one line per listed event:
  - a ☑/☐ box under **Scheduled**
  - a ☑/☐ box under **Unscheduled**
  - the report note under **Problems encountered / Downtime contributors**

  The list is built from formulas that read the Event Details sheet, using a hidden helper column AU. If Event Details is edited in Excel, the list updates.
- Page setup: print area A1:AQ, landscape, one page wide, centred on the page. The sheet opens without gridlines at 55% zoom.

**Event Details**: one row per event
- Columns: ID, Resolved, Tape, Chamber, Equipment, Error, Reason, Solution, Type of Maintenance, Employee, On report as, Report note.
- *On report as* is an Excel formula that applies the rules above.
- A blank chamber or equipment shows as `Unspecified`.
- The header row is styled and frozen, with AutoFilter turned on.

**Chart Data**: the four pivot tables behind the charts (A3, D3, G3 and J3), with the chamber filters above them in row 1. Column N holds the chart titles, which the charts are linked to. After editing Event Details, use **Data > Refresh All** to update the pivot tables and charts.

Edits made in the Excel file stay in that file. To change a maintenance type or report note for good, change it in Access and export again.

### Export PDF report

Works the same way as the Excel export, with four differences:
- The event list is written as fixed values instead of formulas.
- The charts' filter buttons are hidden.
- Only the Weekly Report sheet is exported, to `Failure_Analysis_<year>_WW<nn>.pdf`.
- Excel closes without saving the workbook.

## Requirements

- Microsoft Access, full version. Access Runtime won't work, because the review list is created with `CreateForm`.
- Microsoft Excel 2010 or later on the same PC. Chart titles that follow the chart's chamber filter need Excel 2013 or later; in Excel 2010 they are fixed text.
- DAO reference: *Microsoft Office xx.x Access database engine Object Library*.
- `Failure_analysis` linked to an Access back end. ODBC back ends aren't supported for the shared review table.
- Write access to the back-end folder, so the shared table can be created on first run.
- Write access to the front-end folder, where the output files are saved.

## Setup

`trial-IBAD3_V3.3.1.accdb` already contains:
- the module `modFailureAnalysisReport` (previous version, without the chamber dropdowns)
- the form `frmFailureAnalysisWeekly`
- the review list `frmFailureReviewRows`, in the current layout with Type of Maintenance and On report as
- the query `qryFailureWeekReview`
- links to `Failure_analysis` and `tblFailureReportItemsShared` in the shared back end

**Updating a front end that already has the report**, including `trial-IBAD3_V3.3.1.accdb`. Try it on a copy first.
1. Open the VBA editor (Alt+F11). Right-click `modFailureAnalysisReport` > **Export File** to keep a backup.
2. Right-click `modFailureAnalysisReport` > **Remove** (answer *No* to exporting it again), then **File > Import File** and pick `modFailureAnalysisReport.bas` from this repository.
3. **Debug > Compile**.
4. In the Immediate window (Ctrl+G), run `UpgradeFailureReportForms` once. It adds the chamber dropdowns to `frmFailureAnalysisWeekly`, moving the buttons and hints down to make room. Running it again is safe.

Then open `frmFailureAnalysisWeekly`.

**New front end:** import `modFailureAnalysisReport.bas`, then run `InstallFailureAnalysisForm` from the Immediate window. It creates the form, with the chamber dropdowns, and opens it.

The database also contains `frmFailureAnalysis`, an older copy of the form. Its buttons call the same functions, but it has no chamber dropdowns, so its exports use the three chambers with the most events. Its hint text still describes the manual classification and its work week box defaults to a fixed week (42). Use `frmFailureAnalysisWeekly` instead.

To roll back, remove the module and import the backup from step 1. Then open `frmFailureAnalysisWeekly` in Design view and clear the form's **On Load** property and the **After Update** property of the ISO year and Work week boxes, because the older module doesn't have the functions they call. The chamber dropdowns and their labels can be deleted.
