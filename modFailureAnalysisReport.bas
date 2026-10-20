Attribute VB_Name = "modFailureAnalysisReport"
Option Compare Database
Option Explicit

' Import into the FRONT END: VBA editor > File > Import File.
' Access reference required: Microsoft Office xx.x Access database engine Object Library (DAO).
' Source: linked/local table Failure_analysis, matching the supplied export.
' Run GenerateFailureReportPrompt from Immediate Window or an Access button.

' First run OpenFailureWeekReviewPrompt to edit the report inside Access.
' tblFailureReportItemsShared is stored in the same shared Access back-end as Failure_analysis.
Public Sub OpenFailureWeekReviewPrompt()
    On Error GoTo Failed
    Dim y As Long, ww As Long, s As String
    s = InputBox("ISO work-week year", "Review Failure Analysis", IsoYear(Date))
    If Len(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then Err.Raise vbObjectError + 111, , "Enter a numeric year."
    y = CLng(s)
    s = InputBox("Work week", "Review Failure Analysis", IsoWeek(Date))
    If Len(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then Err.Raise vbObjectError + 112, , "Enter a numeric week."
    ww = CLng(s)
    OpenFailureWeekReview y, ww
    Exit Sub
Failed:
    MsgBox "Could not open weekly review: " & Err.description, vbCritical
End Sub

Public Sub OpenFailureWeekReview(ByVal y As Long, ByVal ww As Long)
    On Error GoTo Failed
    Dim q As DAO.QueryDef, db As DAO.Database
    PrepareFailureWeek y, ww
    Set db = CurrentDb
    On Error Resume Next
    DoCmd.Close acForm, "frmFailureReviewRows", acSaveNo
    DoCmd.Close acQuery, "qryFailureWeekReview", acSaveNo
    db.QueryDefs.Delete "qryFailureWeekReview"
    On Error GoTo Failed
    Set q = db.CreateQueryDef("qryFailureWeekReview", _
        "SELECT f.[ID], f.[When_problem_was_resolved] AS Resolved, " & _
        "f.[Location_of_problem] AS Chamber, " & _
        "f.[Equipment/Machine Parts/Standart System/Process Requirement] AS Equipment, " & _
        "f.[Error], f.[Solution], r.[ScheduleClass], r.[ReportNote] " & _
        "FROM [Failure_analysis] AS f INNER JOIN [tblFailureReportItemsShared] AS r ON f.[ID]=r.[FailureID] " & _
        "WHERE r.[ISOYear]=" & y & " AND r.[WorkWeek]=" & ww & _
        " ORDER BY f.[When_problem_was_resolved], f.[ID]")
    InstallFailureReviewRowsForm
    DoCmd.OpenForm "frmFailureReviewRows", acFormDS, , , acFormEdit
    With Forms("frmFailureReviewRows").Controls("cboClassification")
        .RowSourceType = "Value List"
        .RowSource = "Scheduled;Unscheduled;Ignore"
        .ColumnCount = 1
        .LimitToList = True
        .Requery
    End With
    Exit Sub
Failed:
    MsgBox "Could not open weekly review: " & Err.description, vbCritical
End Sub

Private Sub InstallFailureReviewRowsForm()
    Dim frm As Access.Form, tempName As String, ctl As Object, x As Long
    If FormExists("frmFailureReviewRows") Then Exit Sub
    Set frm = CreateForm
    tempName = frm.name
    frm.Caption = "Weekly events - edit classification and report note"
    frm.RecordSource = "qryFailureWeekReview"
    frm.DefaultView = acDefViewDatasheet
    frm.AllowDatasheetView = True
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = True
    x = 0
    AddReviewText tempName, "ID", "Event ID", x, 1000, True
    AddReviewText tempName, "Resolved", "Resolved", x, 1750, True
    AddReviewText tempName, "Chamber", "Chamber", x, 1200, True
    AddReviewText tempName, "Equipment", "Equipment", x, 2800, True
    AddReviewText tempName, "Error", "Error", x, 3600, True
    AddReviewText tempName, "Solution", "Solution", x, 4200, True
    Set ctl = CreateControl(tempName, acComboBox, acDetail, , "ScheduleClass", x, 0, 2200, 360)
    ctl.name = "cboClassification"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "Scheduled;Unscheduled;Ignore"
    ctl.ColumnCount = 1
    ctl.ListWidth = 2400
    ctl.LimitToList = True
    Set ctl = CreateControl(tempName, acLabel, acDetail, "cboClassification", , x, 0, 2200, 360)
    ctl.Caption = "Classification"
    x = x + 2300
    AddReviewText tempName, "ReportNote", "Report note", x, 6500, False
    DoCmd.Save acForm, tempName
    DoCmd.Close acForm, tempName, acSaveYes
    DoCmd.Rename "frmFailureReviewRows", acForm, tempName
End Sub

Private Sub AddReviewText(ByVal formName As String, ByVal fieldName As String, _
                          ByVal controlName As String, ByRef x As Long, ByVal widthTwips As Long, ByVal readOnly As Boolean)
    Dim ctl As Object, label As Object, textName As String
    Set ctl = CreateControl(formName, acTextBox, acDetail, , fieldName, x, 0, widthTwips, 360)
    textName = "txt" & Replace(controlName, " ", "")
    ctl.name = textName
    ctl.Locked = readOnly
    Set label = CreateControl(formName, acLabel, acDetail, textName, , x, 0, widthTwips, 360)
    label.Caption = controlName
    x = x + widthTwips + 100
End Sub

' Run this once after importing the module. The form stays in this front-end.
Public Sub InstallFailureAnalysisForm()
    On Error GoTo Failed
    Dim frm As Access.Form, tempName As String, ctl As Object
    If FormExists("frmFailureAnalysisWeekly") Then
        DoCmd.OpenForm "frmFailureAnalysisWeekly"
        Exit Sub
    End If
    Set frm = CreateForm
    tempName = frm.name
    frm.Caption = "Weekly Failure Analysis"
    frm.width = 13200
    frm.Section(acDetail).Height = 6500
    frm.Section(acDetail).BackColor = RGB(246, 249, 252)
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 600, 400, 11500, 550)
    ctl.Caption = "Weekly Failure Analysis"
    ctl.FontSize = 20
    ctl.FontBold = True
    ctl.ForeColor = RGB(31, 55, 85)
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 650, 1320, 2700, 350)
    ctl.Caption = "ISO year"
    Set ctl = CreateControl(tempName, acTextBox, acDetail, , , 650, 1740, 2400, 480)
    ctl.name = "txtYear"
    ctl.DefaultValue = CStr(IsoYear(Date))
    ctl.FontSize = 12
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 3600, 1320, 2700, 350)
    ctl.Caption = "Work week"
    Set ctl = CreateControl(tempName, acTextBox, acDetail, , , 3600, 1740, 2400, 480)
    ctl.name = "txtWeek"
    ctl.DefaultValue = CStr(IsoWeek(Date))
    ctl.FontSize = 12
    Set ctl = CreateControl(tempName, acCommandButton, acDetail, , , 650, 2650, 3600, 750)
    ctl.name = "cmdReview"
    ctl.Caption = "Review weekly events"
    ctl.OnClick = "=FA_ReviewClick()"
    Set ctl = CreateControl(tempName, acCommandButton, acDetail, , , 4600, 2650, 3600, 750)
    ctl.name = "cmdExport"
    ctl.Caption = "Export Excel report"
    ctl.OnClick = "=FA_ExportClick()"
    Set ctl = CreateControl(tempName, acCommandButton, acDetail, , , 8550, 2650, 3600, 750)
    ctl.name = "cmdPDF"
    ctl.Caption = "Export PDF report"
    ctl.OnClick = "=FA_ExportPDFClick()"
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 650, 3820, 11300, 650)
    ctl.Caption = "Review: choose Scheduled, Unscheduled or Ignore, then edit the report note. Close the list before exporting."
    ctl.ForeColor = RGB(72, 82, 94)
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 650, 4750, 11300, 900)
    ctl.Caption = "Scheduled / Unscheduled is your review choice. The equipment charts count all events in the selected week."
    ctl.ForeColor = RGB(72, 82, 94)
    DoCmd.Save acForm, tempName
    DoCmd.Close acForm, tempName, acSaveYes
    DoCmd.Rename "frmFailureAnalysisWeekly", acForm, tempName
    DoCmd.OpenForm "frmFailureAnalysisWeekly"
    Exit Sub
Failed:
    MsgBox "Could not create the Failure Analysis form: " & Err.description, vbCritical
End Sub

Public Function FA_ReviewClick() As Boolean
    On Error GoTo Failed
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    OpenFailureWeekReview CLng(frm!txtYear), CLng(frm!txtWeek)
    FA_ReviewClick = True
    Exit Function
Failed:
    MsgBox "Review could not open: " & Err.description, vbExclamation
End Function

Public Function FA_ExportClick() As Boolean
    On Error GoTo Failed
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    GenerateFailureReport CLng(frm!txtYear), CLng(frm!txtWeek)
    FA_ExportClick = True
    Exit Function
Failed:
    MsgBox "Export could not start: " & Err.description, vbExclamation
End Function

Public Function FA_ExportPDFClick() As Boolean
    On Error GoTo Failed
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    GenerateFailureReport CLng(frm!txtYear), CLng(frm!txtWeek), True
    FA_ExportPDFClick = True
    Exit Function
Failed:
    MsgBox "PDF export could not start: " & Err.description, vbExclamation
End Function

Private Function FormExists(ByVal formName As String) As Boolean
    Dim obj As Access.AccessObject
    For Each obj In CurrentProject.AllForms
        If StrComp(obj.name, formName, vbTextCompare) = 0 Then FormExists = True: Exit Function
    Next obj
End Function

Public Sub PrepareFailureWeek(ByVal reportYear As Long, ByVal workWeek As Long)
    Dim db As DAO.Database, source As DAO.Recordset, review As DAO.Recordset
    Dim startDate As Date, lastWeek As Long, id As Long, initialNote As String
    If reportYear < 1900 Or reportYear > 9998 Then Err.Raise vbObjectError + 101, , "Invalid ISO year."
    lastWeek = DateDiff("d", IsoWeekStart(reportYear, 1), IsoWeekStart(reportYear + 1, 1)) \ 7
    If workWeek < 1 Or workWeek > lastWeek Then Err.Raise vbObjectError + 102, , "Invalid work week."
    Set db = CurrentDb
    EnsureSharedReviewTable db
    startDate = IsoWeekStart(reportYear, workWeek)
    Set source = db.OpenRecordset("SELECT * FROM [Failure_analysis] WHERE " & _
        "[When_problem_was_resolved]>=#" & Format$(startDate, "mm\/dd\/yyyy") & _
        "# AND [When_problem_was_resolved]<#" & Format$(DateAdd("d", 7, startDate), "mm\/dd\/yyyy") & "#", dbOpenSnapshot)
    Set review = db.OpenRecordset("tblFailureReportItemsShared", dbOpenDynaset)
    Do Until source.EOF
        If Not IsNull(source.fields("ID").value) Then
            id = CLng(source.fields("ID").value)
            review.FindFirst "[FailureID]=" & id
            If review.NoMatch Then
                review.AddNew
                review!FailureID = id
                review!IsoYear = reportYear
                review!workWeek = workWeek
                review!ScheduleClass = "Ignore"
                review!IncludeInNotes = (InStr(1, NzText(source.fields("Error").value), _
                    "Position/Etching rate check", vbTextCompare) = 0)
                initialNote = NzText(source.fields("Location_of_problem").value) & " - " & _
                    NzText(source.fields("Equipment/Machine Parts/Standart System/Process Requirement").value) & _
                    ": " & NzText(source.fields("Error").value)
                If Len(NzText(source.fields("Solution").value)) > 0 Then _
                    initialNote = initialNote & " - " & NzText(source.fields("Solution").value)
                review!ReportNote = initialNote
                review.Update
            Else
                If NzText(source.fields("Equipment/Machine Parts/Standart System/Process Requirement").value) = "Etching Profile" _
                   And NzText(review!ScheduleClass) <> "Ignore" Then
                    review.Edit
                    review!ScheduleClass = "Ignore"
                    review.Update
                ElseIf NzText(review!ScheduleClass) = "Review" Or Len(NzText(review!ScheduleClass)) = 0 Then
                    review.Edit
                    review!ScheduleClass = "Ignore"
                    review.Update
                End If
                If review!IsoYear <> reportYear Or review!workWeek <> workWeek Then
                    review.Edit
                    review!IsoYear = reportYear
                    review!workWeek = workWeek
                    review.Update
                End If
            End If
        End If
        source.MoveNext
    Loop
    review.Close: source.Close
End Sub

Private Sub EnsureSharedReviewTable(ByVal frontend As DAO.Database)
    Dim sourceLink As DAO.TableDef, newLink As DAO.TableDef
    Dim backend As DAO.Database, connectText As String, backendPath As String
    If TableExists(frontend, "tblFailureReportItemsShared") Then
        If Len(frontend.TableDefs("tblFailureReportItemsShared").Connect) = 0 Then _
            Err.Raise vbObjectError + 121, , "Review table is local; it must be linked to the shared back-end."
        GoTo CopyLocalRows
    End If
    Set sourceLink = frontend.TableDefs("Failure_analysis")
    connectText = sourceLink.Connect
    If Len(connectText) = 0 Or InStr(1, connectText, "DATABASE=", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 122, , "Failure_analysis must be linked to an Access back-end."
    If Left$(connectText, 4) = "ODBC" Then _
        Err.Raise vbObjectError + 123, , "ODBC back-end requires a separately created shared review table."
    backendPath = Mid$(connectText, InStr(1, connectText, "DATABASE=", vbTextCompare) + 9)
    If InStr(backendPath, ";") > 0 Then backendPath = Left$(backendPath, InStr(backendPath, ";") - 1)
    If Len(backendPath) = 0 Then Err.Raise vbObjectError + 124, , "Back-end path is missing."
    Set backend = DBEngine.OpenDatabase(backendPath)
    If Not TableExists(backend, "tblFailureReportItemsShared") Then
        On Error Resume Next ' A second front-end could create the table at the same time.
        backend.Execute "CREATE TABLE [tblFailureReportItemsShared] (" & _
            "[FailureID] LONG CONSTRAINT [pkFailureReportItemsShared] PRIMARY KEY, " & _
            "[ISOYear] SHORT, [WorkWeek] SHORT, [IncludeInNotes] YESNO, " & _
            "[ScheduleClass] TEXT(20), [ReportNote] MEMO)", dbFailOnError
        If Err.number <> 0 Then
            Err.Clear
            If Not TableExists(backend, "tblFailureReportItemsShared") Then _
                Err.Raise vbObjectError + 125, , "Cannot create review table in the shared back-end. Check write access."
        End If
        On Error GoTo 0
    End If
    backend.Close
    Set newLink = frontend.CreateTableDef("tblFailureReportItemsShared")
    newLink.SourceTableName = "tblFailureReportItemsShared"
    newLink.Connect = connectText
    frontend.TableDefs.Append newLink
    frontend.TableDefs("tblFailureReportItemsShared").RefreshLink

CopyLocalRows:
    ' Copy missing rows from the earlier local table. Keep it as a backup.
    If TableExists(frontend, "tblFailureReportItems") Then
        frontend.Execute "INSERT INTO [tblFailureReportItemsShared] " & _
            "([FailureID], [ISOYear], [WorkWeek], [IncludeInNotes], [ScheduleClass], [ReportNote]) " & _
            "SELECT l.[FailureID], l.[ISOYear], l.[WorkWeek], l.[IncludeInNotes], " & _
            "l.[ScheduleClass], l.[ReportNote] FROM [tblFailureReportItems] AS l " & _
            "LEFT JOIN [tblFailureReportItemsShared] AS s ON l.[FailureID]=s.[FailureID] " & _
            "WHERE s.[FailureID] IS NULL", dbFailOnError
    End If
End Sub

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If StrComp(t.name, tableName, vbTextCompare) = 0 Then TableExists = True: Exit Function
    Next t
End Function

Public Sub GenerateFailureReportPrompt()
    Dim y As Long, ww As Long, s As String
    s = InputBox("ISO work-week year (e.g. 2026)", "Failure Analysis", IsoYear(Date))
    If Len(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then MsgBox "Enter a year.", vbExclamation: Exit Sub
    y = CLng(s)
    s = InputBox("Work week (1-53)", "Failure Analysis", IsoWeek(Date))
    If Len(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then MsgBox "Enter a week number.", vbExclamation: Exit Sub
    ww = CLng(s)
    GenerateFailureReport y, ww
End Sub

Public Sub GenerateFailureReport(ByVal reportYear As Long, ByVal workWeek As Long, Optional ByVal outputPdf As Boolean = False)
    On Error GoTo Failed
    Dim weekStart As Date, nextWeek As Date, lastWeek As Long
    Dim db As DAO.Database, rs As DAO.Recordset, review As DAO.Recordset, sql As String
    Dim xl As Object, wb As Object, report As Object, data As Object, summary As Object
    Dim chamber As Object, machines As Object, seen As Object, cm As Object
    Dim chamberKey As Variant, equipment As String, id As String, key As String
    Dim r As Long, n As Long, outputPath As String, c As Long, noteRow As Long
    Dim checkedBox As String, emptyBox As String, quoteMark As String
    Dim classification As String, failureMessage As String

    If reportYear < 1900 Or reportYear > 9998 Then Err.Raise vbObjectError + 101, , "Invalid ISO year."
    lastWeek = DateDiff("d", IsoWeekStart(reportYear, 1), IsoWeekStart(reportYear + 1, 1)) \ 7
    If workWeek < 1 Or workWeek > lastWeek Then Err.Raise vbObjectError + 102, , "This ISO year has weeks 1-" & lastWeek & "."
    weekStart = IsoWeekStart(reportYear, workWeek)
    nextWeek = DateAdd("d", 7, weekStart)

    Set db = CurrentDb
    PrepareFailureWeek reportYear, workWeek
    Set review = db.OpenRecordset("tblFailureReportItemsShared", dbOpenSnapshot)
    sql = "SELECT * FROM [Failure_analysis] WHERE [When_problem_was_resolved] >= #" & _
          Format$(weekStart, "mm\/dd\/yyyy") & "# AND [When_problem_was_resolved] < #" & _
          Format$(nextWeek, "mm\/dd\/yyyy") & "# ORDER BY [When_problem_was_resolved], [ID]"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Set chamber = CreateObject("Scripting.Dictionary")
    Set machines = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    chamber.CompareMode = 1: machines.CompareMode = 1: seen.CompareMode = 1

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set report = wb.Worksheets(1): report.name = "Weekly Report"
    Set data = wb.Worksheets.Add(, report): data.name = "Event Details"
    Set summary = wb.Worksheets.Add(, data): summary.name = "Chart Data"
    checkedBox = ChrW$(&H2611)
    emptyBox = ChrW$(&H2610)
    quoteMark = Chr$(34)

    data.Range("A1:L1").value = Array("ID", "Resolved", "Tape", "Chamber", "Equipment", "Error", "Reason", "Solution", "Maintenance", "Employee", "Classification", "Report note")
    r = 2
    Do Until rs.EOF
        id = NzText(rs.fields("ID").value)
        key = id
        If Len(key) = 0 Then key = "row_" & CStr(r)
        If Not seen.Exists(key) Then
            seen.Add key, True
            chamberKey = NzText(rs.fields("Location_of_problem").value)
            If Len(chamberKey) = 0 Then chamberKey = "Unspecified"
            equipment = NzText(rs.fields("Equipment/Machine Parts/Standart System/Process Requirement").value)
            If Len(equipment) = 0 Then equipment = "Unspecified"
            Increment chamber, CStr(chamberKey)
            If Not machines.Exists(CStr(chamberKey)) Then
                Set cm = CreateObject("Scripting.Dictionary")
                cm.CompareMode = 1
                machines.Add CStr(chamberKey), cm
            End If
            Set cm = machines(CStr(chamberKey))
            Increment cm, equipment
            data.Cells(r, 1).value = id
            data.Cells(r, 2).value = rs.fields("When_problem_was_resolved").value
            data.Cells(r, 3).value = NzText(rs.fields("Date of problem").value)
            data.Cells(r, 4).value = CStr(chamberKey)
            data.Cells(r, 5).value = equipment
            data.Cells(r, 6).value = NzText(rs.fields("Error").value)
            data.Cells(r, 7).value = NzText(rs.fields("Reason_for_error").value)
            data.Cells(r, 8).value = NzText(rs.fields("Solution").value)
            data.Cells(r, 9).value = NzText(rs.fields("Type of Maintenance").value)
            data.Cells(r, 10).value = NzText(rs.fields("Employee").value)
            review.FindFirst "[FailureID]=" & CLng(id)
            If Not review.NoMatch Then
                data.Cells(r, 11).value = NzText(review!ScheduleClass)
                data.Cells(r, 12).value = NzText(review!ReportNote)
            End If
            r = r + 1
        End If
        rs.MoveNext
    Loop
    n = r - 2
    data.columns("B").NumberFormat = "yyyy-mm-dd hh:mm"
    data.columns("A:K").AutoFit
    data.columns("F:H").ColumnWidth = 32
    data.columns("K").ColumnWidth = 19
    data.columns("L").ColumnWidth = 75
    data.Range("A2:L" & ApplicationMax(2, r - 1)).WrapText = False
    data.rows("2:" & ApplicationMax(2, r - 1)).RowHeight = 28
    With data.Range("K2:K" & ApplicationMax(2, r - 1)).Validation
        .Delete
        .Add 3, 1, 1, "Scheduled,Unscheduled,Ignore" ' xlValidateList
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Choose a classification"
        .ErrorMessage = "Select Scheduled, Unscheduled, or Ignore from the list."
    End With
    data.Range("A1:L1").Font.Bold = True
    data.Range("A1:L1").Interior.color = RGB(31, 55, 85)
    data.Range("A1:L1").Font.color = RGB(255, 255, 255)
    data.Range("N2:U3").Merge
    data.Range("N2").value = "Excel edits update this workbook only. To save shared choices, edit the Access review form and export again."
    data.Range("N2:U3").WrapText = True
    data.Range("N2:U3").Interior.color = RGB(235, 242, 250)
    data.Range("A1:L" & ApplicationMax(2, r - 1)).AutoFilter
    data.Activate
    data.Range("A2").Select
    xl.ActiveWindow.FreezePanes = True

    report.Range("C3:AJ3").Merge
    report.Range("C3").value = "Failure Analysis Report (WW" & Format$(workWeek, "00") & ")"
    report.Range("C3").Font.size = 20
    report.Range("C3").Font.Bold = True
    report.Range("C4").value = reportYear & "  |  " & Format$(weekStart, "dd mmm yyyy") & " - " & Format$(nextWeek - 1, "dd mmm yyyy")
    report.Range("C56").value = "Scheduled:"
    report.Range("G56").value = "Unscheduled:"
    report.Range("J56").value = "Problems encountered / Downtime contributors:"
    report.Range("C56:AJ56").Font.Bold = True
    report.Range("E58:E" & ApplicationMax(58, n + 57)).Font.name = "Segoe UI Symbol"
    report.Range("H58:H" & ApplicationMax(58, n + 57)).Font.name = "Segoe UI Symbol"
    report.Range("E58:E" & ApplicationMax(58, n + 57)).Font.size = 15
    report.Range("H58:H" & ApplicationMax(58, n + 57)).Font.size = 15
    report.Range("E58:E" & ApplicationMax(58, n + 57)).HorizontalAlignment = -4108 ' xlCenter
    report.Range("H58:H" & ApplicationMax(58, n + 57)).HorizontalAlignment = -4108
    report.Range("C57").value = "Reviewed in Event Details"
    report.Range("G57").value = "Reviewed in Event Details"
    report.Range("J57").value = "Edit Classification in Event Details; this report updates in Excel. Charts count all events."
    report.columns("A:AJ").ColumnWidth = 4.5
    report.columns("J:AJ").ColumnWidth = 5.5
    report.columns("E").ColumnWidth = 7
    report.columns("H").ColumnWidth = 7
    report.Range("J57:AJ57").Merge
    report.Range("J57:AJ57").WrapText = True

    ' Four panels in the same positions as the existing Current Chart sheet.
    WriteCounts summary, chamber, 1, "All chambers"
    MakeChart report, summary, 1, 3, 4, "Events by chamber"
    Dim choices As Variant
    choices = TopChambers(chamber, 3)
    For c = 0 To 2
        If Len(CStr(choices(c))) > 0 Then
            Set cm = machines(CStr(choices(c)))
            WriteCounts summary, cm, 4 + c * 3, CStr(choices(c)) & " equipment"
            Select Case c
                Case 0: MakeChart report, summary, 4, 23, 4, CStr(choices(c)) & " equipment"
                Case 1: MakeChart report, summary, 7, 3, 29, CStr(choices(c)) & " equipment"
                Case 2: MakeChart report, summary, 10, 23, 30, CStr(choices(c)) & " equipment"
            End Select
        End If
    Next c
    noteRow = 58
    If outputPdf Then
        For c = 2 To n + 1
            classification = NzText(data.Cells(c, 11).value)
            If classification = "Scheduled" Or classification = "Unscheduled" Then
                report.Cells(noteRow, 5).value = IIf(classification = "Scheduled", checkedBox, emptyBox)
                report.Cells(noteRow, 8).value = IIf(classification = "Unscheduled", checkedBox, emptyBox)
                report.Cells(noteRow, 10).value = data.Cells(c, 12).value
                noteRow = noteRow + 1
            End If
        Next c
    ElseIf n > 0 Then
        For noteRow = 58 To n + 57
            report.Cells(noteRow, 37).Formula = SelectedSourceRowFormula(n, noteRow) ' AK helper
            report.Cells(noteRow, 5).Formula = "=IF(AK" & noteRow & "=" & quoteMark & quoteMark & _
                "," & quoteMark & quoteMark & ",IF(INDEX('Event Details'!$K:$K,AK" & noteRow & _
                ")=" & quoteMark & "Scheduled" & quoteMark & "," & quoteMark & checkedBox & quoteMark & _
                "," & quoteMark & emptyBox & quoteMark & "))"
            report.Cells(noteRow, 8).Formula = "=IF(AK" & noteRow & "=" & quoteMark & quoteMark & _
                "," & quoteMark & quoteMark & ",IF(INDEX('Event Details'!$K:$K,AK" & noteRow & _
                ")=" & quoteMark & "Unscheduled" & quoteMark & "," & quoteMark & checkedBox & quoteMark & _
                "," & quoteMark & emptyBox & quoteMark & "))"
            report.Cells(noteRow, 10).Formula = "=IF(AK" & noteRow & "=" & Chr$(34) & Chr$(34) & _
                "," & Chr$(34) & Chr$(34) & ",INDEX('Event Details'!$L:$L,AK" & noteRow & "))"
        Next noteRow
        report.columns("AK").Hidden = True
    End If
    report.Range("J58:J" & ApplicationMax(58, noteRow - 1)).EntireRow.RowHeight = 20
    report.Range("J58:J" & ApplicationMax(58, noteRow - 1)).WrapText = False
    report.Range("C51").value = "Events: " & n & "  |  Review classification and notes before sharing"
    report.PageSetup.Orientation = 2
    report.PageSetup.Zoom = False
    report.PageSetup.FitToPagesWide = 1
    report.PageSetup.FitToPagesTall = False
    report.PageSetup.PrintArea = "C3:AJ" & ApplicationMax(58, noteRow - 1)

    outputPath = CurrentProject.path & "\Failure_Analysis_" & reportYear & "_WW" & Format$(workWeek, "00") & _
                 IIf(outputPdf, ".pdf", ".xlsx")
    If Len(Dir$(outputPath)) > 0 Then
        outputPath = CurrentProject.path & "\Failure_Analysis_" & reportYear & "_WW" & Format$(workWeek, "00") & _
                     "_" & Format$(Now, "yyyymmdd_hhnnss") & IIf(outputPdf, ".pdf", ".xlsx")
    End If
    report.Activate
    If outputPdf Then
        report.ExportAsFixedFormat 0, outputPath
        wb.Close False
        xl.Quit
    Else
        wb.SaveAs outputPath, 51
        xl.Visible = True
        xl.DisplayAlerts = True
    End If
    rs.Close
    Set rs = Nothing
    MsgBox "Created " & n & " events:" & vbCrLf & outputPath, vbInformation
    Exit Sub
Failed:
    failureMessage = Err.description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not review Is Nothing Then review.Close
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    MsgBox "Report could not be created: " & failureMessage & vbCrLf & _
           "Check the Failure_analysis table name, field names, and Excel installation.", vbCritical
End Sub

Private Function SelectedSourceRowFormula(ByVal eventCount As Long, ByVal outputRow As Long) As String
    Dim statuses As String, q As String
    q = Chr$(34)
    statuses = "'Event Details'!$K$2:$K$" & CStr(eventCount + 1)
    SelectedSourceRowFormula = "=IFERROR(AGGREGATE(15,6,ROW(" & statuses & ")/((" & statuses & _
        "=" & q & "Scheduled" & q & ")+(" & statuses & "=" & q & "Unscheduled" & q & _
        ")),ROWS($AK$58:AK" & outputRow & "))," & q & q & ")"
End Function

Private Function IsoWeekStart(ByVal y As Long, ByVal w As Long) As Date
    Dim jan4 As Date
    jan4 = DateSerial(y, 1, 4)
    IsoWeekStart = DateAdd("d", 7 * (w - 1) - Weekday(jan4, vbMonday) + 1, jan4)
End Function

Private Function IsoWeek(ByVal d As Date) As Long
    IsoWeek = DatePart("ww", d, vbMonday, vbFirstFourDays)
End Function

Private Function IsoYear(ByVal d As Date) As Long
    IsoYear = Year(DateAdd("d", 4 - Weekday(d, vbMonday), d))
End Function

Private Function NzText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then NzText = "" Else NzText = Trim$(CStr(v))
End Function

Private Sub Increment(ByVal dict As Object, ByVal key As String)
    If dict.Exists(key) Then dict(key) = CLng(dict(key)) + 1 Else dict.Add key, 1
End Sub

Private Function ApplicationMax(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then ApplicationMax = a Else ApplicationMax = b
End Function

Private Function TopChambers(ByVal counts As Object, ByVal maxItems As Long) As Variant
    Dim result() As String, used As Object, i As Long, k As Variant, best As String, maxCount As Long
    Set used = CreateObject("Scripting.Dictionary")
    ReDim result(0 To maxItems - 1)
    For i = 0 To maxItems - 1
        best = "": maxCount = -1
        For Each k In counts.keys
            If Not used.Exists(CStr(k)) Then
                If CLng(counts(k)) > maxCount Then best = CStr(k): maxCount = CLng(counts(k))
            End If
        Next k
        If Len(best) > 0 Then used.Add best, True
        result(i) = best
    Next i
    TopChambers = result
End Function

Private Sub WriteCounts(ByVal sheet As Object, ByVal counts As Object, ByVal col As Long, ByVal heading As String)
    Dim k As Variant, row As Long
    sheet.Cells(1, col).value = heading
    sheet.Cells(2, col).value = "Category": sheet.Cells(2, col + 1).value = "Events"
    row = 3
    For Each k In counts.keys
        sheet.Cells(row, col).value = CStr(k)
        sheet.Cells(row, col + 1).value = CLng(counts(k))
        row = row + 1
    Next k
    sheet.columns(col).ColumnWidth = 32
    sheet.columns(col + 1).ColumnWidth = 12
End Sub

Private Sub MakeChart(ByVal target As Object, ByVal source As Object, ByVal sourceCol As Long, _
                      ByVal targetCol As Long, ByVal targetRow As Long, ByVal heading As String)
    Dim lastRow As Long, shape As Object, chart As Object
    lastRow = source.Cells(source.rows.count, sourceCol).End(-4162).row ' xlUp
    If lastRow < 3 Then Exit Sub
    Set shape = target.ChartObjects.Add(target.Cells(targetRow, targetCol).Left, _
                target.Cells(targetRow, targetCol).top, 560, 265)
    Set chart = shape.chart
    chart.ChartType = 51 ' xlColumnClustered
    chart.SeriesCollection.NewSeries
    chart.SeriesCollection(1).name = heading
    chart.SeriesCollection(1).XValues = source.Range(source.Cells(3, sourceCol), source.Cells(lastRow, sourceCol))
    chart.SeriesCollection(1).values = source.Range(source.Cells(3, sourceCol + 1), source.Cells(lastRow, sourceCol + 1))
    chart.HasTitle = True
    chart.ChartTitle.text = heading
    chart.HasLegend = False
    chart.SeriesCollection(1).ApplyDataLabels
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
End Sub
