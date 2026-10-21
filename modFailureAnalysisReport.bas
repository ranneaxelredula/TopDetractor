Attribute VB_Name = "modFailureAnalysisReport"
Option Compare Database
Option Explicit

' Import into the FRONT END: VBA editor > File > Import File.
' Access reference required: Microsoft Office xx.x Access database engine Object Library (DAO).
' Source: linked/local table Failure_analysis, matching the supplied export.
' Run GenerateFailureReportPrompt from Immediate Window or an Access button.

' Scheduled / Unscheduled comes from Failure_analysis.[Type of Maintenance]:
'   "Preventive Maintenance" = Scheduled; any other value, or blank = Unscheduled.
'   Events whose equipment is "Etching Profile" are left off the report's event list
'   (they are still counted in the charts).
' The old ScheduleClass column in tblFailureReportItemsShared is kept but no longer used by the review or the report.
' After importing this version, run UpgradeFailureReportForms once from the Immediate window.

' OpenFailureWeekReviewPrompt opens the weekly list to check maintenance types and edit report notes.
' tblFailureReportItemsShared is stored in the same shared Access back-end as Failure_analysis.

' The Weekly Report sheet follows the layout of the Current Chart sheet in "IBAD Top Detractor Updated.xlsx".
' Its top-left chart shows all chambers. Three dropdowns on frmFailureAnalysisWeekly pick the chamber for the
' other three charts; they start on the three chambers with the most events (equal counts in alphabetical order).

Private Const SCHEDULED_TYPE As String = "Preventive Maintenance"
Private Const EXCLUDED_EQUIPMENT As String = "Etching Profile"
Private Const EQUIPMENT_FIELD As String = "Equipment/Machine Parts/Standart System/Process Requirement"
Private Const WEEKLY_FORM As String = "frmFailureAnalysisWeekly"
Private Const REVIEW_FORM As String = "frmFailureReviewRows"
Private Const REVIEW_CHECK_CONTROL As String = "txtOnreportas" ' only exists in the new review list layout
Private Const REVIEW_HINT As String = "Review: check Type of Maintenance and edit report notes. Close the list before exporting."
Private Const RULE_HINT As String = "Preventive Maintenance = Scheduled. Any other or blank type = Unscheduled. " & _
    "Etching Profile events are left off the report's event list. The charts count all events in the selected week."

Private Const CHART_COMBO_PREFIX As String = "cboChart" ' cboChart2, cboChart3, cboChart4 on the weekly form
Private Const CHAMBER_NOTE As String = "lblChamberNote"
Private Const NONE_CHOICE As String = "(none)"
Private Const CHAMBER_BLOCK_TOP As Long = 2400 ' weekly form controls below this move down to make room for the dropdowns
Private Const CHAMBER_BLOCK_HEIGHT As Long = 2050

' Weekly Report rows below the charts.
Private Const LIST_HEADER_ROW As Long = 61
Private Const LIST_FIRST_ROW As Long = 63
Private Const HELPER_COL As String = "AU" ' hidden helper column for the event list, outside the print area

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
    DoCmd.Close acForm, REVIEW_FORM, acSaveNo
    DoCmd.Close acQuery, "qryFailureWeekReview", acSaveNo
    db.QueryDefs.Delete "qryFailureWeekReview"
    On Error GoTo Failed
    ' OnReport uses the same rule as the Excel report (see the constants at the top).
    Set q = db.CreateQueryDef("qryFailureWeekReview", _
        "SELECT f.[ID], f.[When_problem_was_resolved] AS Resolved, " & _
        "f.[Location_of_problem] AS Chamber, " & _
        "f.[" & EQUIPMENT_FIELD & "] AS Equipment, " & _
        "f.[Error], f.[Solution], f.[Type of Maintenance] AS MaintenanceType, " & _
        "IIf(Trim(f.[" & EQUIPMENT_FIELD & "])='" & EXCLUDED_EQUIPMENT & "','Not listed'," & _
        "IIf(Trim(f.[Type of Maintenance])='" & SCHEDULED_TYPE & "','Scheduled','Unscheduled')) AS OnReport, " & _
        "r.[ReportNote] " & _
        "FROM [Failure_analysis] AS f INNER JOIN [tblFailureReportItemsShared] AS r ON f.[ID]=r.[FailureID] " & _
        "WHERE r.[ISOYear]=" & y & " AND r.[WorkWeek]=" & ww & _
        " ORDER BY f.[When_problem_was_resolved], f.[ID]")
    InstallFailureReviewRowsForm
    DoCmd.OpenForm REVIEW_FORM, acFormDS, , , acFormEdit
    If Not HasControl(Forms(REVIEW_FORM), REVIEW_CHECK_CONTROL) Then
        ' Older list with the Classification column: rebuild it once with the new columns.
        DoCmd.Close acForm, REVIEW_FORM, acSaveNo
        DoCmd.DeleteObject acForm, REVIEW_FORM
        InstallFailureReviewRowsForm
        DoCmd.OpenForm REVIEW_FORM, acFormDS, , , acFormEdit
    End If
    Exit Sub
Failed:
    MsgBox "Could not open weekly review: " & Err.description, vbCritical
End Sub

Private Sub InstallFailureReviewRowsForm()
    Dim frm As Access.Form, tempName As String, x As Long
    If FormExists(REVIEW_FORM) Then Exit Sub
    Set frm = CreateForm
    tempName = frm.name
    frm.caption = "Weekly events - edit report note (Type of Maintenance is read-only here)"
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
    AddReviewText tempName, "MaintenanceType", "Type of Maintenance", x, 2600, True
    AddReviewText tempName, "OnReport", "On report as", x, 1500, True ' creates txtOnreportas
    AddReviewText tempName, "ReportNote", "Report note", x, 6500, False
    DoCmd.Save acForm, tempName
    DoCmd.Close acForm, tempName, acSaveYes
    DoCmd.Rename REVIEW_FORM, acForm, tempName
End Sub

Private Sub AddReviewText(ByVal formName As String, ByVal fieldName As String, _
                          ByVal controlName As String, ByRef x As Long, ByVal widthTwips As Long, ByVal readOnly As Boolean)
    Dim ctl As Object, label As Object, textName As String
    Set ctl = CreateControl(formName, acTextBox, acDetail, , fieldName, x, 0, widthTwips, 360)
    textName = "txt" & Replace(controlName, " ", "")
    ctl.name = textName
    ctl.Locked = readOnly
    Set label = CreateControl(formName, acLabel, acDetail, textName, , x, 0, widthTwips, 360)
    label.caption = controlName
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
    frm.caption = "Weekly Failure Analysis"
    frm.width = 13200
    frm.Section(acDetail).Height = 6500 + CHAMBER_BLOCK_HEIGHT
    frm.Section(acDetail).BackColor = RGB(246, 249, 252)
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.OnLoad = "=FA_WeeklyFormLoad([Form])"
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 600, 400, 11500, 550)
    ctl.caption = "Weekly Failure Analysis"
    ctl.FontSize = 20
    ctl.FontBold = True
    ctl.ForeColor = RGB(31, 55, 85)
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 650, 1320, 2700, 350)
    ctl.caption = "ISO year"
    Set ctl = CreateControl(tempName, acTextBox, acDetail, , , 650, 1740, 2400, 480)
    ctl.name = "txtYear"
    ctl.DefaultValue = "=FA_IsoYear()"
    ctl.AfterUpdate = "=FA_WeekChanged([Form])"
    ctl.FontSize = 12
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 3600, 1320, 2700, 350)
    ctl.caption = "Work week"
    Set ctl = CreateControl(tempName, acTextBox, acDetail, , , 3600, 1740, 2400, 480)
    ctl.name = "txtWeek"
    ctl.DefaultValue = "=FA_IsoWeek()"
    ctl.AfterUpdate = "=FA_WeekChanged([Form])"
    ctl.FontSize = 12
    AddChamberControls tempName
    Set ctl = CreateControl(tempName, acCommandButton, acDetail, , , 650, 2650 + CHAMBER_BLOCK_HEIGHT, 3600, 750)
    ctl.name = "cmdReview"
    ctl.caption = "Review weekly events"
    ctl.OnClick = "=FA_ReviewClick()"
    Set ctl = CreateControl(tempName, acCommandButton, acDetail, , , 4600, 2650 + CHAMBER_BLOCK_HEIGHT, 3600, 750)
    ctl.name = "cmdExport"
    ctl.caption = "Export Excel report"
    ctl.OnClick = "=FA_ExportClick()"
    Set ctl = CreateControl(tempName, acCommandButton, acDetail, , , 8550, 2650 + CHAMBER_BLOCK_HEIGHT, 3600, 750)
    ctl.name = "cmdPDF"
    ctl.caption = "Export PDF report"
    ctl.OnClick = "=FA_ExportPDFClick()"
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 650, 3820 + CHAMBER_BLOCK_HEIGHT, 11300, 650)
    ctl.caption = REVIEW_HINT
    ctl.ForeColor = RGB(72, 82, 94)
    Set ctl = CreateControl(tempName, acLabel, acDetail, , , 650, 4750 + CHAMBER_BLOCK_HEIGHT, 11300, 900)
    ctl.caption = RULE_HINT
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
    GenerateFailureReport CLng(frm!txtYear), CLng(frm!txtWeek), False, ChosenChambers(frm)
    FA_ExportClick = True
    Exit Function
Failed:
    MsgBox "Export could not start: " & Err.description, vbExclamation
End Function

Public Function FA_ExportPDFClick() As Boolean
    On Error GoTo Failed
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    GenerateFailureReport CLng(frm!txtYear), CLng(frm!txtWeek), True, ChosenChambers(frm)
    FA_ExportPDFClick = True
    Exit Function
Failed:
    MsgBox "PDF export could not start: " & Err.description, vbExclamation
End Function

' Event properties of frmFailureAnalysisWeekly (=FA_...([Form]) expressions, so the form needs no code module).
Public Function FA_WeeklyFormLoad(ByVal frm As Access.Form) As Boolean
    FillChamberChoices frm, Nz(frm!txtYear, IsoYear(Date)), Nz(frm!txtWeek, IsoWeek(Date))
    FA_WeeklyFormLoad = True
End Function

Public Function FA_WeekChanged(ByVal frm As Access.Form) As Boolean
    FillChamberChoices frm, frm!txtYear, frm!txtWeek
    FA_WeekChanged = True
End Function

Public Function FA_ChamberChanged(ByVal frm As Access.Form) As Boolean
    On Error Resume Next
    UpdateChamberNote frm, WeekChamberCounts(CLng(frm!txtYear), CLng(frm!txtWeek))
    FA_ChamberChanged = True
End Function

' The chambers picked for the top-right, bottom-left and bottom-right charts ("" leaves a chart out).
' Empty when the form has no chamber dropdowns (an older copy of the form): the export then takes the top three.
Private Function ChosenChambers(ByVal frm As Access.Form) As Variant
    Dim result(0 To 2) As String, i As Long
    If Not HasControl(frm, CHART_COMBO_PREFIX & "2") Then Exit Function
    For i = 0 To 2
        result(i) = NzText(frm.Controls(CHART_COMBO_PREFIX & (i + 2)).value)
        If result(i) = NONE_CHOICE Then result(i) = ""
    Next i
    ChosenChambers = result
End Function

' Lists the week's chambers, busiest first, in the three chart dropdowns and selects the top three.
' Every chamber with events that week is listed, so a less busy chamber can be picked too.
Private Sub FillChamberChoices(ByVal frm As Access.Form, ByVal yearValue As Variant, ByVal weekValue As Variant)
    Dim counts As Object, chambers As Variant, choices As String, i As Long, n As Long, message As String
    If Not HasControl(frm, CHART_COMBO_PREFIX & "2") Then Exit Sub
    On Error GoTo NoList
    If Not IsNumeric(yearValue) Or Not IsNumeric(weekValue) Then _
        Err.Raise vbObjectError + 132, , "enter an ISO year and work week."
    Set counts = WeekChamberCounts(CLng(yearValue), CLng(weekValue))
    chambers = SortedChambers(counts)
    For i = 0 To UBound(chambers)
        n = CLng(counts(chambers(i)))
        choices = choices & ValueListItem(CStr(chambers(i))) & ";" & _
            ValueListItem(chambers(i) & " (" & n & IIf(n = 1, " event)", " events)")) & ";"
    Next i
    choices = choices & ValueListItem(NONE_CHOICE) & ";" & ValueListItem("(none - leave this chart out)")
    For i = 0 To 2
        With frm.Controls(CHART_COMBO_PREFIX & (i + 2))
            .RowSource = choices
            If i <= UBound(chambers) Then .value = chambers(i) Else .value = NONE_CHOICE
        End With
    Next i
    UpdateChamberNote frm, counts
    Exit Sub
NoList:
    ' Blank or invalid year/week, or the back end can't be reached: leave the dropdowns empty.
    message = Err.description
    On Error Resume Next
    For i = 0 To 2
        frm.Controls(CHART_COMBO_PREFIX & (i + 2)).RowSource = ""
        frm.Controls(CHART_COMBO_PREFIX & (i + 2)).value = Null
    Next i
    ShowChamberNote frm, "Chambers could not be listed: " & message
End Sub

' Warns when a chamber is picked for two charts, or when a chamber left out of the charts has the same
' number of events as one that is shown (a tie that the default pick broke alphabetically).
Private Sub UpdateChamberNote(ByVal frm As Access.Form, ByVal counts As Object)
    Dim chosen As Variant, chambers As Variant, tied As Collection, leftOut As Collection
    Dim i As Long, j As Long, tieCount As Long
    chosen = ChosenChambers(frm)
    If IsEmpty(chosen) Then Exit Sub
    For i = 0 To 1
        For j = i + 1 To 2
            If Len(chosen(i)) > 0 And StrComp(chosen(i), chosen(j), vbTextCompare) = 0 Then
                ShowChamberNote frm, chosen(i) & " is picked for more than one chart. Pick a different chamber for one of them."
                Exit Sub
            End If
        Next j
    Next i
    chambers = SortedChambers(counts)
    tieCount = -1
    For i = 0 To UBound(chambers) ' busiest first, so the first tie found is the highest one
        If Not IsChosen(chosen, CStr(chambers(i))) Then
            For j = 0 To 2
                If Len(chosen(j)) > 0 And counts.Exists(chosen(j)) Then
                    If CLng(counts(chosen(j))) = CLng(counts(chambers(i))) Then tieCount = CLng(counts(chambers(i)))
                End If
            Next j
            If tieCount >= 0 Then Exit For
        End If
    Next i
    If tieCount < 0 Then ShowChamberNote frm, "": Exit Sub
    Set tied = New Collection
    Set leftOut = New Collection
    For i = 0 To UBound(chambers)
        If CLng(counts(chambers(i))) = tieCount Then
            tied.Add chambers(i)
            If Not IsChosen(chosen, CStr(chambers(i))) Then leftOut.Add chambers(i)
        End If
    Next i
    ShowChamberNote frm, NameList(tied) & " are tied at " & tieCount & IIf(tieCount = 1, " event", " events") & _
        ", so " & NameList(leftOut) & IIf(leftOut.count = 1, " isn't", " aren't") & " in a chart. " & _
        "Pick " & IIf(leftOut.count = 1, "it", "one") & " in a dropdown to swap it in."
End Sub

Private Sub ShowChamberNote(ByVal frm As Access.Form, ByVal text As String)
    If Not HasControl(frm, CHAMBER_NOTE) Then Exit Sub
    With frm.Controls(CHAMBER_NOTE)
        If Len(text) > 0 Then .caption = text
        .Visible = (Len(text) > 0)
    End With
End Sub

Private Function IsChosen(ByVal chosen As Variant, ByVal chamberName As String) As Boolean
    Dim i As Long
    For i = LBound(chosen) To UBound(chosen)
        If StrComp(chosen(i), chamberName, vbTextCompare) = 0 Then IsChosen = True: Exit Function
    Next i
End Function

' "A", "A and B", "A, B and C".
Private Function NameList(ByVal names As Collection) As String
    Dim i As Long
    For i = 1 To names.count
        If i = 1 Then
            NameList = names(i)
        ElseIf i = names.count Then
            NameList = NameList & " and " & names(i)
        Else
            NameList = NameList & ", " & names(i)
        End If
    Next i
End Function

Private Function ValueListItem(ByVal text As String) As String
    ValueListItem = Chr$(34) & Replace(text, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

' Events per chamber (Location_of_problem) in one ISO week, counted the same way as the report's charts.
Private Function WeekChamberCounts(ByVal reportYear As Long, ByVal workWeek As Long) As Object
    Dim db As DAO.Database, rs As DAO.Recordset, counts As Object, seen As Object
    Dim startDate As Date, key As String, chamberKey As String, r As Long
    If reportYear < 1900 Or reportYear > 9998 Then Err.Raise vbObjectError + 101, , "Invalid ISO year."
    If workWeek < 1 Or workWeek > DateDiff("d", IsoWeekStart(reportYear, 1), IsoWeekStart(reportYear + 1, 1)) \ 7 Then _
        Err.Raise vbObjectError + 102, , "Invalid work week."
    startDate = IsoWeekStart(reportYear, workWeek)
    Set counts = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    counts.CompareMode = 1: seen.CompareMode = 1
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT [ID], [Location_of_problem] FROM [Failure_analysis] WHERE " & _
        "[When_problem_was_resolved]>=#" & Format$(startDate, "mm\/dd\/yyyy") & _
        "# AND [When_problem_was_resolved]<#" & Format$(DateAdd("d", 7, startDate), "mm\/dd\/yyyy") & "#", dbOpenSnapshot)
    Do Until rs.EOF
        r = r + 1
        key = NzText(rs.fields("ID").value)
        If Len(key) = 0 Then key = "row_" & CStr(r)
        If Not seen.Exists(key) Then
            seen.Add key, True
            chamberKey = NzText(rs.fields("Location_of_problem").value)
            If Len(chamberKey) = 0 Then chamberKey = "Unspecified"
            Increment counts, chamberKey
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set WeekChamberCounts = counts
End Function

' Adds the chamber dropdowns for the three single-chamber charts. The form must be open in Design view.
Private Sub AddChamberControls(ByVal formName As String)
    Dim ctl As Object, captions As Variant, i As Long, x As Long, controlName As String
    Set ctl = CreateControl(formName, acLabel, acDetail, , , 650, 2500, 6000, 350)
    ctl.caption = "Chambers shown in the charts"
    ctl.FontBold = True
    ctl.ForeColor = RGB(31, 55, 85)
    captions = Array("Top-left chart", "Top-right chart", "Bottom-left chart", "Bottom-right chart")
    For i = 0 To 3
        x = 650 + i * 2910
        If i = 0 Then
            Set ctl = CreateControl(formName, acTextBox, acDetail, , , x, 3270, 2580, 480)
            controlName = "txtChart1"
            ctl.ControlSource = "=""All chambers"""
            ctl.Locked = True
            ctl.TabStop = False
            ctl.BackColor = RGB(233, 237, 242)
        Else
            Set ctl = CreateControl(formName, acComboBox, acDetail, , , x, 3270, 2580, 480)
            controlName = CHART_COMBO_PREFIX & (i + 1)
            ctl.RowSourceType = "Value List"
            ctl.ColumnCount = 2
            ctl.BoundColumn = 1
            ctl.ColumnWidths = "0;2580" ' show "PM4 (18 events)", store "PM4"
            ctl.ListWidth = 3200
            ctl.ListRows = 12
            ctl.LimitToList = True
            ctl.AllowValueListEdits = False
            ctl.AfterUpdate = "=FA_ChamberChanged([Form])"
        End If
        ctl.name = controlName
        ctl.FontSize = 11
        Set ctl = CreateControl(formName, acLabel, acDetail, controlName, , x, 2900, 2580, 330)
        ctl.caption = captions(i)
    Next i
    Set ctl = CreateControl(formName, acLabel, acDetail, , , 650, 3870, 8100, 600)
    ctl.name = CHAMBER_NOTE
    ctl.caption = "Chamber note" ' replaced when a note is shown
    ctl.FontSize = 10
    ctl.BackStyle = 1
    ctl.BackColor = RGB(255, 243, 221)
    ctl.ForeColor = RGB(107, 63, 0)
    ctl.BorderStyle = 1
    ctl.BorderColor = RGB(240, 212, 163)
    ctl.Visible = False
End Sub

Private Function FormExists(ByVal formName As String) As Boolean
    Dim obj As Access.AccessObject
    For Each obj In CurrentProject.AllForms
        If StrComp(obj.name, formName, vbTextCompare) = 0 Then FormExists = True: Exit Function
    Next obj
End Function

Private Function HasControl(ByVal frm As Access.Form, ByVal controlName As String) As Boolean
    Dim ctl As Object
    For Each ctl In frm.Controls
        If StrComp(ctl.name, controlName, vbTextCompare) = 0 Then HasControl = True: Exit Function
    Next ctl
End Function

' Run ONCE after importing this version: Immediate window (Ctrl+G) > UpgradeFailureReportForms > Enter.
' 1. Deletes the old review list (Classification column). It is rebuilt with the new columns on the next Review.
' 2. Updates the two hint labels on frmFailureAnalysisWeekly.
' 3. Makes the ISO year / work week boxes default to the current week instead of a fixed number.
' 4. Adds the chamber dropdowns for the charts, moving the buttons and hints down to make room.
' Running it again is safe: the dropdowns are only added once.
Public Sub UpgradeFailureReportForms()
    On Error GoTo Failed
    Dim frm As Access.Form, ctl As Object, labelsChanged As Long, chambersAdded As Boolean
    If FormExists(REVIEW_FORM) Then
        On Error Resume Next
        DoCmd.Close acForm, REVIEW_FORM, acSaveNo
        On Error GoTo Failed
        DoCmd.DeleteObject acForm, REVIEW_FORM
    End If
    If Not FormExists(WEEKLY_FORM) Then
        MsgBox "Review list removed. " & WEEKLY_FORM & " was not found, so no labels were changed.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    DoCmd.Close acForm, WEEKLY_FORM, acSaveNo
    On Error GoTo Failed
    DoCmd.OpenForm WEEKLY_FORM, acDesign, , , , acHidden
    Set frm = Forms(WEEKLY_FORM)
    For Each ctl In frm.Controls
        If ctl.ControlType = acLabel Then
            If Left$(ctl.caption, 7) = "Review:" Then
                ctl.caption = REVIEW_HINT
                labelsChanged = labelsChanged + 1
            ElseIf Left$(ctl.caption, 23) = "Scheduled / Unscheduled" Then
                ctl.caption = RULE_HINT
                labelsChanged = labelsChanged + 1
            End If
        End If
    Next ctl
    frm.Controls("txtYear").DefaultValue = "=FA_IsoYear()"
    frm.Controls("txtWeek").DefaultValue = "=FA_IsoWeek()"
    If Not HasControl(frm, CHART_COMBO_PREFIX & "2") Then
        frm.Section(acDetail).Height = frm.Section(acDetail).Height + CHAMBER_BLOCK_HEIGHT
        For Each ctl In frm.Controls
            If ctl.top >= CHAMBER_BLOCK_TOP Then ctl.top = ctl.top + CHAMBER_BLOCK_HEIGHT
        Next ctl
        AddChamberControls WEEKLY_FORM
        chambersAdded = True
    End If
    frm.OnLoad = "=FA_WeeklyFormLoad([Form])"
    frm.Controls("txtYear").AfterUpdate = "=FA_WeekChanged([Form])"
    frm.Controls("txtWeek").AfterUpdate = "=FA_WeekChanged([Form])"
    DoCmd.Close acForm, WEEKLY_FORM, acSaveYes
    MsgBox "Upgrade done." & vbCrLf & _
           "- Review list will be rebuilt with Type of Maintenance on the next Review." & vbCrLf & _
           "- Hint labels updated: " & labelsChanged & vbCrLf & _
           "- Year / week boxes now default to the current ISO week." & vbCrLf & _
           "- Chamber dropdowns for the charts: " & IIf(chambersAdded, "added", "already there") & ".", vbInformation
    Exit Sub
Failed:
    MsgBox "Upgrade could not finish: " & Err.description, vbCritical
End Sub

' Used by the Default Value of txtYear / txtWeek on frmFailureAnalysisWeekly.
Public Function FA_IsoYear() As Long
    FA_IsoYear = IsoYear(Date)
End Function

Public Function FA_IsoWeek() As Long
    FA_IsoWeek = IsoWeek(Date)
End Function

Public Sub PrepareFailureWeek(ByVal reportYear As Long, ByVal workWeek As Long)
    Dim db As DAO.Database, source As DAO.Recordset, review As DAO.Recordset
    Dim startDate As Date, lastWeek As Long, ID As Long, initialNote As String
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
            ID = CLng(source.fields("ID").value)
            review.FindFirst "[FailureID]=" & ID
            If review.NoMatch Then
                review.AddNew
                review!FailureID = ID
                review!IsoYear = reportYear
                review!workWeek = workWeek
                review!IncludeInNotes = (InStr(1, NzText(source.fields("Error").value), _
                    "Position/Etching rate check", vbTextCompare) = 0)
                initialNote = NzText(source.fields("Location_of_problem").value) & " - " & _
                    NzText(source.fields("Equipment/Machine Parts/Standart System/Process Requirement").value) & _
                    ": " & NzText(source.fields("Error").value)
                If Len(NzText(source.fields("Solution").value)) > 0 Then _
                    initialNote = initialNote & " - " & NzText(source.fields("Solution").value)
                review!ReportNote = initialNote
                review.Update
            ElseIf review!IsoYear <> reportYear Or review!workWeek <> workWeek Then
                ' Scheduled / Unscheduled now comes from Type of Maintenance, so ScheduleClass is left as it is.
                review.Edit
                review!IsoYear = reportYear
                review!workWeek = workWeek
                review.Update
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

' chartChambers: the chambers for the top-right, bottom-left and bottom-right charts ("" leaves a chart out).
' Left out, the three chambers with the most events are used.
Public Sub GenerateFailureReport(ByVal reportYear As Long, ByVal workWeek As Long, Optional ByVal outputPdf As Boolean = False, _
                                 Optional ByVal chartChambers As Variant)
    On Error GoTo Failed
    Dim weekStart As Date, nextWeek As Date, lastWeek As Long
    Dim db As DAO.Database, rs As DAO.Recordset, review As DAO.Recordset, sql As String
    Dim xl As Object, wb As Object, report As Object, data As Object, summary As Object
    Dim chamber As Object, machines As Object, seen As Object, cm As Object
    Dim chamberKey As Variant, equipment As String, ID As String, key As String
    Dim r As Long, n As Long, outputPath As String, c As Long, noteRow As Long, lastListRow As Long
    Dim checkedBox As String, emptyBox As String, quoteMark As String
    Dim classification As String, failureMessage As String
    Dim weekText As String, choices As Variant, chartAreas As Variant, chosenName As String, chartKey As String
    Dim cache As Object, pt As Object, pageCell As String, skipped As String, co As Object

    If reportYear < 1900 Or reportYear > 9998 Then Err.Raise vbObjectError + 101, , "Invalid ISO year."
    lastWeek = DateDiff("d", IsoWeekStart(reportYear, 1), IsoWeekStart(reportYear + 1, 1)) \ 7
    If workWeek < 1 Or workWeek > lastWeek Then Err.Raise vbObjectError + 102, , "This ISO year has weeks 1-" & lastWeek & "."
    If IsMissing(chartChambers) Then chartChambers = Empty
    If Not IsEmpty(chartChambers) Then CheckChartChambers chartChambers
    weekStart = IsoWeekStart(reportYear, workWeek)
    nextWeek = DateAdd("d", 7, weekStart)
    weekText = WeekLabel(reportYear, workWeek)

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
    xl.visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set report = wb.Worksheets(1): report.name = "Weekly Report"
    Set data = wb.Worksheets.Add(, report): data.name = "Event Details"
    Set summary = wb.Worksheets.Add(, data): summary.name = "Chart Data"
    checkedBox = ChrW$(&H2611)
    emptyBox = ChrW$(&H2610)
    quoteMark = Chr$(34)

    data.Range("A1:L1").value = Array("ID", "Resolved", "Tape", "Chamber", "Equipment", "Error", "Reason", "Solution", "Type of Maintenance", "Employee", "On report as", "Report note")
    r = 2
    Do Until rs.EOF
        ID = NzText(rs.fields("ID").value)
        key = ID
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
            data.Cells(r, 1).value = ID
            data.Cells(r, 2).value = rs.fields("When_problem_was_resolved").value
            data.Cells(r, 3).value = NzText(rs.fields("Date of problem").value)
            data.Cells(r, 4).value = CStr(chamberKey)
            data.Cells(r, 5).value = equipment
            data.Cells(r, 6).value = NzText(rs.fields("Error").value)
            data.Cells(r, 7).value = NzText(rs.fields("Reason_for_error").value)
            data.Cells(r, 8).value = NzText(rs.fields("Solution").value)
            data.Cells(r, 9).value = NzText(rs.fields("Type of Maintenance").value)
            data.Cells(r, 10).value = NzText(rs.fields("Employee").value)
            ' On report as: Etching Profile -> Not listed; Preventive Maintenance -> Scheduled; anything else -> Unscheduled.
            data.Cells(r, 11).Formula = "=IF(TRIM(E" & r & ")=" & quoteMark & EXCLUDED_EQUIPMENT & quoteMark & "," & _
                quoteMark & "Not listed" & quoteMark & ",IF(TRIM(I" & r & ")=" & quoteMark & SCHEDULED_TYPE & quoteMark & "," & _
                quoteMark & "Scheduled" & quoteMark & "," & quoteMark & "Unscheduled" & quoteMark & "))"
            review.FindFirst "[FailureID]=" & CLng(ID)
            If Not review.NoMatch Then data.Cells(r, 12).value = NzText(review!ReportNote)
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
    data.rows("2:" & ApplicationMax(2, r - 1)).rowHeight = 28
    data.Range("A1:L1").Font.Bold = True
    data.Range("A1:L1").Interior.color = RGB(31, 55, 85)
    data.Range("A1:L1").Font.color = RGB(255, 255, 255)
    data.Range("N2:U5").Merge
    data.Range("N2").value = "On report as comes from Type of Maintenance (Preventive Maintenance = Scheduled, other or blank = Unscheduled; " & _
        "Etching Profile = Not listed). Excel edits update this workbook only. To change a type, edit the event in Access and export again. " & _
        "The charts are pivot charts: after editing this sheet, use Data > Refresh All to update them."
    data.Range("N2:U5").WrapText = True
    data.Range("N2:U5").Interior.color = RGB(235, 242, 250)
    data.Range("A1:L" & ApplicationMax(2, r - 1)).AutoFilter
    data.Activate
    data.Range("A2").Select
    xl.ActiveWindow.FreezePanes = True

    ' Same grid as the Current Chart sheet: default column widths, wider E and H for the Scheduled / Unscheduled boxes.
    report.columns("E").ColumnWidth = 14
    report.columns("H").ColumnWidth = 14
    report.rows(3).rowHeight = 63.75
    SetCenteredText report.Range("C3:AP3"), "Failure Analysis Report " & weekText, 48, True
    SetCenteredText report.Range("D" & LIST_HEADER_ROW & ":F" & LIST_HEADER_ROW), "Scheduled:", 22, True
    SetCenteredText report.Range("G" & LIST_HEADER_ROW & ":I" & LIST_HEADER_ROW), "Unscheduled:", 22, True
    SetCenteredText report.Range("J" & LIST_HEADER_ROW & ":AP" & LIST_HEADER_ROW), "Problems encountered / Downtime contributors:", 22, True
    report.rows(LIST_HEADER_ROW).rowHeight = 28.5
    SetCenteredText report.Range("D" & (LIST_HEADER_ROW + 1) & ":F" & (LIST_HEADER_ROW + 1)), "From Type of Maintenance", 11, False
    SetCenteredText report.Range("G" & (LIST_HEADER_ROW + 1) & ":I" & (LIST_HEADER_ROW + 1)), "From Type of Maintenance", 11, False
    SetCenteredText report.Range("J" & (LIST_HEADER_ROW + 1) & ":AP" & (LIST_HEADER_ROW + 1)), _
        "Preventive Maintenance = Scheduled; other or blank types = Unscheduled. Etching Profile events are not listed. Charts count all events.", 11, False
    report.rows(LIST_HEADER_ROW + 1).rowHeight = 30
    lastListRow = ApplicationMax(LIST_FIRST_ROW, n + LIST_FIRST_ROW - 1)
    With report.Range("E" & LIST_FIRST_ROW & ":E" & lastListRow & ",H" & LIST_FIRST_ROW & ":H" & lastListRow)
        .Font.name = "Segoe UI Symbol"
        .Font.size = 20
        .HorizontalAlignment = -4108 ' xlCenter, under the centred Scheduled / Unscheduled headings
    End With
    report.Range("J" & LIST_FIRST_ROW & ":J" & lastListRow).Font.name = "Arial"
    report.Range("J" & LIST_FIRST_ROW & ":J" & lastListRow).Font.size = 20

    ' Four pivot charts in the places used on the Current Chart sheet. Top-left: all chambers. The other three
    ' show one chamber each (picked on the form, else the busiest three); its chamber filter can be changed in Excel.
    If IsEmpty(chartChambers) Then choices = TopChambers(chamber, 3) Else choices = chartChambers
    chartAreas = Array("C4:V31", "W4:AP31", "C32:V59", "W32:AP59")
    If n > 0 Then
        Set cache = wb.PivotCaches.Create(1, "'Event Details'!R1C1:R" & (n + 1) & "C12", 4) ' xlDatabase, Excel 2010 pivot
        Set pt = AddEventPivot(cache, summary.Range("A3"), "ptAllChambers", "Chamber", "")
        summary.Range("N1").value = "Chart titles"
        summary.Range("N2").Formula = "=""Error Count per Chamber""&CHAR(10)&""" & weekText & """"
        AddPivotChart report.Range(chartAreas(0)), pt, summary.Range("N2"), MaxCount(chamber)
        For c = 0 To 2
            chosenName = ""
            If LBound(choices) + c <= UBound(choices) Then chosenName = CStr(choices(LBound(choices) + c))
            If Len(chosenName) > 0 Then
                chartKey = DictKey(chamber, chosenName)
                If Len(chartKey) = 0 Then
                    skipped = skipped & vbCrLf & chosenName & " has no events in " & weekText & ", so its chart was left out."
                Else
                    Set pt = AddEventPivot(cache, summary.Cells(3, 4 + c * 3), "ptChart" & (c + 2), "Equipment", chartKey)
                    pageCell = pt.TableRange2.Cells(1, 2).Address(False, False) ' the chamber filter cell above the table
                    summary.Cells(3 + c, 14).Formula = "=IF(" & pageCell & "=""(All)"",""All chambers"",IF(" & pageCell & _
                        "=""(Multiple Items)"",""Several chambers""," & pageCell & "))&"" Top Detractor""&CHAR(10)&""" & weekText & """"
                    AddPivotChart report.Range(chartAreas(c + 1)), pt, summary.Cells(3 + c, 14), MaxCount(machines(chartKey))
                End If
            End If
        Next c
    End If
    noteRow = LIST_FIRST_ROW
    If outputPdf Then
        xl.Calculate ' make sure the On report as formulas are evaluated before reading them
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
        For noteRow = LIST_FIRST_ROW To lastListRow
            report.Range(HELPER_COL & noteRow).Formula = SelectedSourceRowFormula(n, noteRow)
            report.Cells(noteRow, 5).Formula = "=IF(" & HELPER_COL & noteRow & "=" & quoteMark & quoteMark & _
                "," & quoteMark & quoteMark & ",IF(INDEX('Event Details'!$K:$K," & HELPER_COL & noteRow & _
                ")=" & quoteMark & "Scheduled" & quoteMark & "," & quoteMark & checkedBox & quoteMark & _
                "," & quoteMark & emptyBox & quoteMark & "))"
            report.Cells(noteRow, 8).Formula = "=IF(" & HELPER_COL & noteRow & "=" & quoteMark & quoteMark & _
                "," & quoteMark & quoteMark & ",IF(INDEX('Event Details'!$K:$K," & HELPER_COL & noteRow & _
                ")=" & quoteMark & "Unscheduled" & quoteMark & "," & quoteMark & checkedBox & quoteMark & _
                "," & quoteMark & emptyBox & quoteMark & "))"
            report.Cells(noteRow, 10).Formula = "=IF(" & HELPER_COL & noteRow & "=" & Chr$(34) & Chr$(34) & _
                "," & Chr$(34) & Chr$(34) & ",INDEX('Event Details'!$L:$L," & HELPER_COL & noteRow & "))"
        Next noteRow
        report.columns(HELPER_COL).Hidden = True
    End If
    report.Range("J" & LIST_FIRST_ROW & ":J" & ApplicationMax(LIST_FIRST_ROW, noteRow - 1)).EntireRow.rowHeight = 30
    report.Range("J" & LIST_FIRST_ROW & ":J" & ApplicationMax(LIST_FIRST_ROW, noteRow - 1)).WrapText = False
    report.Range("C" & (LIST_HEADER_ROW - 1)).value = "Events: " & n & "  |  Review report notes before sharing"
    report.Range("C" & (LIST_HEADER_ROW - 1)).Font.size = 14
    report.PageSetup.Orientation = 2
    report.PageSetup.Zoom = False
    report.PageSetup.FitToPagesWide = 1
    report.PageSetup.FitToPagesTall = False
    report.PageSetup.LeftMargin = xl.InchesToPoints(0.25)
    report.PageSetup.RightMargin = xl.InchesToPoints(0.25)
    report.PageSetup.CenterHorizontally = True
    report.PageSetup.PrintArea = "A1:AQ" & (ApplicationMax(LIST_FIRST_ROW, noteRow - 1) + 1)

    outputPath = CurrentProject.path & "\Failure_Analysis_" & reportYear & "_WW" & Format$(workWeek, "00") & _
                 IIf(outputPdf, ".pdf", ".xlsx")
    If Len(Dir$(outputPath)) > 0 Then
        outputPath = CurrentProject.path & "\Failure_Analysis_" & reportYear & "_WW" & Format$(workWeek, "00") & _
                     "_" & Format$(Now, "yyyymmdd_hhnnss") & IIf(outputPdf, ".pdf", ".xlsx")
    End If
    report.Activate
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.Zoom = 55
    If outputPdf Then
        For Each co In report.ChartObjects
            co.chart.ShowAllFieldButtons = False ' no filter buttons in the PDF
        Next co
        report.ExportAsFixedFormat 0, outputPath
        wb.Close False
        xl.Quit
    Else
        wb.SaveAs outputPath, 51
        xl.visible = True
        xl.DisplayAlerts = True
    End If
    rs.Close
    Set rs = Nothing
    review.Close
    Set review = Nothing
    MsgBox "Created " & n & " events:" & vbCrLf & outputPath & skipped, vbInformation
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
        ")),ROWS($" & HELPER_COL & "$" & LIST_FIRST_ROW & ":" & HELPER_COL & outputRow & "))," & q & q & ")"
End Function

' "WW43(10/19-10/25)": the work week and its Monday-Sunday dates, written as on the Current Chart sheet.
Private Function WeekLabel(ByVal reportYear As Long, ByVal workWeek As Long) As String
    Dim weekStart As Date
    weekStart = IsoWeekStart(reportYear, workWeek)
    WeekLabel = "WW" & Format$(workWeek, "00") & "(" & Format$(weekStart, "mm\/dd") & "-" & _
                Format$(DateAdd("d", 6, weekStart), "mm\/dd") & ")"
End Function

Private Sub CheckChartChambers(ByVal chartChambers As Variant)
    Dim i As Long, j As Long
    For i = LBound(chartChambers) To UBound(chartChambers)
        For j = i + 1 To UBound(chartChambers)
            If Len(chartChambers(i)) > 0 And StrComp(chartChambers(i), chartChambers(j), vbTextCompare) = 0 Then _
                Err.Raise vbObjectError + 131, , chartChambers(i) & " is picked for more than one chart. Pick a different chamber for one of them."
        Next j
    Next i
End Sub

Private Sub SetCenteredText(ByVal target As Object, ByVal text As String, ByVal fontSize As Single, ByVal isBold As Boolean)
    target.Merge
    target.value = text
    target.Font.size = fontSize
    target.Font.Bold = isBold
    target.HorizontalAlignment = -4108 ' xlCenter
    target.VerticalAlignment = -4108
    target.WrapText = True
End Sub

' A pivot table on Event Details counting events per rowField, busiest first, filtered to one chamber if given.
Private Function AddEventPivot(ByVal cache As Object, ByVal destination As Object, ByVal tableName As String, _
                               ByVal rowField As String, ByVal chamberItem As String) As Object
    Dim pt As Object
    Set pt = cache.CreatePivotTable(destination, tableName, True, 4)
    If Len(chamberItem) > 0 Then pt.PivotFields("Chamber").Orientation = 3 ' xlPageField
    pt.PivotFields(rowField).Orientation = 1 ' xlRowField
    pt.AddDataField pt.PivotFields("Resolved"), "Events", -4112 ' xlCount: every event has a resolved date
    pt.PivotFields(rowField).AutoSort 2, "Events" ' xlDescending
    If Len(chamberItem) > 0 Then pt.PivotFields("Chamber").CurrentPage = chamberItem
    Set AddEventPivot = pt
End Function

' A column chart on a pivot table, styled like the Current Chart sheet, with a centred title linked to titleCell.
Private Sub AddPivotChart(ByVal area As Object, ByVal pt As Object, ByVal titleCell As Object, ByVal largestCount As Long)
    Dim shape As Object, chart As Object
    Set shape = area.Worksheet.ChartObjects.Add(area.Left, area.top, area.width, area.Height)
    Set chart = shape.chart
    chart.SetSourceData pt.TableRange1
    chart.ChartType = 51 ' xlColumnClustered
    chart.HasTitle = True
    On Error Resume Next
    chart.ChartTitle.Formula = "='" & titleCell.Worksheet.name & "'!" & titleCell.Address
    If Err.number <> 0 Then
        Err.Clear
        chart.ChartTitle.text = CStr(titleCell.value) ' Excel 2010 has no linked titles: fixed text instead
    End If
    chart.ChartTitle.Position = -4105 ' xlChartElementPositionAutomatic: centred above the plot
    On Error GoTo 0
    chart.ChartTitle.Font.size = 32
    chart.ChartTitle.Font.Bold = True
    chart.ChartTitle.HorizontalAlignment = -4108 ' xlCenter, for both title lines
    chart.ChartArea.Format.Line.Visible = -1
    chart.ChartArea.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
    chart.ChartArea.Format.Line.Weight = 0.75
    chart.HasLegend = True
    chart.Legend.Position = -4152 ' xlLegendPositionRight
    chart.Legend.Font.size = 28
    chart.Axes(1).TickLabels.Font.size = 24
    chart.Axes(2).TickLabels.Font.size = 28
    chart.Axes(2).MinimumScale = 0
    If largestCount <= 10 Then chart.Axes(2).MajorUnit = 1 ' whole numbers only for small counts
    chart.ChartGroups(1).GapWidth = 219
    chart.ChartGroups(1).Overlap = -27
    With chart.SeriesCollection(1)
        .Format.Fill.Visible = -1
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = RGB(15, 158, 213)
        .HasDataLabels = True
        .DataLabels.Font.size = 28
        .DataLabels.Format.Fill.Visible = -1
        .DataLabels.Format.Fill.Solid
        .DataLabels.Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
    End With
End Sub

Private Function MaxCount(ByVal counts As Object) As Long
    Dim k As Variant
    For Each k In counts.keys
        If CLng(counts(k)) > MaxCount Then MaxCount = CLng(counts(k))
    Next k
End Function

' The key as stored in a case-insensitive dictionary, or "" when it isn't there.
Private Function DictKey(ByVal dict As Object, ByVal wanted As String) As String
    Dim k As Variant
    For Each k In dict.keys
        If StrComp(CStr(k), wanted, vbTextCompare) = 0 Then DictKey = CStr(k): Exit Function
    Next k
End Function

Private Function IsoWeekStart(ByVal y As Long, ByVal w As Long) As Date
    Dim jan4 As Date
    jan4 = DateSerial(y, 1, 4)
    IsoWeekStart = DateAdd("d", 7 * (w - 1) - Weekday(jan4, vbMonday) + 1, jan4)
End Function

Private Function IsoWeek(ByVal d As Date) As Long
    ' DatePart("ww", d, vbMonday, vbFirstFourDays) can return 53 for some late-December Mondays,
    ' so count whole weeks from the Monday of ISO week 1 instead.
    IsoWeek = DateDiff("d", IsoWeekStart(IsoYear(d), 1), d) \ 7 + 1
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

' The first maxItems chambers by event count ("" where there are fewer chambers).
Private Function TopChambers(ByVal counts As Object, ByVal maxItems As Long) As Variant
    Dim result() As String, chambers As Variant, i As Long
    ReDim result(0 To maxItems - 1)
    chambers = SortedChambers(counts)
    For i = 0 To maxItems - 1
        If i <= UBound(chambers) Then result(i) = chambers(i)
    Next i
    TopChambers = result
End Function

' Chamber names, most events first; equal counts in alphabetical order, so ties are always broken the same way.
Private Function SortedChambers(ByVal counts As Object) As Variant
    Dim names() As String, k As Variant, i As Long, j As Long
    If counts.count = 0 Then SortedChambers = Array(): Exit Function
    ReDim names(0 To counts.count - 1)
    For Each k In counts.keys
        j = i - 1
        Do While j >= 0
            If Not ComesBefore(counts, CStr(k), names(j)) Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = CStr(k)
        i = i + 1
    Next k
    SortedChambers = names
End Function

Private Function ComesBefore(ByVal counts As Object, ByVal a As String, ByVal b As String) As Boolean
    If CLng(counts(a)) <> CLng(counts(b)) Then
        ComesBefore = CLng(counts(a)) > CLng(counts(b))
    Else
        ComesBefore = StrComp(a, b, vbTextCompare) < 0
    End If
End Function
