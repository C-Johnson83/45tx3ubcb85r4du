Attribute VB_Name = "modDefectDrillDown"
Option Explicit

' Run once with the dashboard closed. Requires the existing bar-chart upgrade.
' Re-running repairs an incomplete drill-down installation.
Public Sub InstallDefectDrillDownFixed()
    Dim prj As Object, comp As Object, cm As Object, d As Object
    Dim lb As Object, tb As Object, btn As Object
    Dim p As Long, n As Long, i As Long
    On Error GoTo Failed

    Set prj = ThisWorkbook.VBProject
    Set comp = prj.VBComponents("frmDefectDashboard")
    Set cm = comp.CodeModule
    If Not cm.Find("Private Sub RefreshBarCharts()", 1, 1, cm.CountOfLines, 500) Then
        Err.Raise 5, , "Install the bar-chart upgrade first."
    End If

    ' Remove incomplete prior code injection, if present.
    RemoveDrillProcedure cm, "ShowDrillForBar"
    RemoveDrillProcedure cm, "InitDrillClicks"
    For i = cm.CountOfLines To 1 Step -1
        Select Case Trim$(cm.Lines(i, 1))
            Case "Private mDrillClicks As Collection", "InitDrillClicks"
                cm.DeleteLines i, 1
        End Select
    Next i

    ' Replace only drill-down components, never the existing dashboard.
    Set comp = Nothing
    On Error Resume Next
    Set comp = prj.VBComponents("frmDefectDrillDown")
    On Error GoTo Failed
    If Not comp Is Nothing Then
        If StrComp(comp.Name, "frmDefectDashboard", vbTextCompare) = 0 Then _
            Err.Raise 5, , "Safety stop: refusing to remove dashboard."
        prj.VBComponents.Remove comp
    End If
    Set comp = Nothing
    Set comp = Nothing
    On Error Resume Next
    Set comp = prj.VBComponents("clsDefectBarClick")
    On Error GoTo Failed
    If Not comp Is Nothing Then prj.VBComponents.Remove comp
    Set comp = Nothing

    Set comp = prj.VBComponents.Add(2)
    comp.Name = "clsDefectBarClick"
    AppendGeneratedCode comp.CodeModule, ClickClassCode()

    Set comp = prj.VBComponents.Add(3)
    comp.Name = "frmDefectDrillDown"
    Set d = comp.designer
    MakeDrillLabel d, "lblDrillTitle", "DEFECT DRILL-DOWN", 16, 12, 790, 24, 15, True
    MakeDrillLabel d, "lblDrillCount", "", 16, 41, 790, 20, 10, False
    MakeDrillLabel d, "lblHeaders", "Submission ID        Date of Defect        Z Number         Name                    Defect                    Status", 16, 73, 795, 18, 9, True
    Set lb = d.Controls.Add("Forms.ListBox.1", "lstDrillRecords", True)
    With lb
        .Left = 16: .Top = 94: .Width = 797: .Height = 232
        .ColumnCount = 6
        .ColumnWidths = "112 pt;95 pt;78 pt;120 pt;285 pt;65 pt"
        .Font.Size = 9
        .IntegralHeight = False
    End With
    MakeDrillLabel d, "lblDetails", "SELECT A RECORD TO VIEW ITS FIELDS", 16, 337, 790, 18, 10, True
    Set tb = d.Controls.Add("Forms.TextBox.1", "txtDrillDetails", True)
    With tb
        .Left = 16: .Top = 360: .Width = 797: .Height = 153
        .MultiLine = True: .ScrollBars = 2: .Locked = True
        .Font.Name = "Consolas": .Font.Size = 9
    End With
    Set btn = d.Controls.Add("Forms.CommandButton.1", "cmdDrillClose", True)
    With btn
        .caption = "Close": .Left = 713: .Top = 522
        .Width = 100: .Height = 25
    End With
    AppendGeneratedCode comp.CodeModule, DetailFormCode()

    Set comp = prj.VBComponents("frmDefectDashboard")
    Set cm = comp.CodeModule
    cm.InsertLines cm.CountOfDeclarationLines + 1, "Private mDrillClicks As Collection"
    cm.AddFromString DashboardDrillCode()
    p = cm.ProcStartLine("UserForm_Initialize", 0)
    n = cm.ProcCountLines("UserForm_Initialize", 0)
    For i = p + n - 1 To p Step -1
        If Trim$(cm.Lines(i, 1)) = "End Sub" Then
            cm.InsertLines i, "    InitDrillClicks"
            Exit For
        End If
    Next i
    If i < p Then Err.Raise 5, , "Could not find UserForm_Initialize End Sub."
    MsgBox "Drill-down repaired. Reopen dashboard and click a bar.", vbInformation
    Exit Sub
Failed:
    MsgBox "Installation error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub AppendGeneratedCode(ByVal cm As Object, ByVal generated As String)
    ' New VBA modules may already contain Option Explicit.
    ' Add it only if it is not present, and always before procedures.
    Dim i As Long, found As Boolean
    For i = 1 To cm.CountOfDeclarationLines
        If StrComp(Trim$(cm.Lines(i, 1)), "Option Explicit", vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then cm.InsertLines 1, "Option Explicit"
    cm.AddFromString generated
End Sub

Private Sub RemoveDrillProcedure(ByVal cm As Object, ByVal procName As String)
    Dim p As Long, n As Long
    On Error Resume Next
    p = cm.ProcStartLine(procName, 0)
    If Err.Number = 0 And p > 0 Then
        n = cm.ProcCountLines(procName, 0)
        cm.DeleteLines p, n
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub MakeDrillLabel(ByVal d As Object, ByVal nm As String, _
    ByVal caption As String, ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single, ByVal fs As Single, _
    ByVal isBold As Boolean)
    Dim c As Object
    Set c = d.Controls.Add("Forms.Label.1", nm, True)
    With c
        .caption = caption: .Left = x: .Top = y
        .Width = w: .Height = h
        .Font.Size = fs: .Font.Bold = isBold
        .ForeColor = RGB(32, 52, 85): .BackStyle = 0
    End With
End Sub

Private Function ClickClassCode() As String
    Dim s As String
    s = s & "Public WithEvents Target As MSForms.Label" & vbCrLf
    s = s & "Public Owner As Object" & vbCrLf
    s = s & "Public Panel As String" & vbCrLf
    s = s & "Public Slot As Long" & vbCrLf
    s = s & "Private Sub Target_Click()" & vbCrLf
    s = s & "    Owner.ShowDrillForBar Panel, Slot" & vbCrLf
    s = s & "End Sub" & vbCrLf
    ClickClassCode = s
End Function

Private Function DashboardDrillCode() As String
    Dim s As String
    s = s & "Private Sub InitDrillClicks()" & vbCrLf
    s = s & "    Dim panel As Variant, slot As Long, part As Variant" & vbCrLf
    s = s & "    Dim handler As clsDefectBarClick" & vbCrLf
    s = s & "    Set mDrillClicks = New Collection" & vbCrLf
    s = s & "    For Each panel In Array(""Type"", ""Key"", ""Supervisor"", ""Day"")" & vbCrLf
    s = s & "        For slot = 0 To 7" & vbCrLf
    s = s & "            For Each part In Array(""Name"", ""Bar"", ""Value"")" & vbCrLf
    s = s & "                Set handler = New clsDefectBarClick" & vbCrLf
    s = s & "                Set handler.Target = Me.Controls(""bc"" & panel & part & slot)" & vbCrLf
    s = s & "                Set handler.Owner = Me" & vbCrLf
    s = s & "                handler.Panel = CStr(panel)" & vbCrLf
    s = s & "                handler.Slot = slot" & vbCrLf
    s = s & "                mDrillClicks.Add handler" & vbCrLf
    s = s & "            Next part" & vbCrLf
    s = s & "        Next slot" & vbCrLf
    s = s & "    Next panel" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ShowDrillForBar(ByVal panel As String, ByVal slot As Long)" & vbCrLf
    s = s & "    Dim lb As MSForms.ListBox, idx As Long, startIdx As Long" & vbCrLf
    s = s & "    Dim fieldName As String, selected As String" & vbCrLf
    s = s & "    Dim k As Variant, id As String, status As String" & vbCrLf
    s = s & "    Dim defectDate As Date, cutoff As Date" & vbCrLf
    s = s & "    Dim keys As Collection, results As Collection" & vbCrLf
    s = s & "    Dim detail As Object" & vbCrLf
    s = s & "    Dim raw As String" & vbCrLf
    s = s & "    If mCSV Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    Select Case panel" & vbCrLf
    s = s & "        Case ""Type"": Set lb = Me.lstByType: fieldName = ""Defect""" & vbCrLf
    s = s & "        Case ""Key"": Set lb = Me.lstByKey: fieldName = ""Key""" & vbCrLf
    s = s & "        Case ""Supervisor"": Set lb = Me.lstBySupervisor: fieldName = ""Supervisor""" & vbCrLf
    s = s & "        Case ""Day"": Set lb = Me.lstByDay: fieldName = ""Date of Defect""" & vbCrLf
    s = s & "        Case Else: Exit Sub" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    If panel = ""Day"" And lb.ListCount > 8 Then" & vbCrLf
    s = s & "        startIdx = lb.ListCount - 8" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    idx = startIdx + slot" & vbCrLf
    s = s & "    If idx < 0 Or idx >= lb.ListCount Then Exit Sub" & vbCrLf
    s = s & "    selected = CStr(lb.List(idx, 0))" & vbCrLf
    s = s & "    Select Case Me.cboDateRange.Value" & vbCrLf
    s = s & "        Case ""Last 7 Days"": cutoff = DateAdd(""d"", -6, Date)" & vbCrLf
    s = s & "        Case ""Last 30 Days"": cutoff = DateAdd(""d"", -29, Date)" & vbCrLf
    s = s & "        Case ""Last 90 Days"": cutoff = DateAdd(""d"", -89, Date)" & vbCrLf
    s = s & "        Case Else: cutoff = 0" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    Set results = New Collection" & vbCrLf
    s = s & "    Set keys = mCSV.rowKeys" & vbCrLf
    s = s & "    For Each k In keys" & vbCrLf
    s = s & "        id = CStr(k)" & vbCrLf
    s = s & "        If Not PassField(id, ""Key"", Me.cboKey.Value) Then GoTo SkipRow" & vbCrLf
    s = s & "        If Not PassField(id, ""Supervisor"", Me.cboSupervisor.Value) Then GoTo SkipRow" & vbCrLf
    s = s & "        If Not PassField(id, ""Defect"", Me.cboDefect.Value) Then GoTo SkipRow" & vbCrLf
    s = s & "        status = Trim$(mCSV.GetField(id, ""Resolved""))" & vbCrLf
    s = s & "        If Me.cboStatus.Value = ""Open"" And StrComp(status, ""Yes"", vbTextCompare) = 0 Then GoTo SkipRow" & vbCrLf
    s = s & "        If Me.cboStatus.Value = ""Resolved"" And StrComp(status, ""Yes"", vbTextCompare) <> 0 Then GoTo SkipRow" & vbCrLf
    s = s & "        If cutoff > 0 Then" & vbCrLf
    s = s & "            If Not ParseCSVDate(mCSV.GetField(id, ""Date of Defect""), defectDate) Then GoTo SkipRow" & vbCrLf
    s = s & "            If defectDate < cutoff Or defectDate > Date Then GoTo SkipRow" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If panel = ""Day"" Then" & vbCrLf
    s = s & "            If Not ParseCSVDate(mCSV.GetField(id, ""Date of Defect""), defectDate) Then GoTo SkipRow" & vbCrLf
    s = s & "            If Format$(defectDate, ""yyyy-mm-dd"") <> selected Then GoTo SkipRow" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            raw = Trim$(mCSV.GetField(id, fieldName))" & vbCrLf
    s = s & "            If Len(raw) = 0 Then raw = ""(Blank)""" & vbCrLf
    s = s & "            If StrComp(raw, selected, vbTextCompare) <> 0 Then GoTo SkipRow" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        results.Add id" & vbCrLf
    s = s & "SkipRow:" & vbCrLf
    s = s & "    Next k" & vbCrLf
    s = s & "    Set detail = VBA.UserForms.Add(""frmDefectDrillDown"")" & vbCrLf
    s = s & "    detail.LoadResults mCSV, results, panel & "" : "" & selected" & vbCrLf
    s = s & "    detail.Show vbModal" & vbCrLf
    s = s & "    Unload detail" & vbCrLf
    s = s & "End Sub" & vbCrLf
    DashboardDrillCode = s
End Function

Private Function DetailFormCode() As String
    Dim s As String
    s = s & "Private mCSV As CSVEngine" & vbCrLf
    s = s & "Private mIDs As Collection" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Me.Caption = ""Defect Drill-Down""" & vbCrLf
    s = s & "    Me.StartUpPosition = 1" & vbCrLf
    s = s & "    Me.BackColor = RGB(248, 250, 254)" & vbCrLf
    s = s & "    Me.Width = 840" & vbCrLf
    s = s & "    Me.Height = 570" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub LoadResults(ByVal engine As CSVEngine, ByVal ids As Collection, ByVal heading As String)" & vbCrLf
    s = s & "    Dim id As Variant, n As Long, resolved As String" & vbCrLf
    s = s & "    Set mCSV = engine" & vbCrLf
    s = s & "    Set mIDs = ids" & vbCrLf
    s = s & "    Me.Caption = ""Defect Drill-Down - "" & heading" & vbCrLf
    s = s & "    Me.lblDrillTitle.Caption = heading" & vbCrLf
    s = s & "    Me.lblDrillCount.Caption = CStr(ids.Count) & "" matching defect(s) - select a row for full details""" & vbCrLf
    s = s & "    Me.lstDrillRecords.Clear" & vbCrLf
    s = s & "    For Each id In ids" & vbCrLf
    s = s & "        n = Me.lstDrillRecords.ListCount" & vbCrLf
    s = s & "        Me.lstDrillRecords.AddItem CStr(id)" & vbCrLf
    s = s & "        Me.lstDrillRecords.List(n, 1) = mCSV.GetField(CStr(id), ""Date of Defect"")" & vbCrLf
    s = s & "        Me.lstDrillRecords.List(n, 2) = mCSV.GetField(CStr(id), ""z number"")" & vbCrLf
    s = s & "        Me.lstDrillRecords.List(n, 3) = mCSV.GetField(CStr(id), ""Name"")" & vbCrLf
    s = s & "        Me.lstDrillRecords.List(n, 4) = mCSV.GetField(CStr(id), ""Defect"")" & vbCrLf
    s = s & "        resolved = Trim$(mCSV.GetField(CStr(id), ""Resolved""))" & vbCrLf
    s = s & "        If StrComp(resolved, ""Yes"", vbTextCompare) = 0 Then" & vbCrLf
    s = s & "            Me.lstDrillRecords.List(n, 5) = ""Resolved""" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            Me.lstDrillRecords.List(n, 5) = ""Open""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next id" & vbCrLf
    s = s & "    If ids.Count > 0 Then" & vbCrLf
    s = s & "        Me.lstDrillRecords.ListIndex = 0" & vbCrLf
    s = s & "        ShowSelected" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Me.txtDrillDetails.Value = ""No matching records.""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstDrillRecords_Click()" & vbCrLf
    s = s & "    ShowSelected" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ShowSelected()" & vbCrLf
    s = s & "    Dim i As Long, id As String, fields As Variant, field As Variant" & vbCrLf
    s = s & "    Dim s As String" & vbCrLf
    s = s & "    i = Me.lstDrillRecords.ListIndex" & vbCrLf
    s = s & "    If i < 0 Then Exit Sub" & vbCrLf
    s = s & "    id = CStr(Me.lstDrillRecords.List(i, 0))" & vbCrLf
    s = s & "    fields = Array(""Submission ID"", ""Submitted At"", ""Date of Defect"", _" & vbCrLf
    s = s & "        ""z number"", ""Name"", ""Key"", ""Supervisor"", ""Defect"", ""Sub Type"", _" & vbCrLf
    s = s & "        ""Impact Label Calc"", ""Description Template"", ""Additional Details"", _" & vbCrLf
    s = s & "        ""System Location"", ""System Count"", ""Physical Count"", _" & vbCrLf
    s = s & "        ""System DPCI"", ""Physical DPCI"", ""System PID"", ""Physical PID"", _" & vbCrLf
    s = s & "        ""System VCP/SSP"", ""Physical VCP/SSP"", ""Impacted Labels"", _" & vbCrLf
    s = s & "        ""Verification Method"", ""Over/Under"", ""Physical Location"", _" & vbCrLf
    s = s & "        ""Door"", ""Put Method"", ""Image 1 Path"", ""Image 2 Path"", ""Resolved"")" & vbCrLf
    s = s & "    For Each field In fields" & vbCrLf
    s = s & "        s = s & CStr(field) & Chr$(58) & Chr$(32) & mCSV.GetField(id, CStr(field)) & vbCrLf" & vbCrLf
    s = s & "    Next field" & vbCrLf
    s = s & "    Me.txtDrillDetails.Value = s" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdDrillClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    DetailFormCode = s
End Function
