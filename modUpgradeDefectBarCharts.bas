Attribute VB_Name = "modUpgradeDefectBarCharts"
Option Explicit
' Run UpgradeDefectBarCharts once, with the dashboard closed.
Public Sub UpgradeDefectBarCharts()
    Dim comp As Object, d As Object, cm As Object
    Dim i As Long, panel As Variant, x As Single, y As Single
    Dim atLine As Long, startLine As Long, endLine As Long
    Dim source As String
    On Error GoTo Failed
    Set comp = ThisWorkbook.VBProject.VBComponents("frmDefectDashboard")
    Set d = comp.designer
    Set cm = comp.CodeModule
    If cm.Find("Private Sub RefreshBarCharts()", 1, 1, cm.CountOfLines, 500) Then
        MsgBox "Bar charts are already installed.", vbInformation: Exit Sub
    End If
    For Each panel In Array("Type", "Key", "Supervisor", "Day")
        Select Case CStr(panel)
            Case "Type": x = 18: y = 246
            Case "Key": x = 382: y = 246
            Case "Supervisor": x = 18: y = 406
            Case "Day": x = 382: y = 406
        End Select
        For i = 0 To 7
            MakeChartLabel d, "bc" & panel & "Name" & i, x, y + i * 15, 116, 14, RGB(45, 55, 72), 8
            MakeChartLabel d, "bc" & panel & "Bar" & i, x + 120, y + i * 15 + 2, 1, 10, vbWhite, 8
            MakeChartLabel d, "bc" & panel & "Value" & i, x + 284, y + i * 15, 44, 14, RGB(45, 55, 72), 8
        Next i
    Next panel
    d.Controls("lstByType").Visible = False
    d.Controls("lstByKey").Visible = False
    d.Controls("lstBySupervisor").Visible = False
    d.Controls("lstByDay").Visible = False
    startLine = cm.ProcStartLine("RefreshDashboard", 0)
    endLine = startLine + cm.ProcCountLines("RefreshDashboard", 0) - 1
    For i = endLine To startLine Step -1
        If Trim$(cm.Lines(i, 1)) = "End Sub" Then
            cm.InsertLines i, "    RefreshBarCharts"
            Exit For
        End If
    Next i
    If i < startLine Then Err.Raise 5, , "Could not find RefreshDashboard end."
    cm.AddFromString ChartCode()
    MsgBox "Bar charts installed. Reopen the dashboard to view them.", vbInformation
    Exit Sub
Failed:
    MsgBox "Bar chart upgrade failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub MakeChartLabel(ByVal d As Object, ByVal nm As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, _
    ByVal h As Single, ByVal fore As Long, ByVal fs As Single)
    Dim c As Object
    Set c = d.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Left = x: .Top = y: .Width = w: .Height = h
        .caption = "": .Font.Size = fs: .ForeColor = fore
        .BackStyle = 0: .WordWrap = False
    End With
End Sub

Private Function ChartCode() As String
    Dim s As String
    s = s & "Private Sub RefreshBarCharts()" & vbCrLf
    s = s & "    RenderBarPanel Me.lstByType, ""Type"", RGB(45, 115, 205), False" & vbCrLf
    s = s & "    RenderBarPanel Me.lstByKey, ""Key"", RGB(17, 151, 120), False" & vbCrLf
    s = s & "    RenderBarPanel Me.lstBySupervisor, ""Supervisor"", RGB(137, 85, 190), False" & vbCrLf
    s = s & "    RenderBarPanel Me.lstByDay, ""Day"", RGB(225, 132, 39), True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RenderBarPanel(ByVal lb As MSForms.ListBox, _" & vbCrLf
    s = s & "    ByVal prefix As String, ByVal barColor As Long, _" & vbCrLf
    s = s & "    ByVal latestDates As Boolean)" & vbCrLf
    s = s & "    Dim i As Long, idx As Long, n As Long, maxVal As Long" & vbCrLf
    s = s & "    Dim val As Long, widthVal As Single, startIdx As Long" & vbCrLf
    s = s & "    Dim nameLbl As MSForms.Label, barLbl As MSForms.Label" & vbCrLf
    s = s & "    Dim valueLbl As MSForms.Label, labelText As String" & vbCrLf
    s = s & "    n = lb.ListCount" & vbCrLf
    s = s & "    If latestDates And n > 8 Then startIdx = n - 8 Else startIdx = 0" & vbCrLf
    s = s & "    For i = startIdx To n - 1" & vbCrLf
    s = s & "        val = CLng(lb.List(i, 1))" & vbCrLf
    s = s & "        If val > maxVal Then maxVal = val" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    For i = 0 To 7" & vbCrLf
    s = s & "        Set nameLbl = Me.Controls(""bc"" & prefix & ""Name"" & i)" & vbCrLf
    s = s & "        Set barLbl = Me.Controls(""bc"" & prefix & ""Bar"" & i)" & vbCrLf
    s = s & "        Set valueLbl = Me.Controls(""bc"" & prefix & ""Value"" & i)" & vbCrLf
    s = s & "        idx = startIdx + i" & vbCrLf
    s = s & "        If idx < n Then" & vbCrLf
    s = s & "            labelText = CStr(lb.List(idx, 0))" & vbCrLf
    s = s & "            If Len(labelText) > 21 Then labelText = Left$(labelText, 18) & ""...""" & vbCrLf
    s = s & "            nameLbl.Caption = labelText" & vbCrLf
    s = s & "            val = CLng(lb.List(idx, 1))" & vbCrLf
    s = s & "            valueLbl.Caption = CStr(val)" & vbCrLf
    s = s & "            If maxVal > 0 Then" & vbCrLf
    s = s & "                widthVal = 155! * val / maxVal" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                widthVal = 0" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            barLbl.Width = IIf(widthVal < 1, 1, widthVal)" & vbCrLf
    s = s & "            barLbl.BackStyle = 1" & vbCrLf
    s = s & "            barLbl.BackColor = barColor" & vbCrLf
    s = s & "            barLbl.Visible = (val > 0)" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            nameLbl.Caption = """"" & vbCrLf
    s = s & "            valueLbl.Caption = """"" & vbCrLf
    s = s & "            barLbl.Visible = False" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    ChartCode = s
End Function
