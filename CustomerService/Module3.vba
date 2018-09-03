Dim chartsPath As String
Dim reportPath As String

Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Sub setCharts(wsReport As Worksheet, wsHours As Worksheet, colHours As Integer, wsCharts As Worksheet, reportCol As Integer, title As String, Optional wsReponsabilities As Worksheet)

    For r = 1 To wsReport.UsedRange.Rows.Count
        If wsReport.Cells(r, 1).Value = "" Then
            Exit For
        End If
        If Not wsReponsabilities Is Nothing Then
            wsCharts.Cells(r, 1).Value = wsReport.Cells(r, 1).Value & " " & wsReponsabilities.Cells(r, 5).Value
        Else: wsCharts.Cells(r, 1).Value = wsReport.Cells(r, 1).Value
        End If
        wsCharts.Cells(r, 3).Value = wsHours.Cells(r, colHours).Text
        wsCharts.Cells(r, 4).Value = wsReport.Cells(r, reportCol).Text
        If wsCharts.Cells(r, 3).Value = 0 Then
            wsCharts.Cells(r, 2).Value = 0
        Else
            wsCharts.Cells(r, 2).Value = "=D" & r & "/C" & r
        End If
        wsCharts.Cells(r, 2).NumberFormat = "0.#"
    Next r
    
    Dim lastrow As Long
    lastrow = wsCharts.Cells(wsCharts.Rows.Count, 2).End(xlUp).row
    wsCharts.Cells(1, 2).Value = title
    wsCharts.Cells(lastrow + 1, 1).Value = "Avg"
    wsCharts.Cells(lastrow + 1, 2).Value = Application.WorksheetFunction.Average(wsCharts.Range("B2:B" & lastrow + 1))
    wsCharts.Cells(lastrow + 1, 2).NumberFormat = "0.#"
    
    wsCharts.Rows(1).Font.Bold = True
    wsCharts.Columns("A:D").AutoFit
    wsCharts.Range("A1:Z1").HorizontalAlignment = xlCenter
    
    lastrow = wsCharts.Cells(wsCharts.Rows.Count, 2).End(xlUp).row
    wsCharts.Range("A2:D" & lastrow).Sort key1:=wsCharts.Range("B2:B" & lastrow), _
    order1:=xlAscending, Header:=xlNo
    
    Dim Chrt As Chart
    wsCharts.Shapes.AddChart
    Set Chrt = wsCharts.ChartObjects(1).Chart
    Chrt.SetSourceData Source:=wsCharts.Range("'" & wsCharts.name & "'!$A$2:$B$" & lastrow)
    Chrt.ChartType = xlColumnClustered
    Chrt.Axes(xlCategory).TickLabels.Orientation = 50
    Chrt.Legend.Delete
    Chrt.SeriesCollection(1).ApplyDataLabels
    Chrt.Parent.Left = wsCharts.Range("F4").Left
    Chrt.Parent.Top = wsCharts.Range("F4").Top
    Chrt.Parent.Width = 600
    Chrt.HasTitle = True
    Chrt.ChartTitle.Text = title
    
    avgRow = 0
    
    For r = 2 To wsCharts.UsedRange.Rows.Count
        If wsCharts.Cells(r, 1).Value = "Avg" Then
            avgRow = r
            Exit For
        End If
    Next r
    
    If avgRow > 0 Then
        Chrt.SeriesCollection(1).Points(avgRow - 1).Interior.Color = RGB(255, 0, 0)
    End If

End Sub

Function getColNum(ws As Worksheet, colName As String) As Integer

    For c = 1 To ws.UsedRange.Columns.Count
        If ws.Cells(1, c).Value = colName Then
            getColNum = c
            Exit Function
        End If
    Next c

End Function

Sub removeEntries(ws As Worksheet, agentName As String, lastrow As Long, deleteSheet As Boolean)

    For r2 = 2 To lastrow
        If deleteSheet = True And ws.Cells(r2, 1).Value = agentName And ws.Cells(r2, 2).Value = 0 Then
            ws.Delete
            Exit Sub
        End If
        If r2 < lastrow - 9 And agentName <> ws.Cells(r2, 1).Text And ws.Cells(r2, 1).Text <> "Avg" Then
            ws.Cells(r2, 1).Value = ""
        End If
    Next r2

End Sub

Sub checkRemoveSheet(ws As Worksheet, agentName As String, lastrow As Long)

    For r2 = 2 To lastrow
        If ws.Cells(r2, 1).Value = agentName Then
            If ws.Cells(r2, 2).Value = 0 Then
                ws.Delete
            End If
            Exit Sub
        End If
    Next r2

End Sub

Sub Button2_Click()

    Dim ws As Worksheet
    Dim wbCharts As Workbook
    Dim wsCharts1 As Worksheet
    Dim wsCharts2 As Worksheet
    Dim wsCharts3 As Worksheet
    Dim wsCharts4 As Worksheet
    Dim wbReport As Workbook
    Dim wsReport1 As Worksheet
    Dim wsReport2 As Worksheet
    Dim wbChartWorker As Workbook
    Dim wsChartWorker As Worksheet
    Dim wsChartWorker2 As Worksheet
    Dim wsChartWorker3 As Worksheet
    Dim wsChartWorker4 As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    path = ThisWorkbook.path & "\"
    timeStamp = Format(Now(), "mmddyyyyhhmm")
    reportPath = path + ws.Cells(2, 8).Value
    chartsPath = path + "Charts_" & timeStamp & ".xls"
    
    Workbooks.Add
    ActiveWorkbook.SaveAs chartsPath
    Set wbCharts = Application.Workbooks.Open(chartsPath)
    wbCharts.Worksheets.Add After:=wbCharts.Worksheets(wbCharts.Worksheets.Count)
    wbCharts.Worksheets.Add After:=wbCharts.Worksheets(wbCharts.Worksheets.Count)
    Set wsCharts1 = wbCharts.Worksheets(1)
    Set wsCharts2 = wbCharts.Worksheets(2)
    Set wsCharts3 = wbCharts.Worksheets(3)
    Set wbReport = Application.Workbooks.Open(reportPath)
    Set wsReport1 = wbReport.Worksheets(1)
    Set wsReport2 = wbReport.Worksheets(2)
    
    wsCharts1.name = "Activity"
    Call setCharts(wsReport1, wsReport1, 3, wsCharts1, 19, "Total Activity per Hour " & ws.Cells(2, 7).Value, ws)
    wsCharts2.name = "Donations"
    Call setCharts(wsReport1, wsReport1, 3, wsCharts2, 15, "Donations per Hour")
    wsCharts3.name = "Rush Pickup"
    Call setCharts(wsReport1, wsReport2, getColNum(wsReport2, "Rush Pickup"), wsCharts3, 19, "Rush P/U per Hour")
    wbCharts.Worksheets.Add After:=wbCharts.Worksheets(wbCharts.Worksheets.Count)
    Set wsCharts4 = wbCharts.Worksheets(4)
    wsCharts4.name = "Escalated Issues"
    Call setCharts(wsReport1, wsReport2, getColNum(wsReport2, "Escalated Issues"), wsCharts4, 18, "Escalated Issues per Hour")
    
    wsCharts1.Activate
    
    wbCharts.Save
    
    pathxxx = ThisWorkbook.path + "\Charts_" & timeStamp & "\"
    MkDir pathxxx
    Dim lastrow As Long
    lastrow = wsReport1.Cells(wsReport1.Rows.Count, 2).End(xlUp).row
    Dim agentName As String
    For r = 2 To lastrow
        agentName = wsReport1.Cells(r, 1).Text
        If agentName <> "Avg" Then
            chartsPathxxx = pathxxx + agentName & "_Chart.xls"
            Dim xlobj As Object
            Set xlobj = CreateObject("Scripting.FileSystemObject")
            xlobj.CopyFile chartsPath, chartsPathxxx, True
            
            Set wbChartWorker = Application.Workbooks.Open(chartsPathxxx)
            Set wsChartWorker = wbChartWorker.Worksheets(1)
            Set wsChartWorker2 = wbChartWorker.Worksheets(2)
            Set wsChartWorker3 = wbChartWorker.Worksheets(3)
            Set wsChartWorker4 = wbChartWorker.Worksheets(4)
            Worksheets.Add After:=Worksheets(Worksheets.Count)
            Worksheets.Add After:=Worksheets(Worksheets.Count)
            'Worksheets.Add After:=Worksheets(Worksheets.Count)
            Set wsChartWorker5 = wbChartWorker.Worksheets(5)
            Set wsChartWorker6 = wbChartWorker.Worksheets(6)
            'Set wsChartWorker7 = wbChartWorker.Worksheets(7)
            wsChartWorker5.name = "Statuses1"
            wsChartWorker6.name = "Statuses2"
            'wsChartWorker7.name = "AbandonedCalls"
            
            Application.DisplayAlerts = False

            'Call removeEntries(wsChartWorker, agentName, lastrow, False)
            'Call removeEntries(wsChartWorker2, agentName, lastrow, False)
            'Call removeEntries(wsChartWorker3, agentName, lastrow, True)
            'Call removeEntries(wsChartWorker4, agentName, lastrow, True)
            
            Call checkRemoveSheet(wsChartWorker3, agentName, lastrow)
            Call checkRemoveSheet(wsChartWorker4, agentName, lastrow)
            
            Application.DisplayAlerts = True
            
            'sheet 5: Statuses1
            
            wsChartWorker5.Cells(1, 1).Value = agentName
            wsChartWorker5.Cells(1, 4).Value = ws.Cells(2, 7).Value
            
            wsChartWorker5.Cells(3, 1).Value = wsReport1.Cells(1, 2).Value
            wsChartWorker5.Cells(4, 1).Value = wsReport1.Cells(r, 2).Value
            
            wsChartWorker5.Cells(3, 2).Value = wsReport2.Cells(1, 3).Value + " TV"
            wsChartWorker5.Cells(4, 2).Value = wsReport2.Cells(r, 3).Value
            
            Dim nextCol As Integer
            nextCol = 3
            For c = 4 To wsReport2.UsedRange.Columns.Count
                statVal = wsReport2.Cells(r, c).Value
                If statVal <> "" Then
                    wsChartWorker5.Cells(3, nextCol).Value = wsReport2.Cells(1, c).Value
                    wsChartWorker5.Cells(4, nextCol).Value = statVal / wsChartWorker5.Cells(4, 2).Value
                    wsChartWorker5.Cells(4, nextCol).Style = "Percent"
                    nextCol = nextCol + 1
                End If
            Next c
            
            wsChartWorker5.Rows(1).Font.Bold = True
            wsChartWorker5.Rows(3).Font.Bold = True
            
            wsChartWorker5.Shapes.AddChart
            Set Chrt = wsChartWorker5.ChartObjects(1).Chart
            Chrt.SetSourceData Source:=wsChartWorker5.Range("'Statuses1'!$C$3:$" & ConvertToLetter(nextCol - 1) & "$4")
            Chrt.ChartType = xlPie
            Chrt.PlotBy = xlColumns
            Chrt.PlotBy = xlRows
            Chrt.SeriesCollection(1).ApplyDataLabels
            
            Chrt.Parent.Left = wsChartWorker5.Range("B7").Left
            Chrt.Parent.Top = wsChartWorker5.Range("B7").Top
            Chrt.HasTitle = True
            Chrt.ChartTitle.Text = "TV Statuses"
            
            'sheet 6: Statuses2
            
            NextRow = 1
            For c = 1 To wsReport1.UsedRange.Columns.Count
                Select Case c
                    Case 2, 3, 6, 7, 8, 9, 10, 12, 13, 18
                    Case Else
                        If InStr(wsReport1.Cells(1, c).Value, "Auction Orders .") <= 0 Then
                            wsChartWorker6.Cells(NextRow, 1).Value = wsReport1.Cells(1, c).Value
                            wsChartWorker6.Cells(NextRow, 2).Value = wsReport1.Cells(r, c).Text
                            NextRow = NextRow + 1
                        End If
                End Select
            Next c
            wsChartWorker6.Cells(1, 5).Value = ws.Cells(2, 7).Value
            wsChartWorker6.Rows(1).Font.Bold = True
            wsChartWorker6.Columns(1).Font.Bold = True
            wsChartWorker6.Columns("A").AutoFit
            
            'sheeet 7: AbandonedCalls
            
            'wsChartWorker7.Cells(1, 1).Value = agentName
            'wsChartWorker7.Cells(1, 4).Value = ws.Cells(2, 7).Value
            
            'colAbandoned = getColNum(wsReport1, "Abandoned Calls")
            
            'wsChartWorker7.Cells(3, 1).Value = wsReport1.Cells(1, 4).Value
            'wsChartWorker7.Cells(4, 1).Value = wsReport1.Cells(r, 4).Value
            'wsChartWorker7.Cells(3, 2).Value = wsReport1.Cells(1, colAbandoned).Value
            'wsChartWorker7.Cells(4, 2).Value = wsReport1.Cells(r, colAbandoned).Value
            
            'wsChartWorker7.Rows(1).Font.Bold = True
            'wsChartWorker7.Rows(3).Font.Bold = True
            
            'wsChartWorker7.Shapes.AddChart
            'Set Chrt = wsChartWorker7.ChartObjects(1).Chart
            'Chrt.SetSourceData Source:=wsChartWorker7.Range("'AbandonedCalls'!$A$3:$B$4")
            'Chrt.ChartType = xlPie
            'Chrt.PlotBy = xlColumns
            'Chrt.PlotBy = xlRows
            'Chrt.SeriesCollection(1).ApplyDataLabels
            
            'Chrt.Parent.Left = wsChartWorker7.Range("B7").Left
            'Chrt.Parent.Top = wsChartWorker7.Range("B7").Top
            'Chrt.HasTitle = True
            'Chrt.ChartTitle.Text = "Abandoned Calls"
            
            
            wsChartWorker.Activate
            
            With wbChartWorker
                .Save
                .Close
            End With
        End If
    Next r
    
    wbReport.Close False
    wbCharts.Close False
    
    MsgBox "Generated Report File: " & chartsPath
    
End Sub


