Dim tmpReportPath As String
Dim templatePath As String
Dim finalReportPath As String
Dim kronosPath As String
Dim ticketsNewPath As String
Dim ticketsClosedPath As String
Dim ticketsRepliesPath As String
Dim k4kPath As String
Dim leadsPath As String
Dim auctionPath As String
Dim fuPath As String
Dim statPath As String

Sub updateTemplateParam(name As String, rangeParam As String)

    Dim template As Workbook
    Dim options As Worksheet

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        Set template = .Workbooks.Open(templatePath)
        Set options = template.Worksheets(3)
    End With
    
    options.Cells(1, 3).Value = name
    options.Cells(2, 2).Value = rangeParam
    
    With template
        .Save
        .Close
    End With

End Sub

Sub createReport()

    Workbooks.Add
    ActiveWorkbook.SaveAs finalReportPath
    
    Dim finalReport As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim names As Worksheet
    Dim statuses As Worksheet
    
    Set finalReport = Application.Workbooks.Open(finalReportPath)
    Set ws1 = finalReport.Worksheets(1)
    
    'Dim colHeaders
    'Set colHeaders = Array("Team Member", "Kronos Hours", "Hours minus statuses", "Inbound Calls", "Outbound Calls", "Outbound Calls (.75 pts)", _
    '"Inbound Emails (.25 pts)", "Inbound Emails (.5 pts)", "Inbound Emails (.75 pts)", "Inbound Emails (1 pts)", "Inbound Emails", _
    '"Outbound Emails (.75 pts)", "Outbound Emails", "Closed Emails", "Chats", "Coparts Entered", "Coparts Entered (.40 pts)", "Total", _
    '"Donations", "Leads (not donations)", "Auction Orders", "Escalated Issues", "Arrange Pickup/Rush Pickup")
    
    'For i = 1 To UBound(colHeaders)
        'ws1.Cells(1, 1).Value = colHeaders(i)
    'Next i
    ws1.Cells(1, 1).Value = "Team Member"
    ws1.Cells(1, 2).Value = "Kronos Hours"
    ws1.Cells(1, 3).Value = "Hours minus statuses"
    ws1.Cells(1, 4).Value = "Inbound Calls"
    ws1.Cells(1, 5).Value = "Outbound Calls"
    ws1.Cells(1, 6).Value = "Outbound Calls (.75 pts)"
    ws1.Cells(1, 7).Value = "Inbound Emails (.25 pts)"
    ws1.Cells(1, 8).Value = "Inbound Emails (.5 pts)"
    ws1.Cells(1, 9).Value = "Inbound Emails (.75 pts)"
    ws1.Cells(1, 10).Value = "Inbound Emails (1 pts)"
    ws1.Cells(1, 11).Value = "Inbound Emails - Total"
    ws1.Cells(1, 12).Value = "Inbound Emails - pts"
    ws1.Cells(1, 13).Value = "Outbound Emails (.75 pts)"
    ws1.Cells(1, 14).Value = "Outbound Emails"
    ws1.Cells(1, 15).Value = "Closed Emails"
    ws1.Cells(1, 16).Value = "Chats"
    ws1.Cells(1, 17).Value = "Coparts Entered"
    ws1.Cells(1, 18).Value = "Coparts Entered (.40 pts)"
    ws1.Cells(1, 19).Value = "Total"
    ws1.Cells(1, 20).Value = "Donations"
    ws1.Cells(1, 21).Value = "Leads (not donations)"
    ws1.Cells(1, 22).Value = "Auction Orders"
    ws1.Cells(1, 23).Value = "Escalated Issues"
    
    ws1.Cells(1, 24).Value = "-Arrange Pickup/Rush Pickup"
    
    ws1.Rows(1).Font.Bold = True
    ws1.Range("A1:U1").HorizontalAlignment = xlCenter
    ws1.Columns("A:BZ").AutoFit
    
    Set names = ThisWorkbook.Worksheets(1)
    For r = 2 To names.UsedRange.Rows.Count
        ws1.Cells(r, 1).Value = names.Cells(r, 1).Value
    Next r
    
    finalReport.Worksheets.Add After:=finalReport.Worksheets(finalReport.Worksheets.Count)
    finalReport.Worksheets.Add After:=finalReport.Worksheets(finalReport.Worksheets.Count)
    Set ws2 = finalReport.Worksheets(2)
    ws2.Cells(1, 1).Value = "Agent"
    ws2.Cells(1, 2).Value = "Avg Call Inbound"
    ws2.Cells(1, 3).Value = "Total"
    
    Set statuses = ThisWorkbook.Worksheets(3)
    For r = 2 To statuses.UsedRange.Rows.Count
        ws2.Cells(1, r + 2).Value = statuses.Cells(r, 1).Value
    Next r
    
    ws2.Rows(1).Font.Bold = True
    ws2.Range("A1:AZ1").HorizontalAlignment = xlCenter
    
    With finalReport
        .Save
        .Close
    End With
    
End Sub

Sub createTvReport()

    Workbooks.Add
    ActiveWorkbook.SaveAs finalReportPath
    
    Dim finalReport As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim names As Worksheet
    Dim statuses As Worksheet
    
    Set finalReport = Application.Workbooks.Open(finalReportPath)
    Set ws1 = finalReport.Worksheets(1)
    
    ws1.Cells(1, 1).Value = "Agent"
    ws1.Cells(1, 2).Value = "Avg Call Inbound"
    ws1.Cells(1, 3).Value = "Total"
    
    ws1.Rows(1).Font.Bold = True
    ws1.Range("A1:AZ1").HorizontalAlignment = xlCenter
    
    With finalReport
        .Save
        .Close
    End With
    
End Sub

Sub setTotals(name As String, ws As Worksheet, row As Integer, tmpReportWs As Worksheet, totalsRow As Integer)
    
    ws.Cells(row, 1).Value = name
    
    If totalsRow > 0 Then
        ws.Cells(row, 4).Value = tmpReportWs.Cells(totalsRow, 2).Value
        ws.Cells(row, 5).Value = tmpReportWs.Cells(totalsRow, 3).Value
        ws.Cells(row, 6).Value = ws.Cells(row, 5).Value * 0.75
    End If
    
End Sub

Function getRange(ws As Worksheet, col As String) As Range
    rangeStr = col & "2:" & col & ws.UsedRange.Rows.Count
    Set getRange = ws.Range(rangeStr)
End Function

Function getCompareableStatus(status As String)
    getCompareableStatus = Replace(Replace(Replace(Replace(UCase(status), "'", ""), " ", ""), "/", ""), "\", "")
End Function

Function compare(str1 As String, str2 As String) As Boolean
    
    cmpr1 = getCompareableStatus(str1)
    cmpr2 = getCompareableStatus(str2)
        
    If cmpr1 = cmpr2 Or cmpr1 & "S" = cmpr2 Or cmpr1 = cmpr2 & "S" Then
        compare = True
    Else: compare = False
    End If
End Function

Sub dateToDouble(ws As Worksheet, row As Integer, col As Integer, statusVal, timeFormat As String)

    ws.Cells(row, col).Value = statusVal
    ws.Cells(row, col).NumberFormat = timeFormat
    dd = Val(ws.Cells(row, col).Text)
    If dd <> 0 Then
        ws.Cells(row, col).NumberFormat = "0.##"
        ws.Cells(row, col).Value = dd
    Else
        ws.Cells(row, col).Value = ""
    End If
    
End Sub

Sub setStatuses(name As String, ws As Worksheet, row As Integer, tmpReportWs As Worksheet, statusesRow As Integer, statusesTotalsRow As Integer, totalsRow As Integer)

    Dim statuses As Worksheet
    Dim status As String
    Dim flag As Boolean
    Dim index As Integer
    Dim nextIndex As Integer
    Dim sumStatuses As Integer
    Dim cell2 As Integer
    
    Set statuses = ThisWorkbook.Worksheets(3)
    ws.Cells(row, 1).Value = name
    If totalsRow > 0 Then
        Call dateToDouble(ws, row, 2, tmpReportWs.Cells(totalsRow, 8), "m.ss")
    End If
    
    If statusesTotalsRow > 0 Then
    
        For cell1 = 2 To tmpReportWs.Columns.Count
            flag = False
            status = tmpReportWs.Cells(statusesRow, cell1).Value
            If status <> "" And status <> "Out Of The Office" And status <> "On Vacation" Then
                For cell2 = 4 To ws.Columns.Count
                    If ws.Cells(1, cell2).Value = "" Then
                        Exit For
                    ElseIf compare(ws.Cells(1, cell2).Value, status) Then
                        Call dateToDouble(ws, row, cell2, tmpReportWs.Cells(statusesTotalsRow, cell1).Value, "[h].mm")
                        flag = True
                        Exit For
                    End If
                Next cell2
                If flag = False Then
                    ws.Cells(1, cell2).Value = status
                    Call dateToDouble(ws, row, cell2, tmpReportWs.Cells(statusesTotalsRow, cell1).Value, "[h].mm")
                End If
            End If
        Next cell1
        
        ws.Cells(row, 3).Value = "=SUM(D" & row & ":BZ" & row & ")"
        
        ws.Columns("A:AZ").AutoFit
    End If

End Sub

Function getAddress(ws As Worksheet, row As Integer, col As Integer) As String

    getAddress = ws.Cells(row, col).Address(False, False)

End Function

Sub setTeleVantage(finalReport As Workbook, row As Integer, name As String, statusesOnly As Boolean)

    Dim statuses As Worksheet
    Dim tmpReportWs As Worksheet
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim totalsRow As Integer
    Dim statusesRow As Integer
    Dim statusesTotalsRow As Integer
    Dim hoursStr As String
    
    Set tmpReportWs = Application.Workbooks.Open(tmpReportPath).Sheets(2)
    
    For r = 2 To tmpReportWs.UsedRange.Rows.Count
        If tmpReportWs.Cells(r, 1).Value = "Totals" Then
            totalsRow = r
            Exit For
        End If
    Next r
    
    If Not statusesOnly = True Then
        Set ws1 = finalReport.Worksheets(1)
        Call setTotals(name, ws1, row, tmpReportWs, totalsRow)
    End If
    
    For r = totalsRow + 1 To tmpReportWs.UsedRange.Rows.Count
        If tmpReportWs.Cells(r, 1).Value = "Date" Then
            statusesRow = r
        End If
        If tmpReportWs.Cells(r, 1).Value = "Totals" Then
            statusesTotalsRow = r
        End If
    Next r
    
    If statusesOnly = True Then
        Set ws2 = finalReport.Worksheets(1)
    Else
        Set ws2 = finalReport.Worksheets(2)
    End If
    
    Call setStatuses(name, ws2, row, tmpReportWs, statusesRow, statusesTotalsRow, totalsRow)
    
    If Not statusesOnly = True Then
        Set statuses = ThisWorkbook.Worksheets(3)
        hoursStr = "=B" & row
        For r = 2 To statuses.UsedRange.Rows.Count
            If statuses.Cells(r, 2).Value = False Then
                hoursStr = hoursStr & "-Sheet2!" & ws2.Cells(row, r + 2).Address(False, False)
            End If
        Next r
        
        ws1.Cells(row, 3).Value = hoursStr
    End If
    
    tmpReportWs.Parent.Close
    
End Sub

Sub runTVRRun()

    Call Shell(Chr(34) & "C:\Program Files (x86)\TeleVantage\Client\Reporter\TVRRun.exe" & Chr(34) & " " & templatePath & " -S " & tmpReportPath)
    
End Sub

Function getNumOfExcelProcesses() As Integer

    Dim objProcess, process, strNameOfUser
    ComputerName = "."
    numOfExcels = 0
    Set objProcess = GetObject("winmgmts:{impersonationLevel=impersonate}\\" _
          & ComputerName & "\root\cimv2").ExecQuery("Select * From Win32_Process")
    For Each process In objProcess
        If process.name = "EXCEL.EXE" Then
            numOfExcels = numOfExcels + 1
        End If
    Next
    
    Set objProcess = Nothing
    getNumOfExcelProcesses = numOfExcels
    
End Function

Sub runTVRRunAndWait()
    
    Dim numOfExcels As Integer
    numOfExcels = getNumOfExcelProcesses
    Call runTVRRun
    
    Do While getNumOfExcelProcesses > numOfExcels
        Application.Wait (Now + TimeValue("00:00:02"))
    Loop
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
End Sub

Sub setKronosHours()

    Dim ws As Worksheet
    Dim kronos As Worksheet
    Dim report As Worksheet
    Dim r As Integer
    Dim NameArray() As String
    Dim falgName As Boolean
    Dim cntPayedDays As Double
    
    With Application
        Set kronos = .Workbooks.Open(kronosPath).Sheets(1)
        Set report = .Workbooks.Open(finalReportPath).Sheets(1)
        Set ws = ThisWorkbook.Worksheets(1)
    End With
    
    For rReport = 2 To report.UsedRange.Rows.Count
        falgName = False
        cntPayedDays = 0
        NameArray() = Split(ws.Cells(rReport, 2).Value)
        
        For rKronos = 2 To kronos.UsedRange.Rows.Count
            If kronos.Cells(rKronos, 2).Value = NameArray(1) And kronos.Cells(rKronos - 1, 2).Value = NameArray(0) Then
                falgName = True
            End If
            If falgName = True And kronos.Cells(rKronos, 1).Value = "Subtotal" Then
                report.Cells(rReport, 2).Value = kronos.Cells(rKronos, 3).Value - cntPayedDays
                report.Cells(rReport, 2).NumberFormat = "0.##"
                Exit For
            End If
            If falgName = True And kronos.Cells(rKronos, 6).Value = "Y" Then
                cntPayedDays = cntPayedDays + kronos.Cells(rKronos, 3).Value
            End If
        Next rKronos
    Next rReport
    
    With report.Parent
        .Save
        .Close
    End With
    
    kronos.Parent.Close

End Sub

Sub setInboundEmails(report As Worksheet)

    Dim ws As Worksheet
    Dim ws4 As Worksheet
    Dim wsEmails As Worksheet
    Dim userRange As Range
    Dim qtyRange As Range
    Dim ptsRange As Range
    Dim admireUser As String
    Dim classification As String

    Set ws = ThisWorkbook.Worksheets(1)
    Set ws4 = ThisWorkbook.Worksheets(4)
    Set wsEmails = Application.Workbooks.Open(ticketsRepliesPath).Sheets(1)
    Set userRange = getRange(wsEmails, "B")
    Set qtyRange = getRange(wsEmails, "C")
    
    wsEmails.Cells(1, 4).Value = "pts"
    
    For rowEmail = 2 To wsEmails.UsedRange.Rows.Count
        flag = False
        classification = wsEmails.Cells(rowEmail, 1).Value
        For rowLookup = 2 To ws4.UsedRange.Rows.Count
            If flag = False Then
                For colLookup = 1 To ws4.UsedRange.Columns.Count
                    If ws4.Cells(rowLookup, colLookup).Value = classification Then
                        wsEmails.Cells(rowEmail, 4).Value = ws4.Cells(1, colLookup).Value
                        flag = True
                        Exit For
                    End If
                Next colLookup
            End If
        Next rowLookup
        If flag = False Then
            wsEmails.Cells(rowEmail, 4).Value = "0.25"
        End If
    Next rowEmail
    
    Set ptsRange = getRange(wsEmails, "D")
    
    For row = 2 To ws.UsedRange.Rows.Count
        admireUser = ws.Cells(row, 3).Value
        report.Cells(row, 7).Value = Application.WorksheetFunction.SumIfs(qtyRange, userRange, admireUser, ptsRange, "0.25")
        report.Cells(row, 8).Value = Application.WorksheetFunction.SumIfs(qtyRange, userRange, admireUser, ptsRange, "0.5")
        report.Cells(row, 9).Value = Application.WorksheetFunction.SumIfs(qtyRange, userRange, admireUser, ptsRange, "0.75")
        report.Cells(row, 10).Value = Application.WorksheetFunction.SumIfs(qtyRange, userRange, admireUser, ptsRange, "1")
    Next row
    
    wsEmails.Parent.Close False

End Sub

Sub setEmailsCol(path As String, userCol As String, qtyCol As String, report As Worksheet, reportCol As String)

    Dim we As Worksheet
    Dim wsEmails As Worksheet
    Dim userRange As Range
    Dim qtyRange As Range
    Dim admireUser As String

    Set ws = ThisWorkbook.Worksheets(1)
    Set wsEmails = Application.Workbooks.Open(path).Sheets(1)
    Set userRange = getRange(wsEmails, userCol)
    Set qtyRange = getRange(wsEmails, qtyCol)
    
    For row = 2 To ws.UsedRange.Rows.Count
        admireUser = ws.Cells(row, 3).Value
        report.Cells(row, reportCol).Value = Application.WorksheetFunction.SumIfs(qtyRange, userRange, admireUser)
    Next row
    
    wsEmails.Parent.Close False

End Sub

Sub setEmails(report As Worksheet)

    'Call setEmailsCol(ticketsRepliesPath, "B", "C", report, "G")
    Call setInboundEmails(report)
    Call setEmailsCol(ticketsNewPath, "A", "B", report, "N")
    Call setEmailsCol(ticketsClosedPath, "B", "C", report, "O")

End Sub

Sub setAdmire(report As Worksheet)

    Dim we As Worksheet
    Dim wsK4K As Worksheet
    Dim wsLeads As Worksheet
    Dim wsAuction As Worksheet
    
    Dim onlineRange As Range
    Dim adjusterRange As Range
    Dim enteredByRange As Range
    Dim leadsEnteredByRange As Range
    Dim leadsStatusRange As Range
    
    Dim admireUser As String
    Dim admireFullName As String

    Set ws = ThisWorkbook.Worksheets(1)
    Set wsK4K = Application.Workbooks.Open(k4kPath).Sheets(1)
    Set wsLeads = Application.Workbooks.Open(leadsPath).Sheets(1)
    Set wsAuction = Application.Workbooks.Open(auctionPath).Sheets(1)
    
    Set onlineRange = getRange(wsK4K, "AD")
    Set adjusterRange = getRange(wsK4K, "AW")
    Set enteredByRange = getRange(wsK4K, "BR")
    Set leadsEnteredByRange = getRange(wsLeads, "AE")
    'Set leadsStatusRange = getRange(wsLeads, "E")
    Set auctionUserRange = getRange(wsAuction, "A")
    Set auctionTotalRange = getRange(wsAuction, "E")
    
    For row = 2 To ws.UsedRange.Rows.Count
        admireUser = ws.Cells(row, 3).Value
        admireFullName = ws.Cells(row, 1).Value
        report.Cells(row, 11).Value = _
            report.Cells(row, 7).Value + _
            report.Cells(row, 8).Value + _
            report.Cells(row, 9).Value + _
            report.Cells(row, 10).Value
        report.Cells(row, 12).Value = _
            report.Cells(row, 7).Value * 0.25 + _
            report.Cells(row, 8).Value * 0.5 + _
            report.Cells(row, 9).Value * 0.75 + _
            report.Cells(row, 10).Value
        report.Cells(row, 13).Value = report.Cells(row, 14).Value * 0.75
        'Copart
        report.Cells(row, 17).Value = Application.WorksheetFunction.CountIfs(adjusterRange, admireFullName)
        report.Cells(row, 18).Value = report.Cells(row, 17).Value * 0.4
        report.Cells(row, 18).NumberFormat = "#"
        'Total
        report.Cells(row, 19).Value = "=D" & row & "+F" & row & "+L" & row & "+L" & row & "+P" & row & "+R" & row
        'Donations
        report.Cells(row, 20).Value = Application.WorksheetFunction.CountIfs(enteredByRange, admireUser, onlineRange, False)
        'Leads
        report.Cells(row, 21).Value = Application.WorksheetFunction.CountIfs(leadsEnteredByRange, admireUser)
        'Auction
        report.Cells(row, 22).Value = Application.WorksheetFunction.SumIfs(auctionTotalRange, auctionUserRange, admireUser)
    Next row
    
    wsK4K.Parent.Close False
    wsLeads.Parent.Close False
    wsAuction.Parent.Close False

End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim str As Variant
    For Each str In arr
        If str = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next str
    IsInArray = False
End Function

Sub setAdmireFU(report As Worksheet)

    Dim we As Worksheet
    Dim wsFU As Worksheet
    
    Dim userRange As Range
    Dim statusRange As Range
    
    Dim admireUser As String
    Dim admireFullName As String

    Set ws = ThisWorkbook.Worksheets(1)
    Set wsFU = Application.Workbooks.Open(fuPath).Sheets(1)
    
    Set userRange = getRange(wsFU, "B")
    Set statusRange = getRange(wsFU, "L")
    
    Dim v As String
    Dim users As String
    Dim fus As String
    Dim usersArray() As String
    Dim fuValArray() As String
    
    users = ws.Cells(2, 3).Value
    For row = 3 To ws.UsedRange.Rows.Count
        admireFuUser = LCase(ws.Cells(row, 3).Value)
        users = users & "," & admireFuUser
    Next row
    usersArray = Split(users, ",")
    
    For row = 2 To wsFU.UsedRange.Rows.Count
        admireUser = LCase(wsFU.Cells(row, 2).Value)
        v = wsFU.Cells(row, 12).Value
        If v <> "Arrange Pickup" And v <> "Arrange Rush Pickup" And IsInArray(admireUser, usersArray) = True And IsInArray(v, Split(Right(fus, Len(fus)), ",")) = False Then
            fus = fus & "," & v
        End If
    Next row
    fuValArray = Split(Right(fus, Len(fus) - 1), ",")
    
    For row = 2 To ws.UsedRange.Rows.Count
        admireFuUser = ws.Cells(row, 3).Value
        report.Cells(row, 24).Value = Application.WorksheetFunction.CountIfs(userRange, admireFuUser, statusRange, "Arrange Pickup") + _
                                      Application.WorksheetFunction.CountIfs(userRange, admireFuUser, statusRange, "Arrange Rush Pickup")
        'report.Cells(row, 20).Value = Application.WorksheetFunction.CountIfs(userRange, admireFuUser, statusRange, "Scheduled Pickup")
        'report.Cells(row, 21).Value = Application.WorksheetFunction.CountIfs(userRange, admireFuUser, statusRange, "Night/Weekend")
        'report.Cells(row, 22).Value = Application.WorksheetFunction.CountIfs(userRange, admireFuUser, statusRange, "Copart To Enter")
        For fuCol = 0 To UBound(fuValArray)
            If row = 2 Then
                report.Cells(1, 25 + fuCol).Value = "-" & fuValArray(fuCol)
            End If
            report.Cells(row, 25 + fuCol).Value = Application.WorksheetFunction.CountIfs(userRange, admireFuUser, statusRange, fuValArray(fuCol))
        Next fuCol
    Next row
    
    wsFU.Parent.Close False

End Sub

Sub setAdmireStatistics(report As Worksheet)

    Dim we As Worksheet
    Dim wsStat As Worksheet
    
    Dim userRange As Range
    Dim vinsRange As Range
    Dim mileageRange As Range
    Dim pickupRange As Range
    Dim rushRange As Range
    
    Dim admireUser As String

    Set ws = ThisWorkbook.Worksheets(1)
    Set wsStat = Application.Workbooks.Open(statPath).Sheets(1)
    
    Set userRange = getRange(wsStat, "A")
    Set vinsRange = getRange(wsStat, "B")
    Set mileageRange = getRange(wsStat, "C")
    Set notesRange = getRange(wsStat, "D")
    Set pickupRange = getRange(wsStat, "E")
    Set rushRange = getRange(wsStat, "F")
    
    Dim v As String
    Dim users As String
    Dim fus As String
    Dim usersArray() As String
    Dim fuValArray() As String
    
    colStart = report.UsedRange.Columns.Count
    
    report.Cells(1, colStart + 1).Value = "Missing VINs"
    report.Cells(1, colStart + 2).Value = "Missing VINs Percentage"
    report.Cells(1, colStart + 3).Value = "Missing Mileage"
    report.Cells(1, colStart + 4).Value = "Missing Mileage Percentage"
    report.Cells(1, colStart + 5).Value = "Missing Notes"
    report.Cells(1, colStart + 6).Value = "Missing Notes Percentage"
    report.Cells(1, colStart + 7).Value = "Schedule P/U Contact Log"
    'report.Cells(1, colStart + 6).Value = "Schedule P/U Contact Log Percentage"
    report.Cells(1, colStart + 8).Value = "Rush P/U Contact Log"
    'report.Cells(1, colStart + 8).Value = "Rush P/U Contact Log Percentage"
    
    For row = 2 To ws.UsedRange.Rows.Count
        admireUser = ws.Cells(row, 3).Value
        
        On Error GoTo ErrCol:
        statRow = Application.WorksheetFunction.Match(admireUser, userRange, 0)
        report.Cells(row, colStart + 1).Value = Application.WorksheetFunction.index(vinsRange, statRow, 1)
        report.Cells(row, colStart + 2).Value = report.Cells(row, colStart + 1).Value / report.Cells(row, 20).Value
        report.Cells(row, colStart + 2).NumberFormat = "0.#%"
        report.Cells(row, colStart + 3).Value = Application.WorksheetFunction.index(mileageRange, statRow, 1)
        report.Cells(row, colStart + 4).Value = report.Cells(row, colStart + 3).Value / report.Cells(row, 20).Value
        report.Cells(row, colStart + 4).NumberFormat = "0.#%"
        report.Cells(row, colStart + 5).Value = Application.WorksheetFunction.index(notesRange, statRow, 1)
        report.Cells(row, colStart + 6).Value = report.Cells(row, colStart + 5).Value / report.Cells(row, 20).Value
        report.Cells(row, colStart + 6).NumberFormat = "0.#%"
        report.Cells(row, colStart + 7).Value = Application.WorksheetFunction.index(pickupRange, statRow, 1)
        report.Cells(row, colStart + 8).Value = Application.WorksheetFunction.index(rushRange, statRow, 1)
NextRow:
    Next row

    wsStat.Parent.Close False
    Exit Sub
    
ErrCol:
    Resume NextRow

End Sub

Sub setPickups(report As Worksheet)

    Dim wsK4K As Worksheet
    Dim towerRange As Range
    Dim sumRange As Range
    Dim currRow As Integer
    Dim statCol As Integer
    Dim puCol As Integer
    Dim tower As String
    Dim status As String
    Dim flagTower As Boolean
    Dim flagStatus As Boolean

    Set wsK4K = Application.Workbooks.Open(k4kPath).Sheets(1)
    Set towerRange = getRange(wsK4K, "H")
    Set statusRange = getRange(wsK4K, "J")
    
    report.Cells(1, 1).Value = "Tower"

    For row = 2 To wsK4K.UsedRange.Rows.Count
         flagTower = False
        flagStatus = False
        tower = wsK4K.Cells(row, 8).Value
        status = wsK4K.Cells(row, 10).Value
        
        For row2 = 2 To report.UsedRange.Rows.Count
            If report.Cells(row2, 1).Value = tower Then
                currRow = row2
                flagTower = True
                Exit For
            End If
        Next row2
        
        If flagTower = False Then
            currRow = report.UsedRange.Rows.Count + 1
            report.Cells(currRow, 1).Value = tower
        End If
        
        For col = 2 To report.UsedRange.Columns.Count
            If compare(report.Cells(1, col).Value, "Picked Up") Then
                puCol = col
            End If
            If compare(report.Cells(1, col).Value, status) Then
                statCol = col
                flagStatus = True
                Exit For
            End If
        Next col
        
        If flagStatus = False Then
            statCol = report.UsedRange.Columns.Count + 1
            report.Cells(1, statCol).Value = status
        End If
        
        report.Cells(currRow, statCol).Value = Application.WorksheetFunction.CountIfs(towerRange, tower, statusRange, status)
    Next row
    
    maxCol = report.UsedRange.Columns.Count
    totalCol = maxCol + 1
    report.Cells(1, totalCol).Value = "Total"
    report.Cells(1, totalCol + 1).Value = "Pickup Rate"
    For row = 2 To report.UsedRange.Rows.Count
        Set sumRange = report.Range("B" & row & ":" & report.Cells(row, maxCol).Address(False, False))
        report.Cells(row, totalCol).Value = Application.WorksheetFunction.Sum(sumRange)
        report.Cells(row, totalCol + 1).Value = report.Cells(row, puCol).Value / report.Cells(row, totalCol).Value
        report.Cells(row, totalCol + 1).NumberFormat = "0.##"
    Next row
    
    report.Rows(1).Font.Bold = True
    report.Range("A1:AZ1").HorizontalAlignment = xlCenter
    report.Columns("A:AZ").AutoFit
    
    wsK4K.Parent.Close False

End Sub

Sub Button1_Click()

    Dim runParams As Worksheet
    Dim wbReport As Workbook
    Dim onlineRange As Range
    Dim name As String
    Dim rangeParam As String
    Dim timeStamp As String
    Dim row As Integer
    Dim colFiles As Integer

    Set runParams = ThisWorkbook.Worksheets(1)
    
    colFiles = 6
    path = ThisWorkbook.path + "\"
    timeStamp = Format(Now(), "mmddyyyyhhmm")
    tmpReportPath = path + "TMPActivityHistoryByAgent.xls"
    templatePath = path + runParams.Cells(2, colFiles).Value
    finalReportPath = path + "FinalReport_" & timeStamp & ".xls"
    kronosPath = path + runParams.Cells(3, colFiles).Value
    ticketsRepliesPath = path + runParams.Cells(4, colFiles).Value
    ticketsNewPath = path + runParams.Cells(5, colFiles).Value
    ticketsClosedPath = path + runParams.Cells(6, colFiles).Value
    k4kPath = path + runParams.Cells(7, colFiles).Value
    leadsPath = path + runParams.Cells(8, colFiles).Value
    auctionPath = path + runParams.Cells(9, colFiles).Value
    fuPath = path + runParams.Cells(10, colFiles).Value
    statPath = path + runParams.Cells(11, colFiles).Value
    
    
    rangeParam = runParams.Cells(2, colFiles + 1).Value
        
    Call createReport
    Call setKronosHours
    
    Set wbReport = Application.Workbooks.Open(finalReportPath)
    Call setEmails(wbReport.Sheets(1))
    Call setAdmire(wbReport.Sheets(1))
    Call setAdmireFU(wbReport.Sheets(1))
    Call setAdmireStatistics(wbReport.Sheets(1))
    Call setPickups(wbReport.Sheets(3))
    
    For row = 2 To runParams.UsedRange.Rows.Count
        name = runParams.Cells(row, 1).Value
        Call updateTemplateParam(name, rangeParam)
        Call runTVRRunAndWait
        Call setTeleVantage(wbReport, row, name, False)
    Next row
    
    runParams.Cells(2, colFiles + 2).Value = "FinalReport_" & timeStamp & ".xls"
    
    With wbReport
        .Save
        .Close
    End With
    
    MsgBox "Generated Report File: " & finalReportPath

End Sub


Sub ButtonTV_Click()

    Dim runParams As Worksheet
    Dim wbReport As Workbook
    Dim onlineRange As Range
    Dim name As String
    Dim rangeParam As String
    Dim timeStamp As String
    Dim row As Integer
    Dim colFiles As Integer

    Set runParams = ThisWorkbook.Worksheets(1)
    
    colFiles = 6
    path = ThisWorkbook.path + "\"
    timeStamp = Format(Now(), "mmddyyyyhhmm")
    tmpReportPath = path + "TMPActivityHistoryByAgent.xls"
    templatePath = path + runParams.Cells(2, colFiles).Value
    finalReportPath = path + "FinalReport_" & timeStamp & ".xls"
    
    rangeParam = runParams.Cells(2, colFiles + 1).Value
        
    Call createTvReport
    Set wbReport = Application.Workbooks.Open(finalReportPath)
    
    For row = 2 To runParams.UsedRange.Rows.Count
        name = runParams.Cells(row, 1).Value
        Call updateTemplateParam(name, rangeParam)
        Call runTVRRunAndWait
        Call setTeleVantage(wbReport, row, name, True)
    Next row
    
    runParams.Cells(2, colFiles + 2).Value = "FinalReport_" & timeStamp & ".xls"
    
    With wbReport
        .Save
        .Close
    End With
    
    MsgBox "Generated Report File: " & finalReportPath

End Sub











