Attribute VB_Name = "Module1"
Option Explicit

Sub RunAuditTest()
    Dim wsData As Worksheet, wsControl As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, i As Long, sampleCount As Long
    Dim matThreshold As Double, suspiciousWords() As String
    Dim resultRow As Long

' Set references
    Set wsData = ThisWorkbook.Sheets("GL_Data")
    Set wsControl = ThisWorkbook.Sheets("ControlPanel")
    
' Get user inputs
    matThreshold = wsControl.Range("C3").Value
    sampleCount = wsControl.Range("C4").Value
    suspiciousWords = Split(wsControl.Range("C5").Value, ",")
    
' Delete existing AuditResults sheet if present
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("AuditResults").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
' Create new AuditResults sheet
    Set wsOutput = Worksheets.Add
    wsOutput.Name = "AuditResults"
    wsOutput.Range("A1:F1").Value = Array("Date", "Description", "Amount", "Vendor", "Risk Reason", "Flag")
    resultRow = 2

' Loop through GL data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
    Dim dateVal As String, desc As String, amt As Double, vendor As String
    Dim riskReason As String: riskReason = ""

    dateVal = wsData.Cells(i, 1).Value
    desc = LCase(wsData.Cells(i, 2).Value)
    amt = wsData.Cells(i, 3).Value
    vendor = wsData.Cells(i, 4).Value

' Rule 1: High amount
    If amt > matThreshold Then
    riskReason = "High Amount"
    End If
        
' Rule 2: Suspicious keywords
    Dim word As Variant
    For Each word In suspiciousWords
    If InStr(desc, Trim(LCase(word))) > 0 Then
    If riskReason <> "" Then riskReason = riskReason & ", "
    riskReason = riskReason & "Keyword: " & word
    End If
    Next word
        
' Rule:Weekend Transaction
    If Weekday(wsData.Cells(i, 1).Value, vbMonday) > 5 Then
    If riskReason <> "" Then riskReason = riskReason & ", "
    riskReason = riskReason & "Weekend Date"
    End If

' Rule: Unapproved Vendor
    Dim vendorApproved As Boolean
    vendorApproved = False
    Dim vRow As Long
    vRow = 8
    Do While wsControl.Cells(vRow, 2).Value <> ""
    If Trim(LCase(vendor)) = Trim(LCase(wsControl.Cells(vRow, 2).Value)) Then
    vendorApproved = True
    Exit Do
    End If
    vRow = vRow + 1
    Loop
    If Not vendorApproved Then
    If riskReason <> "" Then riskReason = riskReason & ", "
    riskReason = riskReason & "Unapproved Vendor"
    End If

' If any risk reason, flag it
    If riskReason <> "" Then
    wsOutput.Cells(resultRow, 1).Value = dateVal
    wsOutput.Cells(resultRow, 2).Value = wsData.Cells(i, 2).Value
    wsOutput.Cells(resultRow, 3).Value = amt
    wsOutput.Cells(resultRow, 4).Value = vendor
    wsOutput.Cells(resultRow, 5).Value = riskReason
    wsOutput.Cells(resultRow, 6).Value = "FLAGGED"
    resultRow = resultRow + 1
    End If
    Next i
' Formatting Audit Results
    Call FormatAuditResults(wsOutput, resultRow - 1)
    
' Summary stats (basic)
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = Worksheets("Dashboard")
    On Error GoTo 0
    If wsDash Is Nothing Then
    Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    On Error Resume Next
    wsDash.Name = "Dashboard"
    On Error GoTo 0
    End If
    If wsDash Is Nothing Then
    MsgBox "Failed to create Dashboard sheet.", vbCritical
    Exit Sub
    End If
    Dim totalTrans As Long, flaggedTrans As Long
    Dim riskDict As Object: Set riskDict = CreateObject("Scripting.Dictionary")
    totalTrans = lastRow - 1
    flaggedTrans = resultRow - 2
    wsDash.Range("C4").Value = totalTrans
    wsDash.Range("C5").Value = flaggedTrans
    wsDash.Range("C6").Value = Format(flaggedTrans / totalTrans, "0.0%")

' Count top risk reasons
    For i = 2 To resultRow - 1
    Dim reasons As Variant
    reasons = Split(wsOutput.Cells(i, 5).Value, ",")
    For Each word In reasons
    word = Trim(word)
    If riskDict.exists(word) Then
    riskDict(word) = riskDict(word) + 1
    Else
    riskDict.Add word, 1
    End If
    Next word
    Next i

' Show top 3 reasons
    Dim sortedKeys() As Variant
    sortedKeys = riskDict.keys
    Dim j As Long, k As Long, temp As Variant
' Simple bubble sort by frequency
    For j = 0 To riskDict.Count - 2
    For k = j + 1 To riskDict.Count - 1
    If riskDict(sortedKeys(k)) > riskDict(sortedKeys(j)) Then
    temp = sortedKeys(j)
    sortedKeys(j) = sortedKeys(k)
    sortedKeys(k) = temp
    End If
    Next k
    Next j
' Output top 3 reasons
    For i = 0 To WorksheetFunction.Min(2, riskDict.Count - 1)
    wsDash.Range("C" & (7 + i)).Value = sortedKeys(i) & " (" & riskDict(sortedKeys(i)) & ")"
    Next i



' Audit log
Dim wsLog As Worksheet
On Error Resume Next
Set wsLog = Worksheets("AuditLog")
If wsLog Is Nothing Then
Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
wsLog.Name = "AuditLog"
wsLog.Range("A1:F1").Value = Array("Run Date", "User", "Threshold", "# Flagged", "# Sampled", "Keywords")
End If
On Error GoTo 0
Dim logRow As Long
logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
wsLog.Cells(logRow, 1).Value = Now
wsLog.Cells(logRow, 2).Value = Environ("Username")
wsLog.Cells(logRow, 3).Value = matThreshold
wsLog.Cells(logRow, 4).Value = flaggedTrans
wsLog.Cells(logRow, 5).Value = sampleCount
wsLog.Cells(logRow, 6).Value = wsControl.Range("B3").Value
' Formatting Audit Log
Call FormatAuditResults(wsLog, resultRow - 1)


' Barchart of flag reasons
Dim riskChartRow As Long
riskChartRow = 13
With wsDash
' Populate data
Dim t As Long
For t = 0 To riskDict.Count - 1
.Cells(riskChartRow + 1 + t, 2).Value = sortedKeys(t)
.Cells(riskChartRow + 1 + t, 3).Value = riskDict(sortedKeys(t))
Next t
End With


MsgBox "Audit completed. See 'AuditResults' sheet.", vbInformation
End Sub

Sub GenerateSampleFromAuditResults()
    Dim wsAudit As Worksheet, wsSample As Worksheet
    Dim sampleCount As Long, lastRow As Long
    Dim i As Long, totalAmt As Double
    Dim cumProbs() As Double, randValue As Double
    Dim pickedRows As Object, selectedRow As Long
    Dim attempt As Integer
    
    Set wsAudit = Worksheets("AuditResults")
    sampleCount = Worksheets("ControlPanel").Range("C4").Value

    lastRow = wsAudit.Cells(wsAudit.Rows.Count, "A").End(xlUp).Row
    If lastRow <= 1 Then
    MsgBox "No flagged data found.", vbExclamation
    Exit Sub
    End If
    
    ' Create output sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("SampledTransactions").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsSample = Worksheets.Add
    wsSample.Name = "SampledTransactions"
    wsSample.Range("A1:F1").Value = Array("Date", "Description", "Amount", "Vendor", "Risk Reason", "Flag")
    
    ' Calculate total amount
    totalAmt = 0
    ReDim cumProbs(2 To lastRow)
    For i = 2 To lastRow
    totalAmt = totalAmt + wsAudit.Cells(i, 3).Value
    cumProbs(i) = totalAmt ' running total
    Next i

    ' Normalize cumulative probabilities
    For i = 2 To lastRow
    cumProbs(i) = cumProbs(i) / totalAmt
    Next i

    ' Store picked rows to avoid duplicates
    Set pickedRows = CreateObject("Scripting.Dictionary")
    Randomize
    Do While pickedRows.Count < WorksheetFunction.Min(sampleCount, lastRow - 1)
    randValue = Rnd
    For i = 2 To lastRow
    If randValue <= cumProbs(i) Then
    selectedRow = i
    Exit For
    End If
    Next i
    If Not pickedRows.exists(selectedRow) Then
    pickedRows.Add selectedRow, True
    wsAudit.Rows(selectedRow).Copy Destination:=wsSample.Rows(pickedRows.Count + 1)
    End If
    Loop

    ' Format output
    With wsSample
    .Columns("A:F").AutoFit
    .Range("C2:C" & pickedRows.Count + 1).NumberFormat = "#,##0.00"
    .Range("A2:A" & pickedRows.Count + 1).NumberFormat = "yyyy-mm-dd"
    .Range("F2:F" & pickedRows.Count + 1).Interior.Color = RGB(255, 199, 206)
    End With
    
    ' Formatting Sampled Results
    Call FormatAuditResults(wsSample, sampleCount - 1)

MsgBox "Monetary Unit Sampling complete. " & pickedRows.Count & " transactions selected.", vbInformation
End Sub

Sub FormatAuditResults(ws As Worksheet, lastRow As Long)
' Format header
    With ws.Range("A1:F1")
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    
' Format data rows
    With ws.Range("A2:F" & lastRow)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("C2:C" & lastRow)
    .NumberFormat = "#,##0.00"
    End With
    
    With ws.Range("A2:A" & lastRow)
        .NumberFormat = "yyyy-mm-dd"
    End With
    
' Autofit columns
    ws.Columns("A:F").AutoFit
End Sub


