Attribute VB_Name = "Module1"
Sub TickerColumn()

Dim i As Long
Dim j As Long
Dim h As Long
Dim k As Long
Dim l As Long
Dim lastrow As Long
Dim Counter As Long
Dim change As Double
Dim firstvalue As Double
Dim lastvalue As Double
Dim stockvolume As Double
Dim lastsummaryrow As Integer
Dim maxincrease As Double
Dim maxdecrease As Double
Dim maxvolume As Double

For Each ws In Worksheets
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Pct Change"
ws.Range("L1") = "Stock Volume"

Counter = 2
firstvalue = ws.Cells(2, 3)
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
stockvolume = 0

For i = 2 To lastrow
    stockvolume = stockvolume + ws.Cells(i, 7)
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        ws.Cells(Counter, 9) = ws.Cells(i, 1)
        lastvalue = ws.Cells(i, 6)
        change = lastvalue - firstvalue
        ws.Cells(Counter, 10) = change
        If firstvalue = 0 Then
            ws.Cells(Counter, 11) = change / 1
            Else
            ws.Cells(Counter, 11) = change / firstvalue
        End If
        stockvolume = stockvolume + ws.Cells(i, 7)
        ws.Cells(Counter, 12) = stockvolume
        firstvalue = ws.Cells(i + 1, 3)
        Counter = Counter + 1
        stockvolume = 0
    End If
Next i

lastsummaryrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

ws.Range("K2:K" & lastsummaryrow).Style = "Percent"
ws.Range("K2:K" & lastsummaryrow).Font.Bold = True

For k = 2 To lastsummaryrow
    If ws.Cells(k, 11) < 0 Then
        ws.Cells(k, 11).Interior.Color = RGB(255, 0, 0)
    ElseIf ws.Cells(k, 11) > 0 Then
        ws.Cells(k, 11).Interior.Color = RGB(0, 255, 0)
    Else: ws.Cells(k, 11).Interior.Color = RGB(255, 255, 0)
    End If
Next k

For l = 2 To lastsummaryrow
    If ws.Cells(l, 10) < 0 Then
        ws.Cells(l, 10).Interior.Color = RGB(255, 0, 0)
    ElseIf ws.Cells(l, 10) > 0 Then
        ws.Cells(l, 10).Interior.Color = RGB(0, 255, 0)
    Else: ws.Cells(l, 10).Interior.Color = RGB(255, 255, 0)
    End If
Next l

maxincrease = ws.Cells(2, 11)
maxdecrease = ws.Cells(2, 11)
maxvolume = ws.Cells(2, 12)
Dim incticker As String
Dim decticker As String
Dim volticker As String

For j = 2 To lastsummaryrow
    If maxincrease < ws.Cells(j + 1, 11) Then
        maxincrease = ws.Cells(j + 1, 11)
        incticker = ws.Cells(j + 1, 9)
    End If
    If maxdecrease > ws.Cells(j + 1, 11) Then
        maxdecrease = ws.Cells(j + 1, 11)
        decticker = ws.Cells(j + 1, 9)
    End If
    If maxvolume < ws.Cells(j + 1, 12) Then
        maxvolume = ws.Cells(j + 1, 12)
        volticker = ws.Cells(j + 1, 9)
    End If
Next j

ws.Range("o1") = "Ticker"
ws.Range("p1") = "Value"
ws.Range("n2") = "Greatest Increase %"
ws.Range("n3") = "Greatest Decrease %"
ws.Range("n4") = "Greatest Volume"
ws.Range("o2") = incticker
ws.Range("p2") = maxincrease
ws.Range("o3") = decticker
ws.Range("p3") = maxdecrease
ws.Range("o4") = volticker
ws.Range("p4") = maxvolume

ws.Range("p2:p3").Style = "Percent"

For h = 2 To 3
    If ws.Cells(h, 16) < 0 Then
        ws.Cells(h, 16).Interior.Color = RGB(255, 0, 0)
    ElseIf ws.Cells(h, 16) > 0 Then
        ws.Cells(h, 16).Interior.Color = RGB(0, 255, 0)
    End If
Next h

Next ws
End Sub

