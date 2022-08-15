Sub ChallengeTest():

For Each ws In Worksheets

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim TableRow As Integer
Dim YearStart As Double
Dim YearEnd As Double
Dim Stock As Double

Dim GreatestInc As Double
Dim GreatestDec As Double
Dim GreatestTotal As Double

Stock = 0
TableRow = 2
GreatestInc = 0
GreatestDec = 0
GreatestTotal = 0

For I = 2 To LastRow
    If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1).Value Then
        YearStart = ws.Cells(I, 3).Value
    End If
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        Ticker = ws.Cells(I, 1).Value
        YearEnd = ws.Cells(I, 6).Value
        Stock = Stock + ws.Cells(I, 7).Value
        ws.Range("I" & TableRow).Value = Ticker
        ws.Range("J" & TableRow).Value = YearEnd - YearStart
        ws.Range("K" & TableRow).Value = ((YearEnd - YearStart) / YearStart)
        ws.Range("L" & TableRow).Value = Stock
        If ws.Range("J" & TableRow).Value > 0 Then
            ws.Range("J" & TableRow).Interior.ColorIndex = 4
        Else
            ws.Range("J" & TableRow).Interior.ColorIndex = 3
        End If
        If ws.Range("K" & TableRow).Value > GreatestInc Then
            GreatestInc = ws.Range("K" & TableRow).Value
            ws.Range("P2").Value = ws.Range("I" & TableRow).Value
            ws.Range("Q2").Value = GreatestInc
        End If
        If ws.Range("K" & TableRow).Value < GreatestDec Then
            GreatestDec = ws.Range("K" & TableRow).Value
            ws.Range("P3").Value = ws.Range("I" & TableRow).Value
            ws.Range("Q3").Value = GreatestDec
        End If
        If ws.Range("L" & TableRow).Value > GreatestTotal Then
            GreatestTotal = ws.Range("L" & TableRow).Value
            ws.Range("P4").Value = ws.Range("I" & TableRow).Value
            ws.Range("Q4").Value = GreatestTotal
        End If
        TableRow = TableRow + 1
        Stock = 0
    Else
        Stock = Stock + ws.Cells(I, 7).Value
    
    End If
Next I

ws.Range("K2").EntireColumn.NumberFormat = "0.00%"
ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

End Sub

