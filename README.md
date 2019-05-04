Sub forex_exchange()
Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim rowS As Double
Dim closePrice As Double
Dim openPrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
For Each ws In Worksheets
vol = 0
Count = 0
countp = 0
rowS = 2
ws.Range("K1").Value = "Tickers"
ws.Range("N1").Value = "Volume Total"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percent Change"
LastRow = ws.Range("A1").End(xlDown).Row
For i = 2 To LastRow
If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
vol = vol + ws.Cells(i, 7).Value
Count = Count + 1
countp = countp + 1
Else
Count = Count + 1
vol = vol + ws.Cells(i, 7).Value
ticker = ws.Cells(i, 1).Value
openPrice = ws.Cells(1 + Count - countp, 3).Value
closePrice = ws.Cells(Count + 1, 6).Value
YearlyChange = closePrice - openPrice
PercentChange = ((YearlyChange / openPrice) * 100)
ws.Range("K" & rowS).Value = ticker
ws.Range("N" & rowS).Value = vol
ws.Range("L" & rowS).Value = YearlyChange
ws.Range("M" & rowS).Value = PercentChange
countp = 0
vol = 0
rowS = rowS + 1
End If
Next i
Next ws


End Sub


