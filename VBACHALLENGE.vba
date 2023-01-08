Sub Challenge2()

For Each ws In Worksheets


Dim Ticker, J As Integer
Dim YearlyChange As Double
Dim percentChange As Double
Dim totalStock As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim GreatIncr, GreatDecr, GreatStock As Double




ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "yearlychange"
ws.Range("k1").Value = "Percentchange"
ws.Range("l1").Value = "TotalStock"

'Loop through the rows

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
J = 0
For i = 2 To LastRow

If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
OpeningPrice = ws.Cells(i, 3).Value
totalStock = 0
End If
totalStock = ws.Cells(i, 7).Value + totalStock

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
ClosingPrice = ws.Cells(i, 6).Value
YearlyChange = ClosingPrice - OpeningPrice
percentChange = YearlyChange / OpeningPrice

ws.Range("I" & 2 + J).Value = ws.Cells(i, 1).Value
ws.Range("J" & 2 + J).Value = YearlyChange
ws.Range("k" & 2 + J).Value = percentChange
ws.Range("L" & 2 + J).Value = totalStock

'color
If ws.Range("J" & 2 + J).Value < 0 Then
ws.Range("J" & 2 + J).Interior.ColorIndex = 3
Else
ws.Range("J" & 2 + J).Interior.ColorIndex = 4
End If

'percent
If ws.Range("K" & 2 + J).Value <> 0 Then
ws.Range("K" & 2 + J).Value = Format(percentChange, "Percent")
Else
ws.Range("K" & 2 + J).Value = Format(0, "Percent")
End If

J = J + 1

End If

Next i

'summary

 ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
 ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

GreatDecr = ws.Range("K" & 2 + J).Value
GreatIncr = ws.Range("K" & 2 + J).Value
GreatStock = ws.Cells(2, 12).Value

J = 0
For i = 2 To LastRowI

'Increase

If ws.Range("K" & 2 + J).Value > GreatIncr Then
            GreatIncr = ws.Range("K" & 2 + J).Value
            ws.Range("P2").Value = ws.Range("I" & 2 + J).Value
        Else
            GreatIncr = GreatIncr
        End If
       
'Decrease

If ws.Range("K" & 2 + J).Value < GreatDecr Then
            GreatDecr = ws.Range("K" & 2 + J).Value
            ws.Range("P3").Value = ws.Range("I" & 2 + J).Value
        Else
            GreatDecr = GreatDecr
        End If
       
'Greatest stock volume
     
        If ws.Range("J" & 2 + J).Value > GreatStock Then
            GreatStock = ws.Range("J" & 2 + J).Value
            ws.Range("Q3").Value = ws.Range("I" & 2 + J).Value
        Else
            GreatStock = GreatStock
        End If
J = J + 1
         
Next i

 ws.Range("Q2").Value = Format(GreatIncr, "Percent")
 ws.Range("Q3").Value = Format(GreatDecr, "Percent")
  ws.Range("Q4").Value = Format(GreatStock, "Scientific")

Next ws

End Sub
