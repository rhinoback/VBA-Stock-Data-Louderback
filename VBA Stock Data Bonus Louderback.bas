Attribute VB_Name = "Module1"
Sub StockPracticeBonus():

ÔBonus Code

Dim TotVol As Double
Dim RSum As Integer
Dim R As Long
Dim YearlyChange As Double
Dim PercentChange As Double
Dim GrtPctInc As Double
Dim GrtPctDcr As Double
Dim GrtTotVol As Long


Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"
Cells(1, 14) = "Greatest % Increase"
Cells(1, 15) = "Greatest % Decrease"
Cells(1, 16) = "Greatest Total Volume"

RSum = 2

TotVol = Cells(2, 7).Value
OpenPrice = Cells(2, 3).Value
GrtPctInc = Cells(2, 14).Value
GrtPctDcr = Cells(2, 15).Value
GrtTotVol = Cells(2, 16).Value

    For R = 2 To Cells(Rows.Count, "A").End(xlUp).Row + 1
    If Cells(R, 1).Value = Cells(R + 1, 1).Value Then
    TotVol = TotVol + Cells(R + 1, 7).Value
    OpenPrice = Cells(2, 3).Value
    ClosePrice = Cells(R, 6).Value
    YearlyChange = ClosePrice - OpenPrice
    PercentChange = YearlyChange / OpenPrice
    Cells(R, 11).NumberFormat = "0.00%"
    Cells(2, 14).Value = Application.WorksheetFunction.Max(Range("K:K"))
    Cells(R, 14).NumberFormat = "0.00%"
    Cells(2, 15).Value = Application.WorksheetFunction.Min(Range("K:K"))
    Cells(R, 15).NumberFormat = "0.00%"
    Cells(2, 16).Value = Application.WorksheetFunction.Max(Range("L:L"))
    
    Else
    Cells(RSum, 9).Value = Cells(R, 1).Value
    Cells(RSum, 10).Value = YearlyChange
    Cells(RSum, 11).Value = PercentChange
    Cells(RSum, 12).Value = TotVol
    PercentChange = YearlyChange / OpenPrice * 100
    TotVol = Cells(R + 1, 7).Value
    OpenPrice = Cells(R + 1, 3).Value
    ClosePrice = Cells(R, 6).Value
    RSum = RSum + 1
        
    End If

Next R

For R = 2 To Cells(Rows.Count, "A").End(xlUp).Row + 1
If Cells(R, 10) < 0 Then
Cells(R, 10).Interior.Color = RGB(255, 0, 0)

ElseIf Cells(R, 10) > 0 Then
Cells(R, 10).Interior.Color = RGB(0, 255, 0)

Else
Cells(R, 10).Interior.Color = RGB(255, 255, 255)

End If

Next R

End Sub
