Attribute VB_Name = "Module1"
Sub StockAnalysis():

Dim ticker As String
Dim Openingprice As Double
Dim Closingprice As Double
Dim Total As Double
Dim Tablerow As Integer

For Each ws In Worksheets

Total = 0
Openingprice = ws.Cells(2, "C").Value
Tablerow = 2

ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    Total = Total + ws.Cells(i, "G")

If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
    ws.Cells(Tablerow, "L") = Total
    ws.Cells(Tablerow, "I") = ws.Cells(i, "A")
    Closingprice = ws.Cells(i, "F")
    ws.Cells(Tablerow, "J") = Closingprice - Openingprice

        If ws.Cells(Tablerow, "J") > 0 Then
             ws.Cells(Tablerow, "J").Interior.Color = RGB(0, 255, 0)
        Else
            ws.Cells(Tablerow, "J").Interior.Color = RGB(255, 0, 0)
        End If

        If Openingprice <> 0 Then
            ws.Cells(Tablerow, "K") = FormatPercent((Closingprice - Openingprice) / Openingprice, 2)
        Else
            ws.Cells(Tablerow, "K") = 0
        End If

Tablerow = Tablerow + 1
Openingprice = ws.Cells(i + 1, "C")
Total = 0
End If

Next i

ws.Cells(1, "P") = "Ticker"
ws.Cells(1, "Q") = "Value"
ws.Cells(2, "O") = "Greatest % Increase"
ws.Cells(3, "O") = "Greatest % Decrease"
ws.Cells(4, "O") = "Greatest Total Volume"

Dim maxvalue As Double
Dim minvalue As Double
Dim summarytotal As Double
Dim maxticker As String
Dim minticker As String
Dim totalticker As String
Dim summarytablerow As Integer

maxvalue = ws.Cells(2, "K")
minvalue = ws.Cells(2, "K")
summarytotal = ws.Cells(2, "L")
maxticker = ws.Cells(2, "I")
minticker = ws.Cells(2, "I")
totalticker = ws.Cells(2, "I")
summarytablerow = 2

For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row

    If ws.Cells(i, "K") > maxvalue Then
        maxvalue = ws.Cells(summarytablerow, "K")
        maxticker = ws.Cells(summarytablerow, "I")

    End If

    If ws.Cells(i, "K") < minvalue Then
        minvalue = ws.Cells(summarytablerow, "K")
        minticker = ws.Cells(summarytablerow, "I")

    End If

    If ws.Cells(i, "L") > summarytotal Then
        summarytotal = ws.Cells(summarytablerow, "L")
        totalticker = ws.Cells(summarytablerow, "I")

    End If

summarytablerow = summarytablerow + 1

Next i

ws.Cells(2, "P") = maxticker
ws.Cells(2, "Q") = FormatPercent(maxvalue, 2)
ws.Cells(3, "P") = minticker
ws.Cells(3, "Q") = FormatPercent(minvalue, 2)
ws.Cells(4, "P") = totalticker
ws.Cells(4, "Q") = summarytotal

Next ws

End Sub
