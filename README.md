Option Explicit
Sub Stocks()
'to loop thru each worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
  ws.Activate

'name each cell value in all sheets
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 15).Value = "Analysis"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'variables
Dim i As Long
Dim volume As Double
Dim tickercountrow As String
Dim lastrow As Double
Dim newticker As String
Dim openingprice As Integer
Dim closingprice As Integer
Dim yc As Double 'yc = yearly change
Dim pc As Single

volume = 0
yc = 0

tickercountrow = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'opening price of each stock and closing price
openingprice = Cells(tickercountrow, 3).Value

'to loop through the worksheets
For i = 2 To lastrow
'condition - if ticker is not equal to next ticker, add the ticker name in new ticker column I
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        newticker = Cells(i, 1).Value
        Range("I" & tickercountrow).Value = newticker
        tickercountrow = tickercountrow + 1
'calculate volume
        volume = volume + Cells(i, 7).Value
        Range("L" & tickercountrow).Value = volume             'sum of all the volume with same ticker name
        volume = 0                                               'reset the volume to 0
'calculate yealy change (closing price minus opening price)
        yc = Cells(i, 6).Value - openingprice
        Range("J" & tickercountrow).Value = yc
        yc = 0                                                   'reset the yearly change
'calculate percent change
        pc = (Cells(i, 6).Value - openingprice) / openingprice
        Range("K" & tickercountrow).Value = pc
        Range("K" & tickercountrow).NumberFormat = "0.00%"
        openingprice = Cells(i, 3).Value                 'reset the opening price
    Else
        volume = volume + Cells(i, 7).Value                   'add stock volume to total
    End If
Next i

'format color and percentage
Dim columnI As Integer
Dim j As Integer

columnI = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To columnI
    If Cells(j, 10) >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    Else
        Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j

'to find greatest increase,decrease, and volume
Dim Increase As Integer
Dim Decrease As Integer
Dim Greatest_Volume As Double
Dim percentrow As Long
Dim volumerow As Long

Increase = 0
Decrease = 0
Greatest_Volume = 0

percentrow = Cells(Rows.Count, 11).End(xlUp).Row
volumerow = Cells(Rows.Count, 12).End(xlUp).Row

'to find values for greatest increase and greatest decrease in percentage
For j = 2 To columnI
    If Cells(j, 11).Value > Increase Then
        Cells(j, 11).Value = Increase
        Range("Q2") = Increase
        Range("P2") = Cells(j, 9).Value
    ElseIf Cells(j, 11).Value < Decrease Then
        Decrease = Cells(j, 11).Value
        Range("Q3") = Decrease
        Range("P3") = Cells(j, 9).Value
    End If
Next j

For j = 2 To columnI
    If Cells(j, 12).Value > Greatest_Volume Then
        Greatest_Volume = Cells(j, 12).Value
        Range("Q4") = Greatest_Volume
        Range("P4") = Cells(j, 9).Value
    End If
Next j

Range("Q2:Q3").NumberFormat = "0.00%"            'format the cells for percentage
Range("K2").NumberFormat = "0.00%"
Next ws

End Sub


