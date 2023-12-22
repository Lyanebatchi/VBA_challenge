Sub MultipleYearStockData():
For Each ws In Worksheets

'Declare variables
Dim Worksheet As String
Dim ticker As String
Dim openmarketprice As Double
Dim highmarketprice As Double
Dim lowmarketprice As Double
Dim closemarketprice As Double
Dim volume As Integer
Dim Lastrow As Double
Dim firstrow As Double
Dim i As Double
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Double




'Create Columns name
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "YearlyChange"
ws.Cells(1, 11).Value = "PercentChange"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greastest Total Volume"

'Loop Start
' Define variables
firstrow = 2
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
closemarketprice = 0
openmarketprice = ws.Cells(2, 3)
volume = 0

'Set the Loop to pull the tickers
For i = 2 To Lastrow
    Total = Total + ws.Cells(i, 7)
    If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    ws.Cells(firstrow, 12) = Total
    ws.Cells(firstrow, 9) = ws.Cells(i, 1)
    closemarketprice = ws.Cells(i, 6)
    ws.Cells(firstrow, 10) = closemarketprice - openmarketprice
    If ws.Cells(firstrow, 10) > 0 Then
        ws.Cells(firstrow, 10).Interior.ColorIndex = 4
        
    Else
        ws.Cells(firstrow, 10).Interior.ColorIndex = 3
        
        
    End If
    
    If openmarketprice <> 0 Then
    ws.Cells(firstrow, 11) = FormatPercent((closemarketprice - openmarketprice) / openmarketprice, 2)
    Else
        ws.Cells(firstrow, 11) = 0
    End If
    
    firstrow = firstrow + 1
    openmarketprice = ws.Cells(i + 1, 3)
    Total = 0

'Loop to find values

If ws.Cells(firstrow, 11) > greatestincrease Then
    greatestincrease = ws.Cells(firstrow, 11)
    greatestincreaseticker = ws.Cells(firstrow, 9)
    
End If

If ws.Cells(firstrow, 11) < greatestdecrease Then
    greatestdecrease = ws.Cells(firstrow, 11)
    greatestdecreaseticker = ws.Cells(firstrow, 9)
    
End If
    

If ws.Cells(firstrow, 12) > greatestvolume Then
    greatestvolume = ws.Cells(firstrow, 12)
    greatestvolumeticker = ws.Cells(firstrow, 9)
    
End If


End If
Next i

'Set up location for values
ws.Cells(2, 16) = greatestincreaseticker
ws.Cells(2, 17) = FormatPercent(greatestincrease, 2)

ws.Cells(3, 16) = greatestdecreaseticker
ws.Cells(3, 17) = FormatPercent(greatestdecrease, 2)

ws.Cells(4, 16) = greatestvolumeticker
ws.Cells(4, 17) = greatestvolume
Next ws


MsgBox ("Done!")
End Sub

