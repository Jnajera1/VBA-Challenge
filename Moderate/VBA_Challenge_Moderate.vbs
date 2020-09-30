Sub TotalStockVolume():

'Loop Through All Worksheets
For Each ws In Worksheets

'Set our Initial Variables
Dim i, j As Integer
Dim total As Double
Dim ticker As String
Dim lastrow As Long
'Set Variables for Moderate Segment
Dim ychng As Double
Dim pchng As Double
Dim oprice As Double
Dim cprice As Double
Dim oprice_row As Double

'Determine the last Row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Add Header Name New Columns
ws.Range("I1").Value = "Ticker"
'Move Total Stock Volume to L1 and Yearly Change to J1
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Set Initial Total
total = 0
j = 2
'Change for Moderate Segment
oprice_row = 2
 
'Loop Through Each Year of Stock Data and Grab the Total Amount
For i = 2 To lastrow
    If ws.Range("A" & i).Value = ws.Range("A" & i + 1).Value Then
    total = total + ws.Range("G" & i).Value
    
    Else
        ticker = ws.Range("A" & i).Value
         
    'Calculate Yearly Change and Percent Change
        oprice = ws.Range("C" & oprice_row)
        cprice = ws.Range("F" & i)
        ychng = cprice - oprice

    'Calculate Percent Change
        If oprice = 0 Then
            pchng = 0
        Else
            pchng = ychng / oprice
        End If

    'Insert Grabbed Ticker,Total Volume,Yearly Change and Percent Change into Display Cells
        ws.Range("I" & j).Value = ticker
        ws.Range("L" & j).Value = totalv + ws.Range("G" & i).Value
        ws.Range("J" & j).Value = ychng
        ws.Range("K" & j).Value = pchng
        ws.Range("K" & j).NumberFormat = "0.00%"
         
    'Conditional Formating Yearly Change, Positive Green/ Negative Red
        If ws.Range("J" & j).Value > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4
        Else
            ws.Range("J" & j).Interior.ColorIndex = 3
        End If

    'Add a New Row itno Display Cells for Next Ticker, Set New open rice row and Reset Total
        j = j + 1
        totalv = 0
        oprice_row = i + 1
         
    End If
 Next i
 
 Next ws

End Sub


