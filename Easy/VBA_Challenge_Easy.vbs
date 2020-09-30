Sub TotalStockVolume():

'Loop Through All Worksheets
For Each ws In Worksheets

'Set our Initial Variables
Dim i, j As Integer
Dim total As Double
Dim ticker As String
Dim lastrow As Long

'Determine the last Row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Add Header Name New Columns
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Total Stock Volume"

'Set Initial Total
 total = 0
 j = 2

'Loop Through Each Year of Stock Data and Grab the Total Amount
 For i = 2 To lastrow
     If ws.Range("A" & i).Value = ws.Range("A" & i + 1).Value Then
         total = total + ws.Range("G" & i).Value

     Else
         ticker = ws.Range("A" & i).Value
         ws.Range("I" & j).Value = ticker
         ws.Range("J" & j).Value = total + Range("G" & i).Value
         'Add a New Row Next Ticker and Reset Total
         j = j + 1
         total = 0
     End If

 Next i
Next ws
End Sub
