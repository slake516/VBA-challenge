Attribute VB_Name = "Module1"
Sub VBAStock()

For Each ws In Worksheets

Dim WorksheetName As String
Dim Ticker As String
Dim YrlyChg As Double
Dim totalStock As Double
Dim tickerTableRow As Long
Dim percentchg As Double

'Set initial values
totalStock = 0
percentchg = 0
tickerTableRow = 2

'Find last row
LastRow = ws.Cells(Rows, Count, 1).End(xlUp).Row

'Print column header rows
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'Adjust Column Width
    ws.Range("I1:L1").Columns.AutoFit

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Set ticker name
        Ticker = ws.Cells(i, 1).Value
        'add to totalstock volume
        totalStock = totalStock + ws.Cells(i, 7).Value
        'Print Ticker Name to Column I
        ws.Range("I" & tickerTableRow).Value = Ticker
        'Print Total Stock Volume to Column L
        ws.Range("L" & tickerTableRow).Value = totalStock
    
        tickerTableRow = tickerTableRow + 1
        totalStock = 0
        
    Else
        'Adding to Stock Volume Total
        totalStock = totalStock + ws.Cells(i, 7).Value
        
        'Yearly Change
        YrlyChg = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
        'Print yearly change
        ws.Range("J" & tickerTableRow).Value = YrlyChg
       'End If
        
        'Format Yearly Change Column
        If ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.Color = vbGreen
        Else
           ws.Cells(i, 10).Interior.Color = vbRed
       End If
        
Next i
    
   'End If
   
Next ws
    
End Sub
