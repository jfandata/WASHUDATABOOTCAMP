Sub TotalStockVolumeEASY()

'set ticker variable
Dim ticker As String

'set TotalStockVolume
Dim TotalStockVolume As Double
TotalStockVolume = 0

'store ticker and TotalStockVolume for the year in summary table
Dim currentSummaryRow As Integer
currentSummaryRow = 2

'set headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"

'find last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through the table
For i = 2 To lastRow

    'check to see we are in the same ticker block, if it is not
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'enter the ticker
        ticker = Cells(i, 1).Value
        
        'add the TotalStockVolume
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        
        'print the ticker name into summary table
        Cells(currentSummaryRow, 9).Value = ticker
        
        'print the TotalStockVolume into summary table
        Cells(currentSummaryRow, 10).Value = TotalStockVolume
        
        'add one to the summary table row
        currentSummaryRow = currentSummaryRow + 1
        
        'reset the TotalStockVolume
        TotalStockVolume = 0
        
    'if the row below is the same ticker
    Else
    
        'add the TotalStockVolume
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        
    End If
    
Next i

End Sub