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
Range("J1").Value = "Close P"
Range("K1").Value = "Open P"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Total Stock Volume"

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
        Cells(currentSummaryRow, 13).Value = TotalStockVolume
        
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

Sub openPrices_MODERATE()

'summary table
Dim currentSummaryRow As Integer
currentSummaryRow = 2

'set variables for stockOpen and stockClose dates
Dim yearOpen As Double
Dim yearClose As Double
Dim openP As Double
Dim closeP As Double

yearOpen = 20160101
yearClose = 20161230

'find last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'loop though table
For i = 2 To lastRow

    'check if the next row is the same ticker and matches stockOpen date of 20160101
    If Cells(i + 1, 1).Value = Cells(i, 1).Value And Cells(i, 2).Value = yearOpen Then
    
    'pull the price on stockOpen date
    openP = Cells(i, 3).Value
    
    'print the price into the summary table
    Cells(currentSummaryRow, 11).Value = openP
    
    'add one to the summary table row
    currentSummaryRow = currentSummaryRow + 1
    
    End If
    
Next i
        
End Sub

Sub closePrices_MODERATE()

'summary table
Dim currentSummaryRow As Integer
currentSummaryRow = 2

'set variables for stockOpen and stockClose dates
Dim yearOpen As Double
Dim yearClose As Double
Dim openP As Double
Dim closeP As Double

yearOpen = 20160101
yearClose = 20161230

'find last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'loop though table
For i = 2 To lastRow

    'check if the next row is the same ticker and matches stockOpen date of 20161230
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 2).Value = yearClose Then
    
    'pull the close price
    closeP = Cells(i, 6).Value
    
    'print the price into the summary table
    Cells(currentSummaryRow, 10).Value = closeP
    
    'add one to the summary table row
    currentSummaryRow = currentSummaryRow + 1
    
    End If
    
Next i
        
End Sub

Sub priceChange_MODERATE()

Dim openP As Double
Dim closeP As Double
Dim change As Double

'summary table
Dim currentSummaryRow As Integer
currentSummaryRow = 2

'find last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 11).End(xlUp).Row

'loop through summary table
For i = 2 To lastRow
    
    'take difference to get change in price
    Cells(i, 12).Value = Cells(i, 10).Value - Cells(i, 11).Value
     
Next i

End Sub

Sub cellFormat_MODERATE()

'find last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 12).End(xlUp).Row

'loop through summary table
For i = 2 To lastRow

    'if positive green
    If Cells(i, 12) > 0 Then
    Cells(i, 12).Interior.ColorIndex = 4
    'if negative red
    Else
    Cells(i, 12).Interior.ColorIndex = 3
    End If
Next i
End Sub
