Attribute VB_Name = "Module1"
Sub StockData()
    
    'Variable for counting worksheets
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    'Variable to hold ticker symbol
    Dim Ticker As String
    
    'Variable to hold opening price
    Dim StockOpen As Double
    
    'Variable to hold closing price
    Dim StockClose As Double
    
    'Variable to hold Stock Price difference
    Dim StockDif As Double
    
    'Variable to hold Stock Volume
    Dim StockVolume As Double
    
    'Variable to find last row in sheet
    Dim lastRow As Long
    
    'Variable to find largest total volume amount
    Dim Max_Volume As Double
    
    'Variable to hold greatest % increase
    Dim Greatest_Increase As Double
    
    'Variable to find greatest % decrease
    Dim Greatest_Decrease As Double
    
    'Starting row for summary data output
    Dim startRow As Integer
    
    
    
    
    
    'Set WS_Count to the number of worksheets in the active workbook
    
    
    startRow = 2
    
    StockDif = StockClose - StockOpen
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'initialize value of stock open
    StockOpen = Cells(2, 3).Value
    
    'initialize stock volume value
    StockVolume = 0
    
    'Column Headers for Summary
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    'Row Headers for Summary
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    'loop through the list
    For i = 2 To lastRow
        
        'get each ticker symbol and close price place it in new column
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
        
            'Capture Close Price
            StockClose = Cells(i, 6).Value
            'Capture Ticker
            Ticker = Cells(i, 1).Value
            
            'Display Ticker
            Cells(startRow, 9).Value = Ticker
            'Display Differential between the year
            Cells(startRow, 10).Value = StockClose - StockOpen
            
                'Define Color for the cell
                If Cells(startRow, 10).Value < 0 Then
            
                    Cells(startRow, 10).Interior.ColorIndex = 3
                Else
                    Cells(startRow, 10).Interior.ColorIndex = 4
                End If
                
            'Display difference
            Cells(startRow, 11).Value = Cells(startRow, 10).Value / StockOpen
            Cells(startRow, 11).NumberFormat = "0%"
            
            'Display Stock Volume
            Cells(startRow, 12).Value = StockVolume
            
            'increment next line to display
            startRow = startRow + 1
            
            'initialize StockOpen variable for next group
            StockOpen = Cells(i + 1, 3).Value
            'initialize StockVolume variable for next group
            StockVolume = 0
            
        Else
            StockVolume = StockVolume + Cells(i, 7).Value
            
        End If
        
    Next i
    
    'Return greatest increase info
    Greatest_Increase = Application.WorksheetFunction.Max(Range("K2:K" & lastRow).Value)
    Cells(2, 17).Value = Greatest_Increase
    Cells(2, 17).NumberFormat = "0%"
    
        For i = 2 To lastRow
            If Cells(i, 11).Value = Greatest_Increase Then
                Cells(2, 16).Value = Cells(i, 9).Value
            End If
        Next i
        
    'Return greatest decrease info
    Greatest_Decrease = Application.WorksheetFunction.Min(Range("K2:K" & lastRow).Value)
    Cells(3, 17).Value = Greatest_Decrease
    Cells(3, 17).NumberFormat = "0%"
    
        For i = 2 To lastRow
            If Cells(i, 11).Value = Greatest_Decrease Then
                Cells(3, 16).Value = Cells(i, 9).Value
            End If
        Next i
    
    'Return Greatest Volume info
    Max_Volume = Application.WorksheetFunction.Max(Range("L2:L" & lastRow).Value)
    Cells(4, 17).Value = Max_Volume
    
        For i = 2 To lastRow
            If Cells(i, 12).Value = Max_Volume Then
                Cells(4, 16).Value = Cells(i, 9).Value
            End If
        Next i

Next ws

End Sub
