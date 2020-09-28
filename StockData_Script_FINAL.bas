Attribute VB_Name = "Module1"


Sub StockData()
    
    
    'Variable to hold ticker symbol
    Dim Ticker As String
    
    'Variable to hold opening price
    Dim StockOpen As Double
    
    'Variable to hold closing price
    Dim StockClose As Double
    
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
    
    'Loop through all worksheets
    For Each ws In Worksheets
    
        
        'initialize the starting row for the summary table
        startRow = 2
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'initialize value of stock open
        StockOpen = ws.Cells(2, 3).Value
    
        'initialize stock volume value
        StockVolume = 0
    
        'Column Headers for Summary
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        'Row Headers for Summary
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
        'loop through the list
        For i = 2 To lastRow
            
            'get each ticker symbol and close price place it in new column
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            
                'Capture Close Price
                StockClose = ws.Cells(i, 6).Value
                'Capture Ticker
                Ticker = ws.Cells(i, 1).Value
                
                'Display Ticker
                ws.Cells(startRow, 9).Value = Ticker
                'Display Differential between the year
                ws.Cells(startRow, 10).Value = StockClose - StockOpen
                
                    'Define Color for the cell
                    If ws.Cells(startRow, 10).Value < 0 Then
                        ws.Cells(startRow, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(startRow, 10).Interior.ColorIndex = 4
                    End If
                    
                'Account for stock open of 0
                    If StockOpen = 0 Then
                        ws.Cells(startRow, 11).Value = 0
                    'Display difference
                    ElseIf StockOpen <> 0 Then
                        ws.Cells(startRow, 11).Value = ws.Cells(startRow, 10).Value / StockOpen
                        ws.Cells(startRow, 11).NumberFormat = "0%"
                    End If
                
                'Total StockVolume
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                
                'Display Stock Volume
                ws.Cells(startRow, 12).Value = StockVolume
                
                'increment next line to display
                startRow = startRow + 1
                
                'initialize StockOpen variable for next group
                StockOpen = ws.Cells(i + 1, 3).Value
                
                'initialize StockVolume variable for next group
                StockVolume = 0
                
            Else
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'Return greatest increase info
        Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow).Value)
        ws.Cells(2, 17).Value = Greatest_Increase
        ws.Cells(2, 17).NumberFormat = "0%"
        
            For i = 2 To lastRow
                If ws.Cells(i, 11).Value = Greatest_Increase Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                End If
            Next i
            
        'Return greatest decrease info
        Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow).Value)
        ws.Cells(3, 17).Value = Greatest_Decrease
        ws.Cells(3, 17).NumberFormat = "0%"
        
            For i = 2 To lastRow
                If ws.Cells(i, 11).Value = Greatest_Decrease Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                End If
            Next i
        
        'Return Greatest Volume info
        Max_Volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow).Value)
        ws.Cells(4, 17).Value = Max_Volume
        
            For i = 2 To lastRow
                If ws.Cells(i, 12).Value = Max_Volume Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                End If
            Next i

Next ws

End Sub
