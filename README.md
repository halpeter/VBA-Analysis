# VBA-challenge

Here is my VBA code with comments for this assingment:


Sub StockMarket()
    'loop through all sheets
    For Each ws In Worksheets
    
    'Find the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define variables
    Dim ticker As String
    Dim year_open As Double
    year_open = 0
    Dim year_close As Double
    year_close = 0
    Dim year_change As Double
    year_change = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim total_volume As Single
    total_volume = 0
    Dim ticker_rows As Double
    ticker_rows = 0
    Dim first_row As Double
    first_row = 0
    Dim current_row As Double
    current_row = 0
    Dim percent_change As Double
    percent_change = 0
    'Define variables for bonus
    Dim greatest_percent_increase As Double
    greatest_percent_increase = 0
    Dim greatest_percent_decrease As Double
    greatest_percent_decrease = 0
    Dim greatest_total_volume As Double
    greatest_total_volume = 0
    Dim greatest_percent_increase_ticker As String
    Dim greatest_percent_decrease_ticker As String
    Dim greatest_total_volume_ticker As String
    
      'loop through rows to find total volume
    For i = 2 To LastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'set ticker to ticker value
            ticker = ws.Cells(i, 1).Value
            'Add last value to total volume
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'add last row to ticker rows
            ticker_rows = ticker_rows + 1
            'save current row
            current_row = i
            'calculate year open row number
            first_row = current_row - ticker_rows + 1
            'saves the value of the year open
            year_open = ws.Cells(first_row, 3).Value
            'saves the value of the year close
            year_close = ws.Cells(i, 6).Value
            'calculate the difference between year open and year close
            year_change = year_open - year_close
            'calculate the percent change
            percent_change = year_change / year_open
            
            'print values in table
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("J" & Summary_Table_Row).Value = year_change
            ws.Range("K" & Summary_Table_Row).Value = percent_change
            ws.Range("L" & Summary_Table_Row).Value = total_volume
            'add 1 to summary table row to move next ticker down one
            Summary_Table_Row = Summary_Table_Row + 1
            'reset total volume and ticker rows count
            total_volume = 0
            ticker_rows = 0
        Else
            'add to the total volume
            total_volume = total_volume + ws.Cells(i, 7).Value
            'count the number of ticker rows in each ticker
            ticker_rows = ticker_rows + 1
        End If
        
        'Conditional Formating year change
        If Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
            
        'Bonus Greatest % Increase
        If ws.Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = ws.Cells(i, 11).Value
            greatest_percent_increase_ticker = ws.Cells(i, 1).Value
        End If
        'Bonus Greatest % Decrease
        If ws.Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = ws.Cells(i, 11).Value
            greatest_percent_decrease_ticker = ws.Cells(i, 1).Value
        End If
        'Bonus Greatest Total Volume
        If ws.Cells(i, 12).Value > greatest_total_volume Then
            greatest_total_volume = ws.Cells(i, 12).Value
            greatest_total_volume_ticker = ws.Cells(i, 1).Value
        End If
    
    Next i
    
    'Input bonus values into table
    ws.Cells(2, 15).Value = greatest_percent_increase_ticker
    ws.Cells(2, 16).Value = greatest_percent_increase
    ws.Cells(3, 15).Value = greatest_percent_decrease_ticker
    ws.Cells(3, 16).Value = greatest_percent_decrease
    ws.Cells(4, 15).Value = greatest_total_volume_ticker
    ws.Cells(4, 16).Value = greatest_total_volume
    
    'add headers to each worksheet for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    'autfit data
    ws.Columns("I:P").AutoFit
    'format percent change as a percent
    For i = 2 To LastRow
        ws.Cells(i, 11).NumberFormat = "0.00%"
    Next i
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
    Next ws

End Sub

