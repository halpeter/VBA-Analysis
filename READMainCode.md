# VBA-Analysis

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
            If year_open = 0 Then
                percent_change = year_close
            Else
                percent_change = year_change / year_open
            End If
            
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
    
    Next i
    
    'Find last row of summary table
    SummaryTableLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        'Conditional Formating year change
    For i = 2 To SummaryTableLastRow
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    'add headers to each worksheet for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'autfit data
    ws.Columns("I:P").AutoFit
    'format percent change as a percent
    ws.Columns("K").NumberFormat = "0.00%"
    
    Next ws

End Sub

