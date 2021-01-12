Code for Bonus Summary Table

Sub Bonus()
    'loop through all sheets
    For Each ws In Worksheets
    
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
    
    'Find the last row
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
        For i = 2 To LastRow
             'Bonus Greatest % Increase
            If ws.Cells(i, 11).Value > greatest_percent_increase Then
                greatest_percent_increase = ws.Cells(i, 11).Value
                greatest_percent_increase_ticker = ws.Cells(i, 8).Value
            End If
            'Bonus Greatest % Decrease
            If ws.Cells(i, 11).Value < greatest_percent_decrease Then
                greatest_percent_decrease = ws.Cells(i, 11).Value
                greatest_percent_decrease_ticker = ws.Cells(i, 8).Value
            End If
            'Bonus Greatest Total Volume
            If ws.Cells(i, 12).Value > greatest_total_volume Then
                greatest_total_volume = ws.Cells(i, 12).Value
                greatest_total_volume_ticker = ws.Cells(i, 8).Value
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
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
    
    'format percent change as a percent
        ws.Range("P2:P3").NumberFormat = "0.00%"

    Next ws

End Sub
