Attribute VB_Name = "Module1"
Sub TickerSummary()
'   Loop through each worksheet
For Each ws In Worksheets
    
    '   Find Last Row value
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '   Assign variables
    Dim volume As LongLong
    Dim ticker_count As Integer
    ticker_count = 0
    Dim open_value() As Double
    Dim close_value() As Double
    Dim percent_change As Double
    Dim change As Double
    Dim summary_table_row As Integer
    Dim tickername As String
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As LongLong
    Dim giticker As String
    Dim gdticker As String
    Dim gvticker As String
    summary_table_row = 2
    volume = 0
    change = 0
    greatest_volume = 0
    greatest_decrease = 0
    greatest_increase = 0
    
    '   Place Headers for primary and secondary summary tables
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    
    For i = 2 To LastRow
        '   When ticker i is the same as the next one down...
        If ws.Cells(i, 1).Value = ws.Cells((i + 1), 1) Then
            volume = volume + ws.Cells(i, 7).Value  'add volume to rolling total
            ReDim Preserve open_value(0 To ticker_count) As Double  'resize array parameters with each new row, preserving previous values stored
            ReDim Preserve close_value(0 To ticker_count) As Double  'resize array parameters with each new row, preserving previous values stored
            open_value(0 + ticker_count) = ws.Cells(i, 3).Value 'store opening value of each row
            close_value(0 + ticker_count) = ws.Cells(i, 6).Value    'store closing value of each row
            ticker_count = ticker_count + 1 'track the number of same-named tickers
        
        '   when ticker i is different than the next one down...
        Else
            volume = volume + ws.Cells(i, 7).Value  'add volume to rolling total
            ReDim Preserve open_value(0 To ticker_count) As Double  'resize array parameters with each new row, preserving previous values stored
            ReDim Preserve close_value(0 To ticker_count) As Double  'resize array parameters with each new row, preserving previous values stored
            open_value(0 + ticker_count) = ws.Cells(i, 3).Value 'store opening value of each row
            close_value(0 + ticker_count) = ws.Cells(i, 6).Value    'store closing value of each row
            ticker_count = ticker_count + 1 'track the number of same-named tickers
            tickername = ws.Cells(i, 1).Value   'grab the name of the current ticker
            ws.Range("I" & summary_table_row).Value = tickername    'place ticker name in summary table
            change = close_value(ticker_count - 1) - open_value(0) 'find change from end of current ticker to opening value of same ticker
            ws.Range("J" & summary_table_row).Value = change    'place change value in summary table
            ws.Range("J" & summary_table_row).NumberFormat = "#,##00.00"    'format change value to two decimal places
                If change < 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3   'format negative change values as red
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4   'format positive change values as green
                End If
            percent_change = change / ws.Cells((i - (ticker_count - 1)), 3).Value   'find percent change from end of current ticker to opening of same ticker
            ws.Range("K" & summary_table_row).Value = FormatPercent(percent_change, 2)  'place and format percent change as Percentage with two decimal places
            ws.Range("L" & summary_table_row).Value = volume    'place volume in summary table
            
            'reset volume, ticker count, and change value as we finish current ticker name and move to the next
            volume = 0
            ticker_count = 0
            change = 0
            
            'set next ticker summary data to be placed in next row of summary table
            summary_table_row = summary_table_row + 1
        End If
    Next i
    
    'find greatest increase, greatest decrease, and greatest volume from summary table
    For i = 2 To LastRow
        If ws.Cells(i, 11).Value > greatest_increase Then
            greatest_increase = ws.Cells(i, 11).Value
            giticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < greatest_decrease Then
            greatest_decrease = ws.Cells(i, 11).Value
            gdticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12).Value > greatest_volume Then
            greatest_volume = ws.Cells(i, 12).Value
            gvticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    '   place tickers for greatest increase, decrease, and volume in secondary summary table
    ws.Range("P2").Value = giticker
    ws.Range("P3").Value = gdticker
    ws.Range("P4").Value = gvticker
    
    '   place and format values for greatest increase, decrease, and volume in secondary summary table
    ws.Range("Q2").Value = FormatPercent(greatest_increase, 2)
    ws.Range("Q3").Value = FormatPercent(greatest_decrease, 2)
    ws.Range("Q4").Value = greatest_volume
    ws.Range("Q4").NumberFormat = "##0.00E+0"
    
    'format the columns to AutoFit
    ws.Columns("I:Q").AutoFit
Next ws
End Sub
