Sub stockPerformance()

    For Each ws In Worksheets
    
        'Create variable for Year from tab name
        Dim Year As String
        'Get the worksheet year
        Year = ws.Name
        
        'Add headers and row titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Define variables for calculations
        Dim Ticker As String
        Dim openPrice, closePrice As Double
        Dim yearlyChange, precentChange As Double
        
        'Set initial variable for storing volume
        Dim volume As Double
        volume = 0
        
        'Find last row in ticker column
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Define and set value for the first row in the ticker summary table
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        'Loop through all tickers
        For i = 2 To lastRow
        
            'Check if next row is within same ticker. If not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Establish ticker
                Ticker = ws.Cells(i, 1).Value
                'add daily volume the total volume
                volume = volume + ws.Cells(i, 7).Value
                'Establish closePrice
                closePrice = ws.Cells(i, 6).Value
                'Calculate yearlyChange
                yearlyChange = closePrice - openPrice
                'Calculate percentChange
                percentChange = yearlyChange / openPrice
                
                'Print the ticker to the summary table
                ws.Range("I" & summary_table_row).Value = Ticker
                'Print the yearlyChange to the summary table
                ws.Range("J" & summary_table_row).Value = closePrice - openPrice
                'Format yearlyChange
                    If ws.Range("J" & summary_table_row).Value >= 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    End If
                'Print the percentChange to the summary table with formatting
                ws.Range("K" & summary_table_row).Value = percentChange
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                'Print the volume to the summary table
                ws.Range("L" & summary_table_row).Value = volume
                
                'Add one to the summary table row
                summary_table_row = summary_table_row + 1
                
                'Reset the volume
                volume = 0
                
            'If next cell is within same ticker and date is open Date...
            ElseIf ws.Cells(i, 2).Value = Year + "0102" Then
            
                'Establish openPrice
                openPrice = ws.Cells(i, 3).Value
                'add daily volume the total volume
                volume = volume + ws.Cells(i, 7).Value
        
                'If next cell is within same ticker but not openDate...
            Else
                'add daily volume the total volume
                volume = volume + ws.Cells(i, 7).Value
        
            End If
             
        Next i
        
        'Find Greatest % increase, decrease, and volume
        greatest_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
        greatest_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
        greatest_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        
        'Print Greatest % increase, decrease, and volume and the associated ticker
         ws.Range("Q2").Value = greatest_increase
         ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(greatest_increase, ws.Range("K:K"), 0))
         ws.Range("Q3").Value = greatest_decrease
         ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(greatest_decrease, ws.Range("K:K"), 0))
         ws.Range("Q2:Q3").NumberFormat = "0.00%"
         ws.Range("Q4").Value = greatest_volume
         ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(greatest_volume, ws.Range("L:L"), 0))
         
         'Autofit columns to show data
         ws.Columns("I:Q").AutoFit
     
     Next ws
     
End Sub
