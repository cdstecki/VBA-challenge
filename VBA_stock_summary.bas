Attribute VB_Name = "Module1"
Sub stock_summary()

    'identify variables
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        Dim ticker As String
        Dim yearly_change As Double
        yearly_change = 0
        Dim percent_change As Double
        percent_change = 0
        Dim total_volume_stock As Double
        total_stock_volume = 0
        Dim greatest_increase As Double
        greatest_increase = 0
        Dim opening_price As Double
        opening_price = 0
        Dim closing_price As Double
        closing_price = 0
        Dim summary_table_row As Integer
        summary_table_row = 2
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'identify bonus variables
        Dim greatest_decrease As Double
        Dim greatest_total_volume As Double


    
        'Set Summary Range Headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
    
        'Loop through all stocks
        For i = 2 To last_row
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
            
                'Calculate yearly change and percent change
                open_price = ws.Cells(2, 3).Value
                close_price = ws.Cells(i, 6).Value
                yearly_change = (close_price - open_price)
                percent_change = (close_price - open_price) / open_price
            
                'Add total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
                'Print calculations to the summary table and add color formatting
                ws.Range("I" & summary_table_row).Value = ticker
                ws.Range("J" & summary_table_row).Value = yearly_change
                    If (yearly_change > 0) Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    End If
                ws.Range("K" & summary_table_row).Value = FormatPercent(percent_change)
                ws.Range("L" & summary_table_row).Value = total_stock_volume
    
                'Add one to summary table row
                summary_table_row = summary_table_row + 1
            
                'reset variables to zero
                total_stock_volume = 0
                open_price = 0
                close_price = 0
            
            'If the cell immediately following a row is the same ticker.
        
            Else
        
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
    Next ws

End Sub


