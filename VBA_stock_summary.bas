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
        Dim total_stock_volume As Double
        total_stock_volume = 0
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim summary_table_row As Long
        summary_table_row = 2
        Dim last_row As Long
        Dim i As Long
        
        'Capture last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'identify bonus variables
        Dim greatest_increase_ticker As String
        Dim greatest_increase As Double
        greatest_increase = 0
        Dim greatest_decrease_ticker As String
        Dim greatest_decrease As Double
        greatest_decrease = 0
        Dim greatest_total_volume_ticker As String
        Dim greatest_total_volume As Double
        greatest_total_volume = 0

        'Set Summary Range and Headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
             
        'Set initial open price
        open_price = ws.Cells(2, 3).Value
    
        'Loop through all stocks
        For i = 2 To last_row
        
            'Reset percentage change for each loop run
            percent_change = 0
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                          
                'Calculate Yearly Change and Percent Change
                close_price = ws.Cells(i, 6).Value
                yearly_change = (close_price - open_price)
                'percent_change = (yearly_change / Open_Price)
                If open_price <> 0 Then
                    percent_change = (yearly_change / open_price)
                Else
                    percent_change = 0
                End If
            
                'Add Total Stock Volume
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
            
                'reset variables to zero and capture next open price
                yearly_change = 0
                close_price = 0
                open_price = ws.Cells(i + 1, 3).Value
                
                'identify greatest increase, greatest decrease, and greatest total volume
                If (percent_change > greatest_increase) Then
                
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker
                    
                ElseIf (percent_change < greatest_decrease) Then
                
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker
                    
                End If
                       
                If (total_stock_volume > greatest_total_volume) Then
                    greatest_total_volume = total_stock_volume
                    greatest_total_volume_ticker = ticker
                    
                End If
                
                total_stock_volume = 0
                percent_change = 0
                
            
            'If the cell immediately following a row is the same ticker.
            Else
        
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
            ws.Range("P2").Value = greatest_increase_ticker
            ws.Range("P3").Value = greatest_decrease_ticker
            ws.Range("P4").Value = greatest_total_volume_ticker
            ws.Range("Q2").Value = FormatPercent(greatest_increase)
            ws.Range("Q3").Value = FormatPercent(greatest_decrease)
            ws.Range("Q4").Value = greatest_total_volume
            
    Next ws

End Sub
