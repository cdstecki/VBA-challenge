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
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        Dim last_row As Long
        Dim i As Long
        
        'Capture last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'identify bonus variables
        Dim greatest_decrease As Double
        Dim greatest_total_volume As Double

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
        Open_Price = ws.Cells(2, 3).Value
    
        'Loop through all stocks
        For i = 2 To last_row
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                          
                'Calculate Yearly Change and Percent Change
                Close_Price = ws.Cells(i, 6).Value
                yearly_change = (Close_Price - Open_Price)
                'percent_change = (yearly_change / Open_Price)
                If Open_Price <> 0 Then
                    percent_change = (yearly_change / Open_Price)
                Else
                    percent_change = 0
                End If
            
                'Add Total Stock Volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
                'Print calculations to the summary table and add color formatting
                ws.Range("I" & Summary_Table_Row).Value = ticker
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                    If (yearly_change > 0) Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percent_change)
                ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
    
                'Add one to summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
                'reset variables to zero and capture next open price
                total_stock_volume = 0
                yearly_change = 0
                Close_Price = 0
                Open_Price = ws.Cells(i + 1, 3).Value
                
            
            'If the cell immediately following a row is the same ticker.
            Else
        
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
    Next ws

End Sub

