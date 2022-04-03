Attribute VB_Name = "Module1"
Sub stock_summary()

    'identify variables
    Dim ws As Worksheet
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
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'identify bonus variables
    Dim greatest_decrease As Double
    Dim greatest_total_volume As Double


    
    'Set Summary Range Headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    'Loop through all stocks
    For i = 2 To last_row
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            
            'Calculate yearly change and percent change
            Open_Price = Cells(2, 3).Value
            close_price = Cells(i, 6).Value
            yearly_change = (close_price - Open_Price)
            percent_change = (close_price - Open_Price) / Open_Price
            
            'Add total stock volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
            'Print calculations to the summary table and add color formatting
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = yearly_change
                If (yearly_change > 0) Then
                    Range("J" & summary_table_row).Interior.ColorIndex = 4
                Else
                    Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
            Range("K" & summary_table_row).Value = percent_change
            Range("L" & summary_table_row).Value = total_stock_volume
    
            'Add one to summary table row
            summary_table_row = summary_table_row + 1
            
            'reset variables to zero
            total_stock_volume = 0
            
        ' If the cell immediately following a row is the same ticker.
        
        Else
        
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
        End If
        
    Next i

End Sub
