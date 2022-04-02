Attribute VB_Name = "Module1"
Sub stock_summary()

    'identify variables
    
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    total_stock_volume = 0
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total_volume As Double
    Dim sum_table_row As Integer
    sum_table_row = 2
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set Summary Range Headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    'Loop through all stocks
    For i = 2 To last_row
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            
            'Add total stock volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
            'Print the stock ticker and total stock volume in the summary table
            Range("I" & sum_table_row).Value = ticker
            Range("L" & sum_table_row).Value = total_volume
            
            'Add one to summary table row
            sum_table_row = sum_table_row + 1
            
            'reset the total stock volume
            total_volume = 0
            
        ' If the cell immediately following a row is the same ticker.
        
        Else
        
            total_volume = total_volume + Cells(i, 7).Value
            
        End If
        
    Next i

End Sub
