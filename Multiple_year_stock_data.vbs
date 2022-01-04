Sub multiple_year_stock_data()
    Dim ws As Worksheet
    
    For Each ws In Worksheets

        
        Dim ticker As String
        
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        
        Dim percent_change As Double
        Dim stock_volume As Double
        stock_volume = 0
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Get the number of rows for the dataset
        Dim rows As Long
        rows = ws.Range("A2").End(xlDown).Row
        
        opening_price = ws.Cells(2, 3).Value    'Initial price of 1st listed ticker
        
        For i = 2 To rows
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker = ws.Cells(i, 1).Value
                'Print the Ticker in the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = ticker
                
                'Determining the Yearly Change for each Ticker Symbol
                    closing_price = ws.Cells(i, 6)
                    yearly_change = closing_price - opening_price
                    ws.Range("M" & Summary_Table_Row).Value = yearly_change
                    
                'Determining the Percent Change for each Ticker Symbol
                    If opening_price = 0 Then
                        percent_change = 0
                        ws.Range("N" & Summary_Table_Row).Value = percent_change
                    Else
                        percent_change = (yearly_change / opening_price)
                        ws.Range("N" & Summary_Table_Row).Value = percent_change
                    End If
                    
                'Determining the Total Stock Volume by Ticker Symbol
                    'Add to the Stock Volume
                    stock_volume = stock_volume + ws.Cells(i, 7).Value
                    'Print the Stock Volume to the Summary Table
                    ws.Range("O" & Summary_Table_Row).Value = stock_volume
                    
                    'Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                     
                    
                'Re-initialize Yearly Change, Closing Price, Stock Volume for new ticker
                    yearly_change = 0
                    closing_price = 0
                    stock_volume = 0
                'Set Opening price for new ticker
                    opening_price = ws.Cells(i + 1, 3).Value

            Else
                    stock_volume = stock_volume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        'Conditional formatting to highlight Yearly Change
        Dim length_of_change As Long
        length_of_change = ws.Range("M2").End(xlDown).Row
        For j = 2 To length_of_change
            If ws.Cells(j, 13) < 0 Then
                'Highlights negative yearly changes in red
                ws.Cells(j, 13).Interior.ColorIndex = 3
            ElseIf ws.Cells(j, 13) > 0 Then
                'HIghlights positive yearly changes in green
                ws.Cells(j, 13).Interior.ColorIndex = 4
            Else
                'This is for the situation if the yearly change is equal to zero
                ws.Cells(j, 13).Interior.ColorIndex = 6        'Highlight is yellow
            End If
        Next j
        
        ws.Range("L1").Value = "Ticker Symbol"
        ws.Range("M1").Value = "Yearly Change"
        ws.Range("N1").Value = "Percent Change"
        ws.Range("O1").Value = "Total Stock Volume"

        ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
        ws.Range("R1").Value = "Ticker"
        ws.Range("S1").Value = "Value"
        
        ws.Range("L1:S1").Columns.AutoFit
        ws.Range("Q2:Q4").Columns.AutoFit
        ws.Range("R2:R4").Columns.AutoFit
        ws.Range("S2:S4").Columns.AutoFit
                

        'Finding the Greatest % Increase and associated ticker
        max_increase = Application.WorksheetFunction.Max(ws.Range("N:N"))
        ws.Range("S2") = max_increase
        For k = 2 To length_of_change
            If ws.Cells(k, 14) = max_increase Then
                ws.Range("R2").Value = ws.Cells(k, 12)
            End If
        Next k
            
        'Finding the Greatest % Decrease and associated ticker
        max_decrease = Application.WorksheetFunction.Min(ws.Range("N:N"))
        ws.Range("S3") = max_decrease
        For k = 2 To length_of_change
            If ws.Cells(k, 14) = max_decrease Then
                ws.Range("R3").Value = ws.Cells(k, 12)
            End If
        Next k
        
        'Finding the Greatest Total Volume and associated ticker
        max_vol = Application.WorksheetFunction.Max(ws.Range("O:O"))
        ws.Range("S4") = max_vol
        For k = 2 To length_of_change
            If ws.Cells(k, 15) = max_vol Then
                ws.Range("R4").Value = ws.Cells(k, 12)
            End If
        Next k
        
        'Format Percent Change column, Greatest % Increase cell, & Greatest % Decrease cell as percentages
        ws.Range("N:N").NumberFormat = "0.00%"
        ws.Range("S2:S3").NumberFormat = "0.00%"
        
    
    Next ws
        

End Sub
