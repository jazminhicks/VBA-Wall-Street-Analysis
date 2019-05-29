Sub stocktest()

    
    For Each ws In Worksheets
        
        ' Identify variables
        
        Dim tickername As String

        Dim tickername2 As String
         
        Dim totalvol As Double
        
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        Dim openingprice As Double
        openingprice = ws.Cells(2, 3).Value 'set initial opening price
        
        Dim closingprice As Double
        
        Dim yearly_price_change As Double
        
        Dim percent_change As Double

        Dim greatest_increase As Double
        greatest_increase = 0
        
        Dim greatest_decrease As Double
        greatest_decrease = 0

        Dim greatest_volume As Double
        greatest_volume = 0
        
        ' find last row in each sheet.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Label new columns and Rows
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' Loop through each row, group data of the same ticker
        For I = 2 To LastRow

            'when the next ticker is different, then...
            
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
         
                tickername = ws.Cells(I, 1).Value 'save previous ticker name'
                
                totalvol = totalvol + ws.Cells(I, 7).Value 'save current total volume'
                
                closingprice = ws.Cells(I, 6).Value 'set closing price
                
                yearly_price_change = closingprice - openingprice  'calculate yearly change and save to variable

                
                'calculate percent change and save to variable
                If openingprice > 0 Then
                    percent_change = (yearly_price_change) / openingprice
                    
                
                Else
                    percent_change = 0 'temporarily setting value to zero, will try to edit to "N/A" later
                        
                End If
                

                ' add information to the summary table
                
                ws.Cells(summary_table_row, 9).Value = tickername
                
                ws.Cells(summary_table_row, 10).Value = yearly_price_change
                'ws.Cells(summary_table_row, 10).NumberFormat = "0.00000000#" 'change number format to include more decimal places
        


                'change cell color based on pos, neg values
                If yearly_price_change < 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                
                End If

                ws.Cells(summary_table_row, 11).Value = percent_change
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%" 'change number format to percentage
                



                ws.Cells(summary_table_row, 12).Value = totalvol
                
                summary_table_row = summary_table_row + 1 'reset summary table to next row
                
                totalvol = 0 'reset total volume
                
                openingprice = ws.Cells(I + 1, 3).Value 'reset opening price to first value of the next ticker

            Else
            ' Otherwise, as long as the ticker value is the same, then continue to
            ' add to the total stock volume.
            
                totalvol = totalvol + ws.Cells(I, 7).Value
            
            End If

        ' adjust column size to fit data
        
        
        Next I
        
' PART 3
        'Loop through the summary table to find the greatest % increase, decrease and largest volume
        For j = 2 To LastRow
            
            'if the cell value is positive and larger than the current greatest increase (initial = 0), 
            'then change the greatest increase to that value, place the value and tickername into the new table
            'continue... Same for decrease. 
            If ws.Cells(j, 11).Value > 0 And ws.Cells(j, 11).Value > greatest_increase Then
                greatest_increase = ws.Cells(j, 11).Value
                tickername2 = ws.Cells(j, 9).Value
                ws.Cells(2, 16).Value = tickername2
                ws.Cells(2, 17).Value = greatest_increase
                ws.Cells(2, 17).NumberFormat = "0.00%" ' change number formatting to percentage
            
            ElseIf ws.Cells(j, 11).Value < 0 And ws.Cells(j, 11).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(j, 11).Value
                tickername2 = ws.Cells(j, 9).Value
                ws.Cells(3, 16).Value = tickername2
                ws.Cells(3, 17).Value = greatest_decrease
                ws.Cells(3, 17).NumberFormat = "0.00%" ' change number formatting to percentage
            
            End If
        
        Next j
        
        ' created a separate loop for largest volume just in case it shared a ticker with greatest increase
        For k = 2 To LastRow
            If ws.Cells(k, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(k, 12).Value
                tickername2 = ws.Cells(k, 9).Value
                ws.Cells(4, 16).Value = tickername2
                ws.Cells(4, 17).Value = greatest_volume
            End If
        Next k
        
        

    ws.Columns("I:S").AutoFit 'adjust column width to fit data
    

    Next ws     

    
End Sub

