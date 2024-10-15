Attribute VB_Name = "Module1"
Sub ticker_stock()


    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
    For Each ws In Worksheets

    
    ' --------------------------------------------
    ' RETRIVAL OF DATA
    ' --------------------------------------------
        '  Create a variables

        Dim ticker_name As String
        Dim open_price As Double
        Dim close_price As Double
        Dim quaterly_change As Double
        Dim percent_change As Double
        
        '  Variable to hold a total count on the total volume of trade
        
        Dim ticker_volume As Double
        
        '  Initially set the ticker_volume to be 0 for each row

        ticker_volume = 0

        ' Location for each ticker name in the summary table
        
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        ' Set initial open_price
        
        open_price = ws.Cells(2, 3).Value
        
        ' Create Table titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quaterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Determine the Last Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    ' --------------------------------------------
    ' COLUMN CREATION
    ' --------------------------------------------

        ' Loop through all rows by the ticker name
        
        For i = 2 To lastrow

            'Searches for when the value of the next cell is different than that of the current cell

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              ' Set the ticker name
              ticker_name = ws.Cells(i, 1).Value

              ' Add the ticker volume
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value

              ' Print the ticker name in the summary table
              ws.Range("I" & summary_ticker_row).Value = ticker_name

              ' Print the trade volume for each ticker in the summary table
              ws.Range("L" & summary_ticker_row).Value = ticker_volume

              ' Information about closing price
              close_price = ws.Cells(i, 6).Value

              ' Calculate quaterly change
               quaterly_change = (close_price - open_price)
              
              'Print the quaterly change for each ticker in the summary table
              ws.Range("J" & summary_ticker_row).Value = quaterly_change

              'Check for the non-divisibilty condition when calculating the percent change
                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = quaterly_change / open_price
                
                End If

              'Print the quaterly change for each ticker in the summary table
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              ' Reset the counter. Add one to the summary ticker row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              ticker_volume = 0

              'Reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    ' --------------------------------------------
    ' CONDITIONAL FORMATTING
    ' --------------------------------------------

    'Conditional formatting that will highlight positive change in green and negative change in red
    
    'First find the last row of the summary table

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color quaterly change
       For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
       Next i


        ' --------------------------------------------
        ' THE STOCK PRICE CHANGES
        ' --------------------------------------------
        
    ' Labeling the cells according to the assignment

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"


    'Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume
    
    
        For i = 2 To lastrow_summary_table
        
            'Greatest % increase
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            ' Greatest % decrease
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ' Greatest total volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i

 
    Next ws
        
End Sub