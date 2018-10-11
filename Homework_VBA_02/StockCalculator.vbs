Sub StockVolumeCalc()


    'For each worksheet
    For Each ws In Worksheets
    
        'Last Aggregate Row
        Dim LastAggRow As Integer
        Dim TotalVolume As Double
        Dim TickerFirstValue As Double 'to hold starting Row Number for each Stock
     
    '   Get total Last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Label headers for Aggregates
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Set initial Aggregate row
        LastAggRow = 2
        
        'Set initial FirstValue (row) for Stock
        TickerFirstValue = 2
        
        'Set Total Volume to 0
        TotalVolume = 0
        
    
        'Aggregate Stock volume per stock ticker
        For i = 2 To lastRow
            
            
            ' If the stock ticker is different in the next row, display ticker name & total volume
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                
                'Display Ticker symbol
                ws.Cells(LastAggRow, 9) = ws.Cells(i, 1).Value
                
                'Diplay total Volume
                ws.Cells(LastAggRow, 12) = TotalVolume + ws.Cells(i, 7).Value
                
                
                
                'Display Yearly Change
                ws.Cells(LastAggRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(TickerFirstValue, 3).Value
                'Format the background color
                If (ws.Cells(LastAggRow, 10).Value > 0) Then
                    ws.Cells(LastAggRow, 10).Interior.ColorIndex = 4 'Green
                Else
                    ws.Cells(LastAggRow, 10).Interior.ColorIndex = 3 'Red
                End If
                
                
                'Display Percent Change (((orig start value - new end value)/ orig start value) for that ticker)
                'original start value => (ws.Cells(TickerFirstValue, 3).Value)
                'new end value =>(ws.Cells(i, 3).Value))
                ' Included if statement incase of any 0 values in the stock price
                If (ws.Cells(TickerFirstValue, 3).Value) <> 0 Then
                    ws.Cells(LastAggRow, 11).Value = Format(((ws.Cells(i, 3).Value) - (ws.Cells(TickerFirstValue, 3).Value)) / (ws.Cells(TickerFirstValue, 3).Value), "Percent")
                Else
                    ws.Cells(LastAggRow, 11).Value = 0
                End If
                
                'Increment the Aggregate row #
                LastAggRow = LastAggRow + 1
                
                'Update FirstValue (row) of the next stock
                TickerFirstValue = i + 1
                
                
                'Reset TotalVolume
                TotalVolume = 0
                
             Else
             
                 ' Add Volume of the current row to total
                 TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
               
               
            End If
                          
        Next i
        
        
        'ws.Range("K2:K" & (LastAggRow)).NumberFormat = "0.00"
        'ws.Range("K2:K" & (LastAggRow)).Style = "Percent"
   
        'MsgBox (Application.WorksheetFunction.Max(Columns("K")))
        
        '--For each worksheet, identify the following and their ticker symbol
        ' (1) greatest % increase
        ' (2) greatest % decrease
        ' (3) greatest stock volume
        
        'get row count for column K
        ' if value in Column J is -ve or background is red then calc greatest decrease % percent
        ' else greatest increase %
        
        ' Get row count for column L
        ' calc max stock volume
        
        ' Subtract by 1 to undo the last increment
        LastAggRow = LastAggRow - 1
        
        'Initialize positive and negative percent values
        greatestPosIndex = 2
        greatestNegIndex = 2
        greatestStockVolIndex = 2  'first stock volume (row #2)
        
        For x = 2 To LastAggRow
        
            If Cells(x, 10).Value > 0 Then 'its +ve change
            
                If ws.Cells(greatestPosIndex, 11).Value < ws.Cells(x, 11).Value Then   'current %change is greater than prev. change
                    greatestPosIndex = x
                End If
            
            Else 'its -ve change
                
                If ws.Cells(greatestNegIndex, 11).Value > ws.Cells(x, 11).Value Then   'current %change is lesser than prev. change
                    greatestNegIndex = x
                End If
            
            End If
            
            ' calculate greatest stock volume
             If ws.Cells(greatestStockVolIndex, 12).Value < ws.Cells(x, 12).Value Then    'current volume is greater than prev. change
                    greatestStockVolIndex = x
            End If
             
           ' Stop
        
        Next x
        
        
        
        'Headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Row data
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        'ws.Columns("O1:O5").AutoFit
        
        'Greatest +ve Percent increase
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(2, 17).NumberFormat = "0.00"
        ws.Cells(2, 16).Value = ws.Cells(greatestPosIndex, 9).Value 'Ticker symbol
        ws.Cells(2, 17).Value = ws.Cells(greatestPosIndex, 11).Value '% increase
       
        
        
        
        'Greatest -ve Percent increase
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(3, 17).NumberFormat = "0.00"
        ws.Cells(3, 16).Value = ws.Cells(greatestNegIndex, 9).Value 'Ticker symbol
        ws.Cells(3, 17).Value = ws.Cells(greatestNegIndex, 11).Value '% increase
        
        
        
        'Greatest -total stock volume
        ws.Cells(4, 16).Value = ws.Cells(greatestStockVolIndex, 9).Value 'Ticker symbol
        ws.Cells(4, 17).Value = ws.Cells(greatestStockVolIndex, 12).Value '% increase
        
       
    Next ws
    
End Sub
