Attribute VB_Name = "Module1"
Sub s_tick()

For Each ws In Worksheets
        
        Dim Worksheetname As String
        
        Worksheetname = ws.Name
        
        
        'set variable to hold the ticker symbol
        Dim symbol As String
        
        ' Set an initial variable for holding the total per symbol, opening annual stock price, closing annual stock price, percent change.
        Dim open_price, close_price, symbol_total, pchange, ychange As Double
        
    
        symbol_total = 0
        
        open_price = Cells(2, 3).Value
        
        close_price = 0
        
        
        'track the location for each ticker symbol in the summary table(sumTab)
        Dim sumTab As Long
        
        sumTab = 2
        
        'determine the range of the dataset
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
        'create loop using the lastRow of the dataset
        For i = 2 To lastRow
            
            'check if the symbol remains the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'set symbol and closing price
                symbol = ws.Cells(i, 1).Value
                
                close_price = ws.Cells(i, 6).Value
                
                'calculate yearly change($ & %)
                ychange = (close_price - open_price)
                
                If open_price = 0 Then
                Else
                    pchange = (ychange / open_price)
                End If
                                
                '------------------------------------
                ' SUMMARY TABLE
                '------------------------------------
                
                'print the symbol in the sumTab
                ws.Range("J" & sumTab).Value = symbol
                
                'print the annual change in price
                ws.Range("K" & sumTab).Value = ychange
                
                ws.Range("K" & sumTab).NumberFormat = "0.00"
               
               'if statement to determine the interior color of the yearly change range
                If ychange < 0 Then
                    ws.Range("K" & sumTab).Interior.ColorIndex = 3
                Else
                    ws.Range("K" & sumTab).Interior.ColorIndex = 4
                End If
                
                'print the annual percent change and format to percentage w 2 decimal places
                ws.Range("L" & sumTab).Value = pchange
                
                ws.Range("L" & sumTab).NumberFormat = "0.00%"
                
                'Print the Symbol Total to the Summary Table and format as currency
                ws.Range("M" & sumTab).Value = symbol_total + Cells(i, 7).Value
                
                ws.Range("M" & sumTab).Style = "Currency"
                
                'Change sumTab row
                sumTab = sumTab + 1
                
                'reset symbol total & close price then set new open_price
                symbol_total = 0
                
                close_price = 0
                
                pchange = 0
                
                ychange = 0
                
                open_price = ws.Cells(i + 1, 3).Value
                
                                
            Else
                'sums the total volume for the stock symbol
                symbol_total = symbol_total + ws.Cells(i, 7).Value
            
            End If
            
            Next i
            
        Next ws
        
    End Sub

