Sub StockSummary_PtII()
            
    'Loop through all sheets
    For Each ws In Worksheets
    
        ' Determine last row in worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Assign variables
        Dim TickerName As String
        Dim OpenPrice As Double
        Dim ClosingPrice As Double
        Dim TotalStockVolume As LongLong
        Dim TickerFirstRow As Long
        Dim TickerLastRow As Long
        Dim SummaryRow As Long
    
        'Reset Variables for next ticker
        TickerFirstRow = 2
        SummaryRow = 2
        TickerLastRow = 0
           
        
        
        'Name headers fields for summary
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
    
        'Loop through stock names
            
            'Check if stock name is the same, if not the same then fill in summary
            For i = 2 To LastRow
            
                    'check if tickername has changed and identify opening price
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                
                            'List ticker name in summary
                            TickerName = ws.Cells(i, 1).Value
                            ws.Cells(SummaryRow, 9).Value = TickerName
                                                                      
                            'Opening Price identify
                            OpeningPrice = ws.Cells(TickerFirstRow, 3).Value
                                'TESTCODE Cells(SummaryRow, 10).Value = OpeningPrice
                            
                            'Closing Price identify
                            TickerLastRow = i
                            ClosingPrice = ws.Cells(TickerLastRow, 6).Value
                                'TESTCODE Cells(SummaryRow, 11).Value = ClosingPrice
                            
                            'Yearly Change Calculate and List
                            ws.Cells(SummaryRow, 10) = ClosingPrice - OpeningPrice
                            
                                'Format interior color
                                If ws.Cells(SummaryRow, 10) > 0 Then
                                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                                    
                                    Else: ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                                    
                                End If
                            
                                'Percent change calculate from opening to closing
                                If OpeningPrice = 0 Then
                                    ws.Cells(SummaryRow, 11) = "No opening price at the beginning of year"
                                    
                                    Else
                                    ws.Cells(SummaryRow, 11) = (ClosingPrice / OpeningPrice) - 1
                                
                                End If
                            
                            'Total stock volume calculate
                            TotalStockVolume = Application.WorksheetFunction.Sum(ws.Range("G" & TickerFirstRow & ":G" & TickerLastRow))
                                ws.Cells(SummaryRow, 12) = TotalStockVolume
                          
                            'Advance and reset for next ticker i summary
                            TickerFirstRow = i + 1
                            SummaryRow = SummaryRow + 1
                            TotalStockVolume = 0
                           
                            
                    End If
            Next i
        
            'Format summary rows
            ws.Range("J" & 2 & ":K" & LastRow).NumberFormat = "0.00"
            ws.Range("K" & 2 & ":K" & LastRow).NumberFormat = "0.00%"
            ws.Range("L" & 2 & ":L" & LastRow).NumberFormat = "#,###,###,###,###"
                    
            '-------------------------------------------
            '------------BONUS--------------------------
            '-------------------------------------------
                    
                'Name headers fields for bonus summary
                ws.Cells(1, 15).Value = "Ticker"
                ws.Cells(1, 16).Value = "Value"
                ws.Cells(2, 14).Value = "Greatest % Increase"
                ws.Cells(3, 14).Value = "Greatest % Decrease"
                ws.Cells(4, 14).Value = "Greatest Total Volume"
                
                ' Determine last row in summary
                
                SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                                
                'Set variables and ranges for bonus summary
                Dim PMax As Double
                Dim PMin As Double
                Dim TMax As LongLong
                Dim PMaxRow As Long
                Dim PMinRow As Long
                Dim TMaxRow As Long
                            
                'Find Values for greatest % increase and decrease and greatest volume otal
                PMax = Application.WorksheetFunction.Max(ws.Range("K" & 2 & ":K" & SummaryLastRow))
                PMin = Application.WorksheetFunction.Min(ws.Range("K" & 2 & ":K" & SummaryLastRow))
                TMax = Application.WorksheetFunction.Max(ws.Range("L" & 2 & ":L" & SummaryLastRow))
                                    
                'Fill in values
                ws.Cells(2, 16).Value = PMax
                ws.Cells(3, 16).Value = PMin
                ws.Cells(4, 16).Value = TMax
                
                'Format values
                ws.Cells(2, 16).NumberFormat = "0.00%"
                ws.Cells(3, 16).NumberFormat = "0.00%"
                ws.Cells(4, 16).NumberFormat = "#,###,###,###,###"
                
                
                'Find Tickers for greatest % increase and decrease and greatest volume otal
                For i = 2 To SummaryLastRow
                
                    If ws.Cells(i, 11).Value = PMax Then
                        ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                    
                    End If
                Next i
                
                For j = 2 To SummaryLastRow
                
                    If ws.Cells(j, 11).Value = PMin Then
                        ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
                    
                    End If
                Next j
                
                For k = 2 To SummaryLastRow
                
                    If ws.Cells(k, 12).Value = TMax Then
                        ws.Cells(4, 15).Value = ws.Cells(k, 9).Value
                    
                    End If
                Next k
                
            
            '-------------------------------------------
            '------------END BONUS END------------------
            '-------------------------------------------
                        
            'Worksheet complete. Reset TickerLastRow
            TickerLastRow = 0
        
        Next ws
            
    
    End Sub
                
    
    
            
