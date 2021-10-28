Sub StockSummary()
    
    ' Determine last row in worksheet
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Assign variables
    Dim TickerName As String
    Dim OpenPrice As Double
    Dim ClosingPrice As Double
    Dim TotalStockVolume As LongLong
    Dim TickerFirstRow As Long
    Dim TickerLastRow As Long
    Dim SummaryRow As Double

    TickerFirstRow = 2
    SummaryRow = 2
    
    'Name headers fields for summary
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Format summary rows
    Range("J" & 2 & ":K" & LastRow).NumberFormat = "0.00"
    Range("K" & 2 & ":K" & LastRow).NumberFormat = "0.00%"
    Range("L" & 2 & ":L" & LastRow).NumberFormat = "#,###,###,###,###"
    

    'Loop through stock names (REF: AA = 262; AAC = 525)
        
        'Check if stock name is the same, if not the same then fill in summary
        For i = 2 To LastRow
        
                'check if tickername has changed and identify opening price
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                            
                        'List ticker name in summary
                        TickerName = Cells(i, 1).Value
                        Cells(SummaryRow, 9).Value = TickerName
                                                                  
                        'Opening Price identify
                        OpeningPrice = Cells(TickerFirstRow, 3).Value
                            'TESTCODE Cells(SummaryRow, 10).Value = OpeningPrice
                        
                        'Closing Price identify
                        TickerLastRow = i
                        ClosingPrice = Cells(TickerLastRow, 6).Value
                            'TESTCODE Cells(SummaryRow, 11).Value = ClosingPrice
                        
                        'Yearly Change Calculate and List
                        Cells(SummaryRow, 10) = ClosingPrice - OpeningPrice
                        
                            'Format interior color
                            If Cells(SummaryRow, 10) > 0 Then
                                Cells(SummaryRow, 10).Interior.ColorIndex = 4
                                
                                Else: Cells(SummaryRow, 10).Interior.ColorIndex = 3
                                
                            End If
                        
                        'Percent change calculate from opening to closing
                        Cells(SummaryRow, 11) = (ClosingPrice / OpeningPrice) - 1
                        
                        'Total stock volume calculate
                        TotalStockVolume = Application.WorksheetFunction.Sum(Range("G" & TickerFirstRow & ":G" & TickerLastRow))
                            Cells(SummaryRow, 12) = TotalStockVolume
                      
                        'Advance and reset for next ticker i summary
                        TickerFirstRow = i + 1
                        SummaryRow = SummaryRow + 1
                        TotalStockVolume = 0
                       
                        
                End If
        Next i
        
        
        'Check if stock name is the same, if not the same then fill in summary
        

End Sub
            


        
