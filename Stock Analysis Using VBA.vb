Sub MultiYearStock()

For Each ws In Worksheets

    Dim ticker As String
    
    Dim tickerStartRow As Long
    tickerStartRow = 2
    
    Dim volume, openValues, closeValues As Double
    volume = 0
    
    Dim yearlyChange As Double
    yearlyChange = 0
    
    Dim percentChange As Double
    percentChange = 0
    
    'Dim openValues, closeValues As Double
    openValues = 0
    closeValues = 0
    
    Dim summaryRow As Integer
    summaryRow = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Column Header for Ticker and Total Stock Volume
    
        ws.Range("I1") = "Ticker"
        
        ws.Range("J1") = "Yearly Change"
        
        ws.Range("K1") = "Percent Change"
        
        ws.Range("L1") = "Total Stock Volume"
    
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        
            'Display unique Ticker Name in the Summary Table
            
            ticker = ws.Cells(i, 1)
            
            ws.Range("I" & summaryRow) = ticker
            
            
            'Pulling the open value of the year by Ticker
            openValues = ws.Cells(tickerStartRow, 3)
            ws.Range("M" & summaryRow) = openValues
            
            
            'Pulling the close value of the year by Ticker
            closeValues = ws.Cells(i, 6)
            ws.Range("N" & summaryRow) = closeValues
            
            'Calculate Yearly Change
            yearlyChange = closeValues - openValues
            
            'Display Yearly Changes
            ws.Range("J" & summaryRow) = yearlyChange
            
            
            'Calculate Percent Change
            If openValues > 0 Then
            
                percentChange = (closeValues - openValues) / openValues
                
            Else
            
                percentChange = 0
                
            End If
            
            If yearlyChange > 0 Then
            
                ws.Range("J" & summaryRow).Interior.ColorIndex = 4
                
            Else
            
                ws.Range("J" & summaryRow).Interior.ColorIndex = 3
            
            End If
            
            'Display Percent Change
            ws.Range("K" & summaryRow) = percentChange
            
            'Calculate Stock Volume
            volume = volume + ws.Cells(i, 7)
            
            'Display Stock Volume in the summary table
            ws.Range("L" & summaryRow) = volume
            
            'Incrementing the row number for the summary table
            summaryRow = summaryRow + 1
            
            tickerStartRow = i + 1
            
            volume = 0
            openValues = 0
            closeValues = 0
            percentChange = 0
            
        Else
        
            volume = volume + ws.Cells(i, 7)
            
            
        End If
        
        
    Next i
     
    
    'Format Percent Change from Number to Percent
    ws.Range("K1", ws.Cells(Rows.Count, 11)).NumberFormat = "0.00%"
      
      
    

Next ws

End Sub

