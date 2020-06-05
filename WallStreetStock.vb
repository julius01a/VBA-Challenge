Sub WallStreetStock()

For Each ws In Worksheets
    'setting variables to fetch stock data
    Dim ticker As String
    Dim openValue As Double
    Dim closeValue As Double
    
    
    
    'settig summary table variables
    Dim outTableTicker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalStockVolume As Double
    Dim tableRow As Integer
    tableRow = 2
    totalStockVolume = 0
    
    
    'setting summary table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    
    'initialising the loop for tickers
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    openValue = ws.Cells(2, 3).Value
    
    'starting the loop for tickers
    For r = 2 To LastRow
        
        'if ticker is the same then keep adding volume
        totalStockVolume = totalStockVolume + ws.Cells(r, 7).Value
        
        'if ticker changes
        If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
           ticker = ws.Cells(r, 1).Value
           ws.Range("I" & tableRow) = ticker
           
         'define close values and yearly change
           closeValue = ws.Cells(r, 6).Value
           yearlyChange = closeValue - openValue
           
         'total stock volume
           ws.Range("J" & tableRow).Value = yearlyChange
           
         'working on yearly percentage change
         '---this has bugged me for so long, we have to eliminate when open vale = 0!!!!!----
         
           If openValue = 0 Then
              percentChange = 0
           Else
              percentChange = yearlyChange / openValue
           End If
        
           ws.Range("K" & tableRow).Value = percentChange
           ws.Range("K" & tableRow).NumberFormat = "0.00%"
           
           ws.Range("L" & tableRow).Value = totalStockVolume
           
         'reset values for the next ticker
           openValue = ws.Cells(r + 1, 3).Value
           totalStockVolume = 0
           tableRow = tableRow + 1
           yearlyChange = 0
           percentChange = 0
        End If
        
    Next r
    
    'setting report table colour
    
    LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    For r = 2 To LastRow
    
        If ws.Cells(r, 10).Value >= 0 Then
            ws.Cells(r, 10).Interior.ColorIndex = 4
        
        Else
            ws.Cells(r, 10).Interior.ColorIndex = 3
        End If
        
    Next r
      
      
      
    'challenges
    'setting variables
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestTotalVolume As Double
    Dim minRow As Integer
    Dim maxRow As Integer
    Dim totalVolumeRow As Double
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestTotalVolume = 0
    
    'setting headers & titles
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
      
    For r = 2 To LastRow
        'find the greatest increase
        If ws.Cells(r, 11).Value > greatestIncrease Then
            greatestIncrease = ws.Cells(r, 11).Value
            maxRow = r
        End If
        
        'find the greatest decrease
        If ws.Cells(r, 11).Value < greatestDecrease Then
            greatestDecrease = ws.Cells(r, 11).Value
            minRow = r
        End If
        
        'find the max total volume
        If ws.Cells(r, 12).Value > greatestTotalVolume Then
            greatestTotalVolume = ws.Cells(r, 12).Value
            totalVolumeRow = r
        End If
                    
    Next r
            
        
    'fill the ticker name
    ws.Range("O2") = ws.Cells(maxRow, 9).Value
    ws.Range("O3") = ws.Cells(minRow, 9).Value
    ws.Range("O4") = ws.Cells(totalVolumeRow, 9).Value
        
        
    'fill the extreme values
    ws.Range("P2") = greatestIncrease
    ws.Range("P3") = greatestDecrease
    ws.Range("P4") = greatestTotalVolume
            
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
        
      
Next ws

End Sub