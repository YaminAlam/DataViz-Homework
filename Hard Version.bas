Attribute VB_Name = "Module11"
Sub VolumeTotal():
      
    
    'this loop will go through each of the worksheets
    For Each Ws In Worksheets
    
        'defining our header variables
        Dim tickerHeader As String
        Dim tickerVolHeader As String
        Dim yearlyChangeHeader As String
        Dim percentChangeHeader As String
        Dim valueHeader As String
        
        
        'defining variables for loop that will process the data table
        Dim currentTicker As String
        Dim prevTicker As String
        Dim nextTicker As String
        Dim lastRow As Long
        Dim currentTickerVol As LongLong
        Dim stockCount As Integer
        Dim stockDate As Long
        Dim maxPercChange
        Dim maxPercChangeTicker
        Dim minPercChange
        Dim minPercChangeTicker
        Dim maxVol
        Dim maxVolTicker
       
        
        'defining variables that will be used to relay information
        Dim stockVol As LongLong
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percChange As Double
    
        'giving preliminary values
        stockCount = 1
        tickerHeader = "Ticker"
        tickerVolHeader = "Total Stock Value"
        yearlyChangeHeader = "Yearly Change"
        percentChangeHeader = "Percent Change"
        lastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        maxPercChange = 0
        minPercChange = 0
        maxVol = 0
        valueHeader = "Value"
        
        
        'inserting headers
        Ws.Cells(1, 9).Value = tickerHeader
        Ws.Cells(1, 10).Value = yearlyChangeHeader
        Ws.Cells(1, 11).Value = percentChangeHeader
        Ws.Cells(1, 12).Value = tickerVolHeader
        Ws.Cells(1, 16).Value = tickerHeader
        Ws.Cells(1, 17).Value = valueHeader
        Ws.Cells(2, 15).Value = "Greatest % Increase"
        Ws.Cells(3, 15).Value = "Greatest % Decrease"
        Ws.Cells(4, 15).Value = "Greatest Total Volume"

    
        'this loop will go through the rows of the current sheet to retrieve the stock names and their volumes
        For rowNum = 2 To lastRow
            
            'giving values to the ticker variables that will be used when the loop is processing all the information
            currentTicker = Ws.Cells(rowNum, 1).Value
            prevTicker = Ws.Cells(rowNum - 1, 1).Value
            nextTicker = Ws.Cells(rowNum + 1, 1).Value
            currentTickerVol = Ws.Cells(rowNum, 7).Value
            
    
            'creating conditional that depends on whether current ticker matches previous ticker
            If currentTicker <> prevTicker Then
            
                'if the current ticker does not match with the previous ticker, retrieve its volume and title to relay
                stockCount = stockCount + 1
                
                'saving the opening price as that would be the first instance of the stock
                openingPrice = Ws.Cells(rowNum, 3).Value
                
                'relaying the Stock name and Volume
                Ws.Cells(stockCount, 9).Value = currentTicker
                stockVol = currentTickerVol
                Ws.Cells(stockCount, 12).Value = stockVol
            Else
                
                'if the current ticker does match the previous ticker, add its volume onto the total stock volume
                stockVol = stockVol + currentTickerVol
                Ws.Cells(stockCount, 12).Value = stockVol
                
                
                'if the current ticker does not match the next ticker, then the stock has finished and we have gone on to the next one
                If currentTicker <> nextTicker Then
                
                    'since the nextTicker is a new Stock, we can be confident that the closing price of the current Ticker is the closing price of the Stock
                    closingPrice = Ws.Cells(rowNum, 6).Value
                    
                    'with the opening Price already retreived, we can calculate yearly change and percent change now
                    yearlyChange = closingPrice - openingPrice
                    
                    'calculating the percent change while keeping in mind that division by zero is not possible
                    If openingPrice <> 0 Then
                        percChange = Round((yearlyChange / openingPrice) * 100, 2)
                    Else
                        percChange = 0
                    End If
                    Ws.Cells(stockCount, 10) = yearlyChange
                    
                    'if yearly change is positive, the cell color is green--if it is negative than the cell color is red;
                    'a change of 0 will be left white--I'M LOOKING AT YOU PLNT
                    
                    If Ws.Cells(stockCount, 10).Value > 0 Then
                        Ws.Cells(stockCount, 10).Interior.ColorIndex = 4
                    ElseIf Ws.Cells(stockCount, 10).Value < 0 Then
                        Ws.Cells(stockCount, 10).Interior.ColorIndex = 3
                    Else
                    End If
                    
                    Ws.Cells(stockCount, 11).Value = percChange & "%"
                
                Else
                    
                    
                End If
                    
            End If
            'this will record the highest perc change and replace it if a higher change is found
            'it will also retreive the ticker of that change
            If percChange > maxPercChange Then
                maxPercChange = percChange
                Ws.Cells(2, 17).Value = maxPercChange
                maxPercChangeTicker = Ws.Cells(stockCount, 9).Value
                Ws.Cells(2, 16).Value = maxPercChangeTicker
            Else
            End If
                    
            'similar logic to the highest perc change version, but this time it looks for lower values
            If percChange < minPercChange Then
                minPercChange = percChange
                Ws.Cells(3, 17).Value = minPercChange
                minPercChangeTicker = Ws.Cells(stockCount, 9).Value
                Ws.Cells(3, 16).Value = minPercChangeTicker
                        
            Else
            End If
                    
            'again similar logic to the above, it only differs from the highPercChange by looking at volume
            If stockVol > maxVol Then
                maxVol = stockVol
                Ws.Cells(4, 17).Value = maxVol
                maxVolTicker = Ws.Cells(stockCount, 9).Value
                Ws.Cells(4, 16).Value = maxVolTicker
            Else
            End If
        Next rowNum
        
    Next Ws
            
End Sub



