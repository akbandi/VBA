Attribute VB_Name = "Module1"
Sub Multiyearstock()
    
    Dim column      As Integer
    Dim Index       As Integer
    Dim ticker      As String
    Dim volume  As LongLong
    Dim GreatestIncrease         As Double
    Dim GreatestDecrease     As Double
    Dim TotalVolume As LongLong
    
    column = 1
    GreatestIncrease = 0
    GreatestDecrease = 0
    TotalVolume = 0
    
    For Each ws In Worksheets
    
        Index = 2
        ws.Activate
        
        Dim worksheetName As String
        worksheetName = ws.Name
        
        MsgBox (ws.Name)
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Open Price"
        ws.Cells(1, 11).Value = "Close Price"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = "Greatest % Decrease "
        ws.Cells(4, 17).Value = "Greatest Total Volume"
        ws.Cells(1, 18).Value = "Ticker"
        ws.Cells(1, 19).Value = "Value"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            
            If ws.Cells(i - 1, column).Value <> ws.Cells(i, column).Value Then
                
                'Open price
                openprice = ws.Cells(i, 3).Value
                
            End If
            
            ticker = ws.Cells(i, 1).Value
            
            ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                
                'populating Ticker to exel
                ws.Cells(Index, 9).Value = ticker
                
                'populating openprice to excel
                ws.Cells(Index, 10).Value = openprice
                
                'populating close price to excel
                closeprice = ws.Cells(i, 6).Value
                ws.Cells(Index, 11).Value = closeprice
                
                'Yearly change calculation
                yearlychange = closeprice - openprice
                ws.Cells(Index, 12).Value = yearlychange
                
                'percent change calculation
                percentchange = (closeprice - openprice) / openprice
                ws.Cells(Index, 13).Value = percentchange
                           
                'Total volume of each stock
                volume = volume + Cells(i, 7).Value
                ws.Cells(Index, 14).Value = volume
        
                'Bonus - Greatest % increase and decrease
                If ws.Cells(Index, 13).Value > GreatestIncrease Then
                    GreatestIncrease = ws.Cells(Index, 13).Value
                    Cells(2, 18).Value = ticker
                    Cells(2, 19).Value = GreatestIncrease
                    
                End If
                
                If ws.Cells(Index, 13).Value < GreatestDecrease Then
                    GreatestDecrease = ws.Cells(Index, 13).Value
                    Cells(3, 18).Value = ticker
                    Cells(3, 19).Value = GreatestDecrease
                    
                End If
                
                If ws.Cells(Index, 14).Value > TotalVolume Then
                    TotalVolume = ws.Cells(Index, 14).Value
                    Cells(4, 18).Value = ticker
                    Cells(4, 19).Value = TotalVolume
                End If

                'Formating cells
                ws.Cells(Index, 13).NumberFormat = "0.00%"
                Cells(2, 19).NumberFormat = "0.00%"
                Cells(3, 19).NumberFormat = "0.00%"
                
                If yearlychange >= 0 Then
                    
                    ws.Cells(Index, 12).Interior.ColorIndex = 4
                    
                Else
                    
                    ws.Cells(Index, 12).Interior.ColorIndex = 3
                    
                End If
                
                volume = 0
                Index = Index + 1
                
            Else
                
                'else add up next volume if the cell is not different
               volume = volume + Cells(i, 7).Value

                
            End If
            
        Next i
        
    Next ws
    
End Sub
