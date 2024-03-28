Sub StockAnalysis()
    
    ' List of variables used
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerSymbol As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTable As Range
    Dim summaryRow As Long
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    ' Looping through al lthe worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' The summary table range
        Set summaryTable = ws.Range("I2:L" & lastRow)
        
        ' Clear previous summary data and formatting for each run
        summaryTable.ClearContents
        summaryTable.FormatConditions.Delete
        
        ' Initialize summary table headers and % table
        ws.Cells(1, 9).value = "Ticker"
        ws.Cells(1, 10).value = "Yearly Change"
        ws.Cells(1, 11).value = "Percent Change"
        ws.Cells(1, 12).value = "Total Stock Volume"
        ws.Cells(1, 15).value = "Ticker"
        ws.Cells(1, 16).value = "Value"
        
        
        summaryRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Loop through each row of data and find total of each ticker
        For i = 2 To lastRow
            
            If ws.Cells(i, 1).value <> ws.Cells(i - 1, 1).value Then
                If i <> 2 Then
                    yearlyChange = closingPrice - openingPrice
                    If openingPrice <> 0 Then
                        percentChange = (closingPrice - openingPrice) / openingPrice
                    Else
                        percentChange = 0
                    End If
                    
                    ' The summary table
                    ws.Cells(summaryRow, 9).value = tickerSymbol
                    ws.Cells(summaryRow, 10).value = yearlyChange
                    ws.Cells(summaryRow, 11).value = percentChange
                    ws.Cells(summaryRow, 12).value = totalVolume
                    
                    ' Apply conditional formatting to yearly change column and Percent change
                    ApplyConditionalFormatting ws.Cells(summaryRow, 10), yearlyChange
                    ApplyConditionalFormatting ws.Cells(summaryRow, 11), percentChange
                    
                    
                    ' The greatest increase, decrease, and volume
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = tickerSymbol
                        ws.Cells(2, 15).value = greatestIncreaseTicker
                        ws.Cells(2, 16).value = greatestIncrease

                    ElseIf percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = tickerSymbol
                        ws.Cells(3, 15).value = greatestDecreaseTicker
                        ws.Cells(3, 16).value = greatestDecrease
                        
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = tickerSymbol
                        ws.Cells(4, 15).value = greatestVolumeTicker
                        ws.Cells(4, 16).value = greatestVolume
                    End If
                    
                    summaryRow = summaryRow + 1
                End If
                
                ' New variables for the new ticker symbol
                tickerSymbol = ws.Cells(i, 1).value
                openingPrice = ws.Cells(i, 3).value
                totalVolume = 0
            End If
            
            ' Total volume and closing price
            totalVolume = totalVolume + ws.Cells(i, 7).value
            closingPrice = ws.Cells(i, 6).value
            
            ' If it's the last row
            If i = lastRow Then
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = (closingPrice - openingPrice) / openingPrice
                Else
                    percentChange = 0
                End If
                
                ' Data to summary table
                ws.Cells(summaryRow, 9).value = tickerSymbol
                ws.Cells(summaryRow, 10).value = yearlyChange
                ws.Cells(summaryRow, 11).value = percentChange
                ws.Cells(summaryRow, 12).value = totalVolume
                
            End If
        Next i
        
        ' Output greatest increase, decrease, and volume
        ws.Cells(2, 14).value = "Greatest % Increase"
        ws.Cells(3, 14).value = "Greatest % Decrease"
        ws.Cells(4, 14).value = "Greatest Total Volume"
        ws.Cells(2, 15).value = greatestIncreaseTicker
        ws.Cells(3, 15).value = greatestDecreaseTicker
        ws.Cells(4, 15).value = greatestVolumeTicker
    Next ws
End Sub

Sub ApplyConditionalFormatting(cell As Range, value As Double)
    Dim rng As Range
    
    ' Define the range
    Set rng = cell
    
    ' Clear existing conditional formatting
    rng.FormatConditions.Delete
    
    ' Apply conditional formatting for positive change (green) and negative change (red)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 200, 0) ' Green
    End With
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(200, 0, 0) ' Red
    End With
End Sub


