Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()
    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim rowNum As Long
    Dim outputRow As Long ' variable to keep track of the output row in the worksheet
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of data in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set initial values
        ticker = ""
        openingPrice = 0
        closingPrice = 0
        yearlyChange = 0
        percentChange = 0
        totalVolume = 0
        outputRow = 2 ' Start from the second row for output
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        maxPercentIncreaseTicker = ""
        maxPercentDecreaseTicker = ""
        maxTotalVolumeTicker = ""
        
        ' Set column headers for the analysis
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through all rows of data
        For rowNum = 2 To lastRow
            ' Check if it's a new ticker symbol
            If ws.Cells(rowNum, 1).Value <> ws.Cells(rowNum + 1, 1).Value Then
                ' Set the ticker symbol
                ticker = ws.Cells(rowNum, 1).Value
                ' Set the closing price for the ticker
                closingPrice = ws.Cells(rowNum, 6).Value
                
                ' Calculate the yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                
                ' Output the analysis results
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Apply conditional formatting to the "Yearly Change" column
                If yearlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red color for negative amounts
                ElseIf yearlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green color for positive amounts
                End If
                
                ' Check and update maximum values
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = ticker
                End If
                
                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                
                ' Move to the next output row
                outputRow = outputRow + 1
                
                ' Reset the values for the next ticker
                ticker = ""
                openingPrice = 0
                closingPrice = 0
                yearlyChange = 0
                percentChange = 0
                totalVolume = 0
            Else
                ' Check if it's the first day of the ticker
                If openingPrice = 0 Then
                    openingPrice = ws.Cells(rowNum, 3).Value
                End If
                
                ' Sum the total stock volume for the ticker
                totalVolume = totalVolume + ws.Cells(rowNum, 7).Value
            End If
        Next rowNum
        
        ' Apply formatting to the analysis results
        ws.Range("K:K").NumberFormat = "0.00%"
        
        ' Output the stock with the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume"
        ws.Cells(2, 16).Value = maxPercentIncreaseTicker
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(3, 16).Value = maxPercentDecreaseTicker
        ws.Cells(3, 17).Value = maxPercentDecrease
        ws.Cells(4, 16).Value = maxTotalVolumeTicker
        ws.Cells(4, 17).Value = maxTotalVolume
        
        ' Apply formatting to the stock analysis results
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,##0"
    Next ws
End Sub


