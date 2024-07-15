Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim summaryRow As Integer
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim greatestIncreaseTicker As String, greatestDecreaseTicker As String, greatestVolumeTicker As String

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        summaryRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Add headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                closePrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Update summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = totalVolume
                ws.Cells(summaryRow, 11).Value = quarterlyChange
                ws.Cells(summaryRow, 12).Value = percentChange
                
                ' Conditional Formatting
                If quarterlyChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(summaryRow, 11).Interior.Color = vbRed
                End If
                If percentChange > 0 Then
                    ws.Cells(summaryRow, 12).Interior.Color = vbGreen
                Else
                    ws.Cells(summaryRow, 12).Interior.Color = vbRed
                End If
                
                ' Find greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                summaryRow = summaryRow + 1
                totalVolume = 0 ' Reset total volume for next ticker
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Display greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
    Next ws

End Sub

