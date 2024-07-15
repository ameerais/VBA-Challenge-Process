Attribute VB_Name = "Module1"
Sub MultipleYearStockAnalysis()

    ' Declare variables
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
    Dim i As Long
    Dim startRow As Long ' Variable to store the start row for each ticker
    Dim sheetName As String ' Variable to store the sheet name
    
    ' Initialize the greatest values
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        sheetName = ws.Name
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        summaryRow = 2
        
        ' Add headers for the summary table (if not already added)
        If ws.Cells(1, 9).Value <> "Ticker" Then
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Total Volume"
            ws.Cells(1, 11).Value = "Quarterly Change"
            ws.Cells(1, 12).Value = "Percent Change"
        End If
        
        ' Initialize total volume to zero for the new sheet
        totalVolume = 0
        
        ' Set the start row for the first ticker
        startRow = 2
        openPrice = ws.Cells(startRow, 3).Value
        
        ' Loop through each row in the current sheet
        For i = 2 To lastRow
            ' Accumulate volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if the next row has a different ticker symbol or we reached the last row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
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
                
                ' Apply Conditional Formatting
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
                
                ' Move to the next summary row and reset values for the next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                If i <> lastRow Then
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
            End If
        Next i
    Next ws
    
    ' Output the greatest values
    With ThisWorkbook.Worksheets(1) ' You can change this to a specific summary sheet if needed
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(2, 16).Value = greatestIncreaseTicker
        .Cells(2, 17).Value = greatestIncrease
        .Cells(3, 15).Value = "Greatest % Decrease"
        .Cells(3, 16).Value = greatestDecreaseTicker
        .Cells(3, 17).Value = greatestDecrease
        .Cells(4, 15).Value = "Greatest Total Volume"
        .Cells(4, 16).Value = greatestVolumeTicker
        .Cells(4, 17).Value = greatestVolume
    End With

End Sub

