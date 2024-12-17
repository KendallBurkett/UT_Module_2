Sub QuarterlyStockDataSummary()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim tickers As Collection
    Dim summaryRow As Long
    Dim summaryCol As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Name Like "Q*" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            
            summaryCol = lastCol + 2
            ws.Cells(1, summaryCol).Value = "Ticker"
            ws.Cells(1, summaryCol + 1).Value = "Quarterly Change"
            ws.Cells(1, summaryCol + 2).Value = "Percentage Change"
            ws.Cells(1, summaryCol + 3).Value = "Total Volume"
            
            Set tickers = New Collection
            
            Dim i As Long
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value
                On Error Resume Next
                tickers.Add ticker, CStr(ticker)
                On Error GoTo 0
            Next i
            
            
            maxIncrease = -999999
            maxDecrease = 999999
            maxVolume = 0
            
            summaryRow = 2
            
            
            Dim item As Variant
            For Each item In tickers
                ticker = item
                openingPrice = ws.Cells(Application.WorksheetFunction.Match(ticker, ws.Columns(1), 0), 3).Value
                closingPrice = ws.Cells(Application.WorksheetFunction.Match(ticker, ws.Columns(1), 0) + _
                                Application.WorksheetFunction.CountIf(ws.Columns(1), ticker) - 1, 6).Value
                totalVolume = Application.WorksheetFunction.SumIf(ws.Columns(1), ticker, ws.Columns(7))
                
                quarterlyChange = closingPrice - openingPrice
                percentageChange = Round((quarterlyChange / openingPrice) * 100, 2)
                
                
                ws.Cells(summaryRow, summaryCol).Value = ticker
                ws.Cells(summaryRow, summaryCol + 1).Value = quarterlyChange
                ws.Cells(summaryRow, summaryCol + 2).Value = Format(percentageChange, "0.00") & "%"
                ws.Cells(summaryRow, summaryCol + 3).Value = totalVolume
                
                
                With ws.Cells(summaryRow, summaryCol + 1).Interior
                    If quarterlyChange > 0 Then
                        .Color = RGB(144, 238, 144) ' Light green for positive change
                    ElseIf quarterlyChange < 0 Then
                        .Color = RGB(255, 182, 193) ' Light red for negative change
                    Else
                        .Pattern = xlNone ' No color for zero change
                    End If
                End With
                
                
                If percentageChange > maxIncrease Then
                    maxIncrease = percentageChange
                    maxIncreaseTicker = ticker
                End If
                If percentageChange < maxDecrease Then
                    maxDecrease = percentageChange
                    maxDecreaseTicker = ticker
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
                
                summaryRow = summaryRow + 1
            Next item
            
            
            ws.Cells(1, summaryCol + 5).Value = "Greatest % Increase"
            ws.Cells(1, summaryCol + 6).Value = "Ticker"
            ws.Cells(1, summaryCol + 7).Value = "Value"
            ws.Cells(2, summaryCol + 6).Value = maxIncreaseTicker
            ws.Cells(2, summaryCol + 7).Value = Format(maxIncrease, "0.00") & "%"
            
            ws.Cells(3, summaryCol + 5).Value = "Greatest % Decrease"
            ws.Cells(3, summaryCol + 6).Value = maxDecreaseTicker
            ws.Cells(3, summaryCol + 7).Value = Format(maxDecrease, "0.00") & "%"
            
            ws.Cells(4, summaryCol + 5).Value = "Greatest Total Volume"
            ws.Cells(4, summaryCol + 6).Value = maxVolumeTicker
            ws.Cells(4, summaryCol + 7).Value = maxVolume
        End If
    Next ws
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
