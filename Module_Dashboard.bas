'==============================================================================
' Module_Dashboard — Intraday Charts & Visualization
'==============================================================================
Option Explicit

' ── Main Dashboard Refresh ──
Public Sub RefreshDashboard()
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHT_DASH)
    
    Application.ScreenUpdating = False
    
    ' Clear existing charts
    Dim cht As ChartObject
    For Each cht In wsDash.ChartObjects
        cht.Delete
    Next cht
    
    ' Clear existing data
    wsDash.Cells.Clear
    
    ' Build summary header
    wsDash.Range("A1").Value = "EQUITY DERIVATIVES CROSS TRACKER"
    wsDash.Range("A1").Font.Size = 16
    wsDash.Range("A1").Font.Bold = True
    wsDash.Range("A2").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsDash.Range("A2").Font.Color = RGB(128, 128, 128)
    
    ' Get today's date
    Dim today As Date
    today = Int(Now)
    
    wsDash.Range("A3").Value = "Showing trades for: " & Format(today, "yyyy-mm-dd")
    wsDash.Range("A3").Font.Bold = True
    
    ' Build intraday premium chart
    Call BuildIntradayPremiumChart(wsDash, today)
    
    ' Build trade count summary
    Call BuildTradeSummary(wsDash, today)
    
    ' Build product breakdown
    Call BuildProductBreakdown(wsDash, today)
    
    Application.ScreenUpdating = True
End Sub

' ── Intraday Premium % Chart (main visualization) ──
Private Sub BuildIntradayPremiumChart(wsDash As Worksheet, tradeDate As Date)
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    Dim lr As Long
    lr = LastRow(wsLog, COL_LOG_KEY)
    If lr <= 1 Then Exit Sub
    
    ' Collect unique products for today
    Dim products As Object
    Set products = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lr
        If Int(wsLog.Cells(i, COL_LOG_TRADEDATE).Value) = tradeDate Then
            Dim prod As String
            prod = wsLog.Cells(i, COL_LOG_PRODUCT).Value
            If Not products.Exists(prod) Then
                products.Add prod, prod
            End If
        End If
    Next i
    
    If products.Count = 0 Then
        wsDash.Range("A5").Value = "No trades logged for today."
        Exit Sub
    End If
    
    ' Build data table for chart starting at row 5
    Dim dataStartRow As Long
    dataStartRow = 5
    
    wsDash.Cells(dataStartRow, 1).Value = "Trade Time"
    wsDash.Cells(dataStartRow, 2).Value = "Product"
    wsDash.Cells(dataStartRow, 3).Value = "Premium %"
    wsDash.Cells(dataStartRow, 4).Value = "IDB Flag"
    wsDash.Cells(dataStartRow, 5).Value = "Lots"
    
    Dim row As Long
    row = dataStartRow + 1
    
    For i = 2 To lr
        If Int(wsLog.Cells(i, COL_LOG_TRADEDATE).Value) = tradeDate Then
            wsDash.Cells(row, 1).Value = wsLog.Cells(i, COL_LOG_TRADETIME).Value
            wsDash.Cells(row, 1).NumberFormat = "hh:mm:ss"
            wsDash.Cells(row, 2).Value = wsLog.Cells(i, COL_LOG_PRODUCT).Value
            wsDash.Cells(row, 3).Value = wsLog.Cells(i, COL_LOG_PREMIUM).Value
            wsDash.Cells(row, 3).NumberFormat = "0.00%"
            wsDash.Cells(row, 4).Value = wsLog.Cells(i, COL_LOG_IDBFLAG).Value
            wsDash.Cells(row, 5).Value = wsLog.Cells(i, COL_LOG_LOTS).Value
            
            ' Color IDB rows
            If wsLog.Cells(i, COL_LOG_IDBFLAG).Value = "LIKELY IDB" Then
                wsDash.Range(wsDash.Cells(row, 1), wsDash.Cells(row, 5)).Interior.Color = RGB(234, 247, 237)
            End If
            
            row = row + 1
        End If
    Next i
    
    Dim dataEndRow As Long
    dataEndRow = row - 1
    
    If dataEndRow < dataStartRow + 1 Then Exit Sub
    
    ' ── Create scatter chart: Premium % over time ──
    Dim chartLeft As Double
    chartLeft = wsDash.Cells(dataStartRow, 7).Left
    
    Dim chtObj As ChartObject
    Set chtObj = wsDash.ChartObjects.Add(Left:=chartLeft, Top:=wsDash.Cells(dataStartRow, 1).Top, _
                                          Width:=550, Height:=350)
    
    With chtObj.Chart
        .ChartType = xlXYScatterLines
        
        ' Add a series per product
        Dim pKeys As Variant
        pKeys = products.Keys
        Dim p As Long
        
        For p = 0 To UBound(pKeys)
            ' Collect data points for this product
            Dim xVals() As Date, yVals() As Double
            Dim count As Long
            count = 0
            
            ' First pass: count
            Dim j As Long
            For j = dataStartRow + 1 To dataEndRow
                If wsDash.Cells(j, 2).Value = pKeys(p) Then count = count + 1
            Next j
            
            If count > 0 Then
                ReDim xVals(1 To count)
                ReDim yVals(1 To count)
                
                Dim idx As Long
                idx = 1
                For j = dataStartRow + 1 To dataEndRow
                    If wsDash.Cells(j, 2).Value = pKeys(p) Then
                        xVals(idx) = wsDash.Cells(j, 1).Value
                        yVals(idx) = wsDash.Cells(j, 3).Value * 100  ' Display as 100.xx
                        idx = idx + 1
                    End If
                Next j
                
                Dim ser As Series
                Set ser = .SeriesCollection.NewSeries
                ser.Name = pKeys(p)
                ser.XValues = xVals
                ser.Values = yVals
                ser.MarkerStyle = xlMarkerStyleCircle
                ser.MarkerSize = 6
            End If
        Next p
        
        .HasTitle = True
        .ChartTitle.Text = "Intraday Premium % — " & Format(tradeDate, "yyyy-mm-dd")
        
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Time"
        .Axes(xlCategory).TickLabels.NumberFormat = "hh:mm"
        
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Premium (100.xx%)"
        .Axes(xlValue).TickLabels.NumberFormat = "0.00"
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        .PlotArea.Interior.Color = RGB(245, 245, 250)
    End With
End Sub

' ── Trade Count Summary Table ──
Private Sub BuildTradeSummary(wsDash As Worksheet, tradeDate As Date)
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    Dim lr As Long
    lr = LastRow(wsLog, COL_LOG_KEY)
    
    Dim summaryRow As Long
    summaryRow = 22
    
    wsDash.Cells(summaryRow, 1).Value = "TODAY'S SUMMARY"
    wsDash.Cells(summaryRow, 1).Font.Bold = True
    wsDash.Cells(summaryRow, 1).Font.Size = 13
    
    summaryRow = summaryRow + 1
    
    Dim totalTrades As Long, idbTrades As Long, clientTrades As Long, noDataTrades As Long
    totalTrades = 0: idbTrades = 0: clientTrades = 0: noDataTrades = 0

    Dim totalNotional As Double
    totalNotional = 0

    Dim i As Long
    For i = 2 To lr
        If Int(wsLog.Cells(i, COL_LOG_TRADEDATE).Value) = tradeDate Then
            totalTrades = totalTrades + 1
            totalNotional = totalNotional + wsLog.Cells(i, COL_LOG_NOTIONAL).Value

            Select Case wsLog.Cells(i, COL_LOG_IDBFLAG).Value
                Case "LIKELY IDB": idbTrades = idbTrades + 1
                Case "LIKELY CLIENT": clientTrades = clientTrades + 1
                Case Else: noDataTrades = noDataTrades + 1
            End Select
        End If
    Next i

    wsDash.Cells(summaryRow, 1).Value = "Total Crosses:"
    wsDash.Cells(summaryRow, 2).Value = totalTrades
    summaryRow = summaryRow + 1

    wsDash.Cells(summaryRow, 1).Value = "Likely IDB:"
    wsDash.Cells(summaryRow, 2).Value = idbTrades
    wsDash.Cells(summaryRow, 2).Interior.Color = RGB(198, 239, 206)
    summaryRow = summaryRow + 1

    wsDash.Cells(summaryRow, 1).Value = "Likely Client:"
    wsDash.Cells(summaryRow, 2).Value = clientTrades
    wsDash.Cells(summaryRow, 2).Interior.Color = RGB(255, 199, 206)
    summaryRow = summaryRow + 1

    wsDash.Cells(summaryRow, 1).Value = "No Close Data:"
    wsDash.Cells(summaryRow, 2).Value = noDataTrades
    wsDash.Cells(summaryRow, 2).Interior.Color = RGB(255, 235, 156)
    summaryRow = summaryRow + 1

    wsDash.Cells(summaryRow, 1).Value = "Total Notional:"
    wsDash.Cells(summaryRow, 2).Value = totalNotional
    wsDash.Cells(summaryRow, 2).NumberFormat = "#,##0"
End Sub

' ── Product Breakdown Pie Chart ──
Private Sub BuildProductBreakdown(wsDash As Worksheet, tradeDate As Date)
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    Dim lr As Long
    lr = LastRow(wsLog, COL_LOG_KEY)
    
    ' Count trades per product
    Dim prodCount As Object
    Set prodCount = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lr
        If Int(wsLog.Cells(i, COL_LOG_TRADEDATE).Value) = tradeDate Then
            Dim prod As String
            prod = wsLog.Cells(i, COL_LOG_PRODUCT).Value
            If prodCount.Exists(prod) Then
                prodCount(prod) = prodCount(prod) + 1
            Else
                prodCount.Add prod, 1
            End If
        End If
    Next i
    
    If prodCount.Count = 0 Then Exit Sub
    
    ' Write data for pie chart
    Dim pieRow As Long
    pieRow = 29
    
    wsDash.Cells(pieRow, 1).Value = "Product"
    wsDash.Cells(pieRow, 2).Value = "Trades"
    wsDash.Cells(pieRow, 1).Font.Bold = True
    wsDash.Cells(pieRow, 2).Font.Bold = True
    
    Dim keys As Variant
    keys = prodCount.Keys
    Dim p As Long
    For p = 0 To UBound(keys)
        wsDash.Cells(pieRow + 1 + p, 1).Value = keys(p)
        wsDash.Cells(pieRow + 1 + p, 2).Value = prodCount(keys(p))
    Next p
    
    ' Create pie chart
    Dim pieEndRow As Long
    pieEndRow = pieRow + prodCount.Count
    
    Dim chtObj As ChartObject
    Set chtObj = wsDash.ChartObjects.Add( _
        Left:=wsDash.Cells(pieRow, 4).Left, _
        Top:=wsDash.Cells(22, 4).Top, _
        Width:=350, Height:=250)
    
    With chtObj.Chart
        .ChartType = xlPie
        .SetSourceData wsDash.Range(wsDash.Cells(pieRow, 1), wsDash.Cells(pieEndRow, 2))
        .HasTitle = True
        .ChartTitle.Text = "Trades by Product"
        .ApplyDataLabels xlDataLabelsShowPercent
    End With
End Sub
