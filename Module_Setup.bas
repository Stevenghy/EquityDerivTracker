'==============================================================================
' Module_Setup — Initialize Workbook Structure (run once)
'==============================================================================
Option Explicit

' ── Run this once to set up all sheets ──
Public Sub InitializeWorkbook()
    Application.ScreenUpdating = False
    
    ' ── ScanInput Sheet ──
    Dim wsScan As Worksheet
    Set wsScan = EnsureSheet(SHT_SCAN)
    Call FormatAsTable(wsScan, Array("Scan Timestamp", "Product", "Lots", "Notional", _
                                     "Condition", "Cross Level", "Trade Time"))
    wsScan.Columns("A:A").ColumnWidth = 20
    wsScan.Columns("B:B").ColumnWidth = 18
    wsScan.Columns("F:F").NumberFormat = "#,##0.0000"
    wsScan.Columns("G:G").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    ' ── PrevClose Sheet ──
    Dim wsPC As Worksheet
    Set wsPC = EnsureSheet(SHT_CLOSE)
    Call FormatAsTable(wsPC, Array("Product", "Prev Close Level", "Close Date"))
    wsPC.Columns("B:B").NumberFormat = "#,##0.0000"
    wsPC.Columns("C:C").NumberFormat = "yyyy-mm-dd"
    
    ' Add sample products (customize these)
    Dim sampleProducts As Variant
    sampleProducts = Array("HSI", "HSCEI", "NKY", "SPX", "SX5E", "KOSPI2", "TWSE", "AS51")
    
    Dim i As Long
    For i = 0 To UBound(sampleProducts)
        wsPC.Cells(i + 2, COL_PC_PRODUCT).Value = sampleProducts(i)
        wsPC.Cells(i + 2, COL_PC_CLOSE).Value = 0  ' Fill manually or via Bloomberg
        wsPC.Cells(i + 2, COL_PC_DATE).Value = Date - 1
    Next i
    
    ' ── DataLog Sheet ──
    Dim wsLog As Worksheet
    Set wsLog = EnsureSheet(SHT_LOG)
    Call FormatAsTable(wsLog, Array("DedupeKey", "Scan Timestamp", "Product", "Lots", _
                                    "Notional", "Condition", "Cross Level", "Trade Time", _
                                    "Prev Close", "Premium %", "Premium Rounded", _
                                    "IDB Flag", "Manual Override", "Notes", "Trade Date"))
    wsLog.Columns("A:A").ColumnWidth = 45
    wsLog.Columns("G:G").NumberFormat = "#,##0.0000"
    wsLog.Columns("J:J").NumberFormat = "0.0000%"
    wsLog.Columns("K:K").NumberFormat = "0.00%"
    wsLog.Columns("H:H").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsLog.Columns("O:O").NumberFormat = "yyyy-mm-dd"
    
    ' Hide dedup key column (still used internally)
    wsLog.Columns("A:A").Hidden = True
    
    ' ── Dashboard Sheet ──
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHT_DASH)
    wsDash.Cells.Clear
    
    ' ── Add control buttons to ScanInput ──
    Call AddControlButtons
    
    Application.ScreenUpdating = True
    
    MsgBox "Workbook initialized!" & vbCrLf & vbCrLf & _
           "Sheets created:" & vbCrLf & _
           "• ScanInput — Bloomberg scan landing zone" & vbCrLf & _
           "• PrevClose — Previous day close levels" & vbCrLf & _
           "• DataLog — Master trade log (auto-deduped)" & vbCrLf & _
           "• Dashboard — Intraday charts" & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Fill in PrevClose levels for your products" & vbCrLf & _
           "2. Configure Bloomberg scan formulas in ScanInput" & vbCrLf & _
           "3. Click 'Export to Log' after each scan", _
           vbInformation, "Setup Complete"
End Sub

' ── Add macro buttons to ScanInput sheet ──
Private Sub AddControlButtons()
    Dim wsScan As Worksheet
    Set wsScan = ThisWorkbook.Sheets(SHT_SCAN)
    
    ' Remove existing buttons
    Dim shp As Shape
    For Each shp In wsScan.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp
    
    ' Export button
    Dim btn As Shape
    Set btn = wsScan.Shapes.AddFormControl(xlButtonControl, _
              wsScan.Cells(1, 9).Left, wsScan.Cells(1, 9).Top, 120, 30)
    btn.OnAction = "ExportScanToLog"
    btn.TextFrame.Characters.Text = "Export to Log"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True
    
    ' Clear scan button
    Set btn = wsScan.Shapes.AddFormControl(xlButtonControl, _
              wsScan.Cells(1, 9).Left + 130, wsScan.Cells(1, 9).Top, 120, 30)
    btn.OnAction = "ClearScanInput"
    btn.TextFrame.Characters.Text = "Clear Scan"
    btn.TextFrame.Characters.Font.Size = 10
    
    ' Auto scan start
    Set btn = wsScan.Shapes.AddFormControl(xlButtonControl, _
              wsScan.Cells(1, 9).Left + 260, wsScan.Cells(1, 9).Top, 130, 30)
    btn.OnAction = "StartAutoScan"
    btn.TextFrame.Characters.Text = "Start Auto Scan"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True
    
    ' Auto scan stop
    Set btn = wsScan.Shapes.AddFormControl(xlButtonControl, _
              wsScan.Cells(1, 9).Left + 400, wsScan.Cells(1, 9).Top, 130, 30)
    btn.OnAction = "StopAutoScan"
    btn.TextFrame.Characters.Text = "Stop Auto Scan"
    btn.TextFrame.Characters.Font.Size = 10
    
    ' Refresh dashboard
    Set btn = wsScan.Shapes.AddFormControl(xlButtonControl, _
              wsScan.Cells(1, 9).Left, wsScan.Cells(2, 9).Top + 5, 120, 30)
    btn.OnAction = "RefreshDashboard"
    btn.TextFrame.Characters.Text = "Refresh Charts"
    btn.TextFrame.Characters.Font.Size = 10
    
    ' Update prev close
    Set btn = wsScan.Shapes.AddFormControl(xlButtonControl, _
              wsScan.Cells(1, 9).Left + 130, wsScan.Cells(2, 9).Top + 5, 130, 30)
    btn.OnAction = "UpdatePrevClose"
    btn.TextFrame.Characters.Text = "Update Prev Close"
    btn.TextFrame.Characters.Font.Size = 10
    
    ' Export CSV
    Set btn = wsScan.Shapes.AddFormControl(xlButtonControl, _
              wsScan.Cells(1, 9).Left + 270, wsScan.Cells(2, 9).Top + 5, 120, 30)
    btn.OnAction = "ExportToCSV"
    btn.TextFrame.Characters.Text = "Export CSV"
    btn.TextFrame.Characters.Font.Size = 10
End Sub
