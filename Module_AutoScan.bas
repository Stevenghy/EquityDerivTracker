'==============================================================================
' Module_AutoScan — Unattended Bloomberg Scanning & Scheduling
'==============================================================================
Option Explicit

Private nextRunTime As Date

' ── Start auto-scanning (call once to begin cycle) ──
Public Sub StartAutoScan()
    ' Verify we're in market hours
    If Not IsMarketHours() Then
        MsgBox "Outside market hours (" & MARKET_OPEN_HOUR & ":00 - " & MARKET_CLOSE_HOUR & ":00). " & _
               "Auto-scan will start at next market open.", vbInformation, "Auto Scan"
        ' Schedule for next market open
        Dim nextOpen As Date
        If Hour(Now) >= MARKET_CLOSE_HOUR Then
            nextOpen = Int(Now) + 1 + TimeSerial(MARKET_OPEN_HOUR, 0, 0)
        Else
            nextOpen = Int(Now) + TimeSerial(MARKET_OPEN_HOUR, 0, 0)
        End If
        nextRunTime = nextOpen
        Application.OnTime nextOpen, "StartAutoScan"
        Exit Sub
    End If
    
    ' Run scan
    Call RunBloombergScan
    
    ' Export to log
    Call ExportScanToLog
    
    ' Schedule next run
    nextRunTime = Now + TimeSerial(0, AUTO_SCAN_INTERVAL_MIN, 0)
    
    ' Only schedule if next run is still in market hours
    If Hour(nextRunTime) < MARKET_CLOSE_HOUR Then
        Application.OnTime nextRunTime, "StartAutoScan"
        Debug.Print "Next scan scheduled: " & Format(nextRunTime, "hh:nn:ss")
    Else
        ' Schedule end-of-day summary
        Call EndOfDaySummary
        
        ' Schedule for next market open
        Dim tmrwOpen As Date
        tmrwOpen = Int(Now) + 1 + TimeSerial(MARKET_OPEN_HOUR, 0, 0)
        Application.OnTime tmrwOpen, "StartAutoScan"
        Debug.Print "Market closed. Next scan: " & Format(tmrwOpen, "yyyy-mm-dd hh:nn")
    End If
End Sub

' ── Stop auto-scanning ──
Public Sub StopAutoScan()
    On Error Resume Next
    Application.OnTime nextRunTime, "StartAutoScan", , False
    On Error GoTo 0
    MsgBox "Auto-scan stopped.", vbInformation, "Auto Scan"
End Sub

' ── Check if we're in market hours ──
Private Function IsMarketHours() As Boolean
    Dim h As Long
    h = Hour(Now)
    IsMarketHours = (h >= MARKET_OPEN_HOUR And h < MARKET_CLOSE_HOUR)
    ' Skip weekends
    If Weekday(Now, vbMonday) > 5 Then IsMarketHours = False
End Function

' ── Run Bloomberg Scan (populate ScanInput) ──
Public Sub RunBloombergScan()
    Dim wsScan As Worksheet
    Set wsScan = ThisWorkbook.Sheets(SHT_SCAN)
    
    ' ──────────────────────────────────────────────────────
    ' BLOOMBERG API INTEGRATION
    ' ──────────────────────────────────────────────────────
    ' Option 1: If using Bloomberg COM API (BLP API)
    ' Uncomment and adapt the section below:
    '
    ' Dim blp As Object
    ' Set blp = CreateObject("Bloomberg.Data.1")  ' or Bloomberg.API
    '
    ' You would typically use:
    ' - BDP (Bloomberg Data Point) for single values
    ' - BDH (Bloomberg Data History) for time series
    ' - BSRCH (Bloomberg Search) for scanning crosses
    '
    ' Example BSRCH approach:
    ' Results = blp.BSearch("COMDTY:CROSS", "fields=PRODUCT,LOTS,NOTIONAL,...")
    '
    ' Option 2: If using Bloomberg Excel Add-in formulas
    ' The scan sheet would have BDS/BSRCH formulas that auto-refresh
    ' In that case, just trigger a calculate:
    '
    wsScan.Calculate
    Application.Wait Now + TimeSerial(0, 0, 5)  ' Wait for BB data
    wsScan.Calculate  ' Recalc to ensure all data loaded
    '
    ' ──────────────────────────────────────────────────────
    ' IMPORTANT: Adapt the above to your specific Bloomberg
    ' terminal setup. The scan query and field mapping depends
    ' on your Bloomberg subscription and the specific index
    ' futures you're tracking.
    '
    ' The key fields to populate in ScanInput are:
    ' Col A: Timestamp (when scan ran)
    ' Col B: Product name (e.g. "HSI", "NKY", "SPX")
    ' Col C: Number of lots
    ' Col D: Notional size
    ' Col E: Condition (e.g. "CROSS", "BLOCK")
    ' Col F: Crossing level (absolute price)
    ' Col G: Trade time
    ' ──────────────────────────────────────────────────────
    
    Debug.Print "Bloomberg scan completed at " & Format(Now, "hh:nn:ss")
End Sub

' ── Update Previous Close levels from Bloomberg ──
Public Sub UpdatePrevClose()
    Dim wsPC As Worksheet
    Set wsPC = ThisWorkbook.Sheets(SHT_CLOSE)
    
    Dim lr As Long
    lr = LastRow(wsPC, COL_PC_PRODUCT)
    
    ' ──────────────────────────────────────────────────────
    ' BLOOMBERG CLOSE PRICE UPDATE
    ' ──────────────────────────────────────────────────────
    ' For each product in PrevClose sheet, pull yesterday's
    ' closing level via BDP:
    '
    ' Dim i As Long
    ' For i = 2 To lr
    '     Dim ticker As String
    '     ticker = wsPC.Cells(i, COL_PC_PRODUCT).Value
    '     ' wsPC.Cells(i, COL_PC_CLOSE).Formula = "=BDP(""" & ticker & " Index"",""PX_LAST"")"
    '     ' wsPC.Cells(i, COL_PC_DATE).Formula = "=BDP(""" & ticker & " Index"",""PREV_CLOSE_DT"")"
    ' Next i
    '
    ' wsPC.Calculate
    ' ──────────────────────────────────────────────────────
    
    Debug.Print "Previous close levels updated at " & Format(Now, "hh:nn:ss")
    MsgBox "Previous close levels updated.", vbInformation, "PrevClose Update"
End Sub

' ── End of Day Summary (auto-runs after market close) ──
Private Sub EndOfDaySummary()
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    Dim today As Date
    today = Int(Now)
    
    Dim lr As Long
    lr = LastRow(wsLog, COL_LOG_KEY)
    
    Dim totalTrades As Long, idbCount As Long
    totalTrades = 0: idbCount = 0
    
    Dim i As Long
    For i = 2 To lr
        If Int(wsLog.Cells(i, COL_LOG_TRADEDATE).Value) = today Then
            totalTrades = totalTrades + 1
            If wsLog.Cells(i, COL_LOG_IDBFLAG).Value = "LIKELY IDB" Then
                idbCount = idbCount + 1
            End If
        End If
    Next i
    
    ' Export daily CSV
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\DailyExport_" & Format(today, "yyyymmdd") & ".csv"
    Call ExportToCSV(csvPath)
    
    ' Refresh dashboard
    Call RefreshDashboard
    
    ' Save workbook
    ThisWorkbook.Save
    
    Debug.Print "End of day summary: " & totalTrades & " trades (" & idbCount & " IDB)"
End Sub
