'==============================================================================
' Module_ScanExport — Deduplicated Export from ScanInput → DataLog
'==============================================================================
Option Explicit

' ── Main Export Macro (assign to button) ──
Public Sub ExportScanToLog()
    Dim wsScan As Worksheet, wsLog As Worksheet
    Set wsScan = ThisWorkbook.Sheets(SHT_SCAN)
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    Dim scanLastRow As Long
    scanLastRow = LastRow(wsScan, COL_SCAN_PRODUCT)
    
    If scanLastRow <= 1 Then
        MsgBox "No scan data found in ScanInput.", vbExclamation, "Export"
        Exit Sub
    End If
    
    Dim logNextRow As Long
    logNextRow = LastRow(wsLog, COL_LOG_KEY) + 1
    
    Dim added As Long, skipped As Long, noClose As Long
    added = 0: skipped = 0: noClose = 0
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i As Long
    For i = 2 To scanLastRow
        ' Read scan row
        Dim product As String
        Dim lots As Double, notional As Double, crossLevel As Double
        Dim condition As String
        Dim scanTS As Date, tradeTime As Date
        
        scanTS = wsScan.Cells(i, COL_SCAN_TIMESTAMP).Value
        product = Trim(CStr(wsScan.Cells(i, COL_SCAN_PRODUCT).Value))
        lots = Val(wsScan.Cells(i, COL_SCAN_LOTS).Value)
        notional = Val(wsScan.Cells(i, COL_SCAN_NOTIONAL).Value)
        condition = Trim(CStr(wsScan.Cells(i, COL_SCAN_CONDITION).Value))
        crossLevel = Val(wsScan.Cells(i, COL_SCAN_CROSSLEVEL).Value)
        tradeTime = wsScan.Cells(i, COL_SCAN_TRADETIME).Value
        
        If product = "" Or crossLevel = 0 Then GoTo NextRow
        
        ' Generate dedup key
        Dim dedupeKey As String
        dedupeKey = MakeDedupeKey(product, tradeTime, crossLevel)
        
        ' Check for duplicate
        If KeyExists(dedupeKey) Then
            skipped = skipped + 1
            GoTo NextRow
        End If
        
        ' Get previous close
        Dim prevClose As Double
        prevClose = GetPrevClose(product)
        
        ' Calculate premium
        Dim premium As Double, premRounded As Double
        Dim idbFlag As String
        
        If prevClose > 0 Then
            premium = CalcPremium(crossLevel, prevClose)
            premRounded = Round(premium * 100, 2) / 100
            
            If IsCleanPremium(premium) Then
                idbFlag = "LIKELY IDB"
            Else
                idbFlag = "LIKELY CLIENT"
            End If
        Else
            premium = 0: premRounded = 0
            idbFlag = "NO CLOSE DATA"
            noClose = noClose + 1
        End If
        
        ' Write to DataLog
        With wsLog
            .Cells(logNextRow, COL_LOG_KEY).Value = dedupeKey
            .Cells(logNextRow, COL_LOG_SCANTS).Value = scanTS
            .Cells(logNextRow, COL_LOG_PRODUCT).Value = product
            .Cells(logNextRow, COL_LOG_LOTS).Value = lots
            .Cells(logNextRow, COL_LOG_NOTIONAL).Value = notional
            .Cells(logNextRow, COL_LOG_CONDITION).Value = condition
            .Cells(logNextRow, COL_LOG_CROSSLEVEL).Value = crossLevel
            .Cells(logNextRow, COL_LOG_TRADETIME).Value = tradeTime
            .Cells(logNextRow, COL_LOG_PREVCLOSE).Value = prevClose
            .Cells(logNextRow, COL_LOG_PREMIUM).Value = premium
            .Cells(logNextRow, COL_LOG_PREMROUND).Value = premRounded
            .Cells(logNextRow, COL_LOG_IDBFLAG).Value = idbFlag
            .Cells(logNextRow, COL_LOG_OVERRIDE).Value = ""
            .Cells(logNextRow, COL_LOG_NOTES).Value = ""
            .Cells(logNextRow, COL_LOG_TRADEDATE).Value = Int(tradeTime)
            
            ' Conditional formatting for IDB flag
            If idbFlag = "LIKELY IDB" Then
                .Cells(logNextRow, COL_LOG_IDBFLAG).Interior.Color = RGB(198, 239, 206)  ' Green
                .Cells(logNextRow, COL_LOG_IDBFLAG).Font.Color = RGB(0, 97, 0)
            ElseIf idbFlag = "LIKELY CLIENT" Then
                .Cells(logNextRow, COL_LOG_IDBFLAG).Interior.Color = RGB(255, 199, 206)  ' Red
                .Cells(logNextRow, COL_LOG_IDBFLAG).Font.Color = RGB(156, 0, 6)
            Else
                .Cells(logNextRow, COL_LOG_IDBFLAG).Interior.Color = RGB(255, 235, 156)  ' Yellow
                .Cells(logNextRow, COL_LOG_IDBFLAG).Font.Color = RGB(156, 101, 0)
            End If
            
            ' Format premium as percentage
            .Cells(logNextRow, COL_LOG_PREMIUM).NumberFormat = "0.0000%"
            .Cells(logNextRow, COL_LOG_PREMROUND).NumberFormat = "0.00%"
            .Cells(logNextRow, COL_LOG_TRADETIME).NumberFormat = "yyyy-mm-dd hh:nn:ss"
            .Cells(logNextRow, COL_LOG_TRADEDATE).NumberFormat = "yyyy-mm-dd"
        End With
        
        logNextRow = logNextRow + 1
        added = added + 1
        
NextRow:
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Auto-fit columns
    wsLog.Columns("A:O").AutoFit
    
    ' Summary
    Dim msg As String
    msg = "Export Complete!" & vbCrLf & vbCrLf
    msg = msg & "Added: " & added & " trades" & vbCrLf
    msg = msg & "Skipped (dupes): " & skipped & vbCrLf
    If noClose > 0 Then
        msg = msg & "Missing close data: " & noClose & " (update PrevClose sheet)"
    End If
    
    MsgBox msg, vbInformation, "Scan Export"
    
    ' Refresh dashboard
    Call RefreshDashboard
End Sub

' ── Clear ScanInput after export (optional) ──
Public Sub ClearScanInput()
    Dim wsScan As Worksheet
    Set wsScan = ThisWorkbook.Sheets(SHT_SCAN)
    
    Dim lr As Long
    lr = LastRow(wsScan, COL_SCAN_PRODUCT)
    
    If lr > 1 Then
        wsScan.Rows("2:" & lr).Delete
        MsgBox "ScanInput cleared.", vbInformation
    End If
End Sub
