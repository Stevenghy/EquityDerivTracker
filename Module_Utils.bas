'==============================================================================
' Module_Utils — Helper Functions
'==============================================================================
Option Explicit

' ── Get last used row in a sheet column ──
Public Function LastRow(ws As Worksheet, Optional col As Long = 1) As Long
    If Application.WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
        LastRow = 1
    Else
        LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    End If
End Function

' ── Generate dedup key from trade data ──
Public Function MakeDedupeKey(product As String, tradeTime As Date, crossLevel As Double) As String
    MakeDedupeKey = UCase(Trim(product)) & "|" & Format(tradeTime, "yyyy-mm-dd hh:nn:ss") & "|" & Format(crossLevel, "0.0000")
End Function

' ── Check if dedup key exists in DataLog ──
Public Function KeyExists(key As String) As Boolean
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    Dim lr As Long
    lr = LastRow(wsLog, COL_LOG_KEY)
    
    If lr <= 1 Then
        KeyExists = False
        Exit Function
    End If
    
    Dim rng As Range
    Set rng = wsLog.Range(wsLog.Cells(2, COL_LOG_KEY), wsLog.Cells(lr, COL_LOG_KEY))
    
    Dim found As Range
    Set found = rng.Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    
    KeyExists = Not found Is Nothing
End Function

' ── Get previous close for a product ──
Public Function GetPrevClose(product As String) As Double
    Dim wsPC As Worksheet
    Set wsPC = ThisWorkbook.Sheets(SHT_CLOSE)
    
    Dim lr As Long
    lr = LastRow(wsPC, COL_PC_PRODUCT)
    
    Dim i As Long
    For i = 2 To lr
        If UCase(Trim(wsPC.Cells(i, COL_PC_PRODUCT).Value)) = UCase(Trim(product)) Then
            GetPrevClose = wsPC.Cells(i, COL_PC_CLOSE).Value
            Exit Function
        End If
    Next i
    
    GetPrevClose = 0  ' Not found
End Function

' ── Calculate premium % ──
Public Function CalcPremium(crossLevel As Double, prevClose As Double) As Double
    If prevClose = 0 Then
        CalcPremium = 0
    Else
        CalcPremium = (crossLevel / prevClose)
    End If
End Function

' ── Check if premium is "clean" (IDB-like, max 2dp) ──
Public Function IsCleanPremium(premium As Double) As Boolean
    ' Premium is expressed as e.g. 1.0031 (meaning 100.31%)
    ' Convert to display format: 100.31
    Dim displayPct As Double
    displayPct = premium * 100  ' e.g. 100.31
    
    ' Round to 2dp and compare
    Dim rounded As Double
    rounded = Round(displayPct, 2)
    
    ' Check if difference is negligible (floating point tolerance)
    IsCleanPremium = (Abs(displayPct - rounded) < 0.001)
End Function

' ── Count decimal places in a number ──
Public Function CountDP(val As Double) As Long
    Dim s As String
    s = CStr(val * 100)  ' Convert to percentage display
    
    Dim dotPos As Long
    dotPos = InStr(s, ".")
    
    If dotPos = 0 Then
        CountDP = 0
    Else
        ' Trim trailing zeros
        Dim trimmed As String
        trimmed = s
        Do While Right(trimmed, 1) = "0"
            trimmed = Left(trimmed, Len(trimmed) - 1)
        Loop
        If Right(trimmed, 1) = "." Then
            CountDP = 0
        Else
            CountDP = Len(trimmed) - dotPos
        End If
    End If
End Function

' ── Ensure sheet exists, create if not ──
Public Function EnsureSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    
    Set EnsureSheet = ws
End Function

' ── Format sheet as table with headers ──
Public Sub FormatAsTable(ws As Worksheet, headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i - LBound(headers) + 1).Value = headers(i)
        ws.Cells(1, i - LBound(headers) + 1).Font.Bold = True
    Next i
    
    ws.Rows(1).AutoFilter
    ws.Rows(1).Interior.Color = RGB(68, 84, 106)
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
End Sub

' ── Export DataLog to CSV ──
Public Sub ExportToCSV(Optional filePath As String = "")
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    If filePath = "" Then
        Dim wbPath As String
        wbPath = ThisWorkbook.Path

        ' ThisWorkbook.Path returns an https:// URL on OneDrive — not valid for file I/O
        If wbPath = "" Or Left(wbPath, 4) = "http" Then
            wbPath = "C:\Users\Steve\OneDrive\Documents\1__________\Work\MARIANA\tracker\EquityDerivTracker"
        End If

        filePath = wbPath & "\DataLog_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    End If
    
    Dim lr As Long, lc As Long
    lr = LastRow(wsLog, COL_LOG_KEY)
    lc = COL_LOG_TRADEDATE
    
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    
    Dim r As Long, c As Long
    Dim lineStr As String
    
    For r = 1 To lr
        lineStr = ""
        For c = 1 To lc
            If c > 1 Then lineStr = lineStr & ","
            Dim cellVal As String
            cellVal = Replace(CStr(wsLog.Cells(r, c).Value), ",", ";")
            lineStr = lineStr & """" & cellVal & """"
        Next c
        Print #fNum, lineStr
    Next r
    
    Close #fNum
    MsgBox "Exported to: " & filePath, vbInformation, "CSV Export"
End Sub
