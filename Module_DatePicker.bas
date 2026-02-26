'==============================================================================
' Module_DatePicker — Dashboard Date Navigation
'==============================================================================
Option Explicit

Private Const DATE_CELL As String = "Z1"  ' Hidden cell to store current view date

' ── Get current dashboard date (defaults to today) ──
Public Function GetDashDate() As Date
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(SHT_DASH)
    
    If IsDate(wsDash.Range(DATE_CELL).Value) And wsDash.Range(DATE_CELL).Value > 0 Then
        GetDashDate = Int(wsDash.Range(DATE_CELL).Value)
    Else
        GetDashDate = Int(Now)
    End If
End Function

' ── Set dashboard date ──
Private Sub SetDashDate(d As Date)
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(SHT_DASH)
    wsDash.Range(DATE_CELL).Value = d
End Sub

' ── Go to previous trading day ──
Public Sub DashPrevDay()
    Dim current As Date
    current = GetDashDate()
    
    ' Step back, skipping weekends and holidays
    Dim prev As Date
    prev = GetPrevMarketDay(current)
    
    ' Check if there's any data for this date; if not, keep going back (max 30 days)
    Dim attempts As Long
    attempts = 0
    Do While Not HasTradesOnDate(prev) And attempts < 30
        prev = GetPrevMarketDay(prev)
        attempts = attempts + 1
    Loop
    
    ' If no data found in 30 days, just show the date anyway
    If attempts >= 30 Then prev = GetPrevMarketDay(current)
    
    SetDashDate prev
    Call RefreshDashboardForDate(prev)
End Sub

' ── Go to next trading day ──
Public Sub DashNextDay()
    Dim current As Date
    current = GetDashDate()
    
    ' Don't go past today
    If current >= Int(Now) Then
        MsgBox "Already showing the latest date.", vbInformation, "Dashboard"
        Exit Sub
    End If
    
    ' Step forward, skipping weekends and holidays
    Dim nxt As Date
    nxt = GetNextMarketDay(current)
    
    ' Don't go past today
    If nxt > Int(Now) Then nxt = Int(Now)
    
    SetDashDate nxt
    Call RefreshDashboardForDate(nxt)
End Sub

' ── Jump to today ──
Public Sub DashToday()
    Dim today As Date
    today = Int(Now)
    SetDashDate today
    Call RefreshDashboardForDate(today)
End Sub

' ── Jump to specific date via input box ──
Public Sub DashPickDate()
    Dim current As Date
    current = GetDashDate()
    
    Dim inputDate As String
    inputDate = InputBox("Enter date to view:" & vbCrLf & vbCrLf & _
                         "Format: YYYY-MM-DD", _
                         "Select Dashboard Date", _
                         Format(current, "yyyy-mm-dd"))
    
    If inputDate = "" Then Exit Sub
    
    If Not IsDate(inputDate) Then
        MsgBox "Invalid date. Please use YYYY-MM-DD format.", vbExclamation, "Date Picker"
        Exit Sub
    End If
    
    Dim selectedDate As Date
    selectedDate = Int(CDate(inputDate))
    
    If selectedDate > Int(Now) Then
        MsgBox "Cannot view future dates.", vbExclamation, "Date Picker"
        Exit Sub
    End If
    
    SetDashDate selectedDate
    Call RefreshDashboardForDate(selectedDate)
End Sub

' ── Get next market day (skips weekends + holidays) ──
Public Function GetNextMarketDay(Optional fromDate As Date = 0) As Date
    If fromDate = 0 Then fromDate = Date
    
    Dim candidate As Date
    candidate = fromDate + 1
    
    Dim wsHol As Worksheet
    On Error Resume Next
    Set wsHol = ThisWorkbook.Sheets(SHT_HOLIDAYS)
    On Error GoTo 0
    
    Dim maxLookforward As Long
    maxLookforward = 0
    
    Do
        Dim skipDay As Boolean
        skipDay = False
        
        ' Skip weekends
        If Weekday(candidate, vbMonday) > 5 Then skipDay = True
        
        ' Skip holidays
        If Not skipDay And Not wsHol Is Nothing Then
            Dim lr As Long
            lr = LastRow(wsHol, COL_HOL_DATE)
            Dim i As Long
            For i = 2 To lr
                If IsDate(wsHol.Cells(i, COL_HOL_DATE).Value) Then
                    If Int(wsHol.Cells(i, COL_HOL_DATE).Value) = Int(candidate) Then
                        skipDay = True
                        Exit For
                    End If
                End If
            Next i
        End If
        
        If Not skipDay Then
            GetNextMarketDay = candidate
            Exit Function
        End If
        
        candidate = candidate + 1
        maxLookforward = maxLookforward + 1
    Loop While maxLookforward <= 30
    
    ' Fallback
    GetNextMarketDay = fromDate + 1
End Function

' ── Check if DataLog has trades for a given date ──
Private Function HasTradesOnDate(tradeDate As Date) As Boolean
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHT_LOG)
    
    Dim lr As Long
    lr = LastRow(wsLog, COL_LOG_KEY)
    
    Dim i As Long
    For i = 2 To lr
        If Int(wsLog.Cells(i, COL_LOG_TRADEDATE).Value) = Int(tradeDate) Then
            HasTradesOnDate = True
            Exit Function
        End If
    Next i
    
    HasTradesOnDate = False
End Function
