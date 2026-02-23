'==============================================================================
' ThisWorkbook — Auto_Open for Unattended Mode
'==============================================================================
' When this workbook is opened (e.g. by Windows Task Scheduler),
' it will automatically start scanning if the file was opened
' with a command-line flag or if the AUTO_SCAN_ON_OPEN flag is set.
'==============================================================================

Option Explicit

Private Const AUTO_SCAN_ON_OPEN As Boolean = False  ' Set True for unattended mode

Private Sub Workbook_Open()
    ' Check if opened in unattended mode
    If AUTO_SCAN_ON_OPEN Then
        ' Update previous close first
        Call UpdatePrevClose

        ' Wait a moment for Bloomberg to initialize
        Application.Wait Now + TimeSerial(0, 0, 10)

        ' Start auto scanning
        Call StartAutoScan

    ElseIf InStr(1, Command(), "/autoscan", vbTextCompare) > 0 Then
        ' Also check: if command line contains "/autoscan"
        Call UpdatePrevClose
        Application.Wait Now + TimeSerial(0, 0, 10)
        Call StartAutoScan
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Stop any scheduled scans
    On Error Resume Next
    Call StopAutoScan
    On Error GoTo 0
End Sub
