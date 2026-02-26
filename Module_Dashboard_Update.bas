'==============================================================================
' Module_Dashboard UPDATE — Add these changes to your existing Module_Dashboard
'==============================================================================

' ── REPLACE your existing RefreshDashboard with this ──
' (keeps backward compatibility — no date param = today)

Public Sub RefreshDashboard()
    Call RefreshDashboardForDate(GetDashDate())
End Sub

' ── NEW: Refresh for any date ──
Public Sub RefreshDashboardForDate(tradeDate As Date)
    Dim wsDash As Worksheet
    Set wsDash = EnsureSheet(SHT_DASH)
    
    Application.ScreenUpdating = False
    
    ' Clear existing charts
    Dim cht As ChartObject
    For Each cht In wsDash.ChartObjects
        cht.Delete
    Next cht
    
    ' Clear data area (but preserve date cell Z1)
    wsDash.Range("A:Y").Clear
    
    ' ── Navigation Bar ──
    wsDash.Range("A1").Value = "EQUITY DERIVATIVES CROSS TRACKER"
    wsDash.Range("A1").Font.Size = 16
    wsDash.Range("A1").Font.Bold = True
    
    wsDash.Range("A2").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsDash.Range("A2").Font.Color = RGB(128, 128, 128)
    
    ' Date display with navigation hint
    Dim dateLabel As String
    If Int(tradeDate) = Int(Now) Then
        dateLabel = Format(tradeDate, "yyyy-mm-dd") & "  (TODAY)"
    Else
        dateLabel = Format(tradeDate, "yyyy-mm-dd  (dddd)")
    End If
    wsDash.Range("A3").Value = "Viewing: " & dateLabel
    wsDash.Range("A3").Font.Bold = True
    wsDash.Range("A3").Font.Size = 12
    
    ' Store the date
    wsDash.Range("Z1").Value = tradeDate
    wsDash.Columns("Z").Hidden = True
    
    ' Build charts for the selected date
    Call BuildIntradayPremiumChart(wsDash, tradeDate)
    Call BuildTradeSummary(wsDash, tradeDate)
    Call BuildProductBreakdown(wsDash, tradeDate)
    
    ' Add navigation buttons
    Call AddDashNavButtons(wsDash)
    
    Application.ScreenUpdating = True
End Sub

' ── Navigation Buttons on Dashboard ──
Private Sub AddDashNavButtons(wsDash As Worksheet)
    ' Remove existing nav buttons
    Dim shp As Shape
    For Each shp In wsDash.Shapes
        If shp.Type = msoFormControl Then
            If Left(shp.Name, 4) = "nav_" Or _
               shp.TextFrame.Characters.Text = Chr(9664) & " Prev Day" Or _
               shp.TextFrame.Characters.Text = "Next Day " & Chr(9654) Or _
               shp.TextFrame.Characters.Text = "Today" Or _
               shp.TextFrame.Characters.Text = "Pick Date" Then
                shp.Delete
            End If
        End If
    Next shp
    
    Dim btnTop As Double
    btnTop = wsDash.Cells(3, 6).Top  ' Row 3, column F area
    Dim btnLeft As Double
    btnLeft = wsDash.Cells(3, 6).Left
    
    ' ◀ Prev Day
    Dim btn As Shape
    Set btn = wsDash.Shapes.AddFormControl(xlButtonControl, _
              btnLeft, btnTop, 90, 25)
    btn.OnAction = "DashPrevDay"
    btn.TextFrame.Characters.Text = Chr(9664) & " Prev Day"
    btn.TextFrame.Characters.Font.Size = 9
    btn.Name = "nav_prev"
    
    ' Today
    Set btn = wsDash.Shapes.AddFormControl(xlButtonControl, _
              btnLeft + 95, btnTop, 65, 25)
    btn.OnAction = "DashToday"
    btn.TextFrame.Characters.Text = "Today"
    btn.TextFrame.Characters.Font.Size = 9
    btn.TextFrame.Characters.Font.Bold = True
    btn.Name = "nav_today"
    
    ' Next Day ▶
    Set btn = wsDash.Shapes.AddFormControl(xlButtonControl, _
              btnLeft + 165, btnTop, 90, 25)
    btn.OnAction = "DashNextDay"
    btn.TextFrame.Characters.Text = "Next Day " & Chr(9654)
    btn.TextFrame.Characters.Font.Size = 9
    btn.Name = "nav_next"
    
    ' Pick Date
    Set btn = wsDash.Shapes.AddFormControl(xlButtonControl, _
              btnLeft + 260, btnTop, 80, 25)
    btn.OnAction = "DashPickDate"
    btn.TextFrame.Characters.Text = "Pick Date"
    btn.TextFrame.Characters.Font.Size = 9
    btn.Name = "nav_pick"
End Sub
