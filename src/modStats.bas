Attribute VB_Name = "modStats"
Option Explicit
' =====================================================================
' modStats - Live hit-rate bars (K9)
'   rate = hits / (hits + misses), shown in % with 1 decimal.
'   Bar color: green >= 80 %, amber 50-79 %, red < 50 %,
'   grey when the level has no accesses yet.
'   Reads and writes count (STOREs are counted in L1D and L2).
'   Drawn on the Memory sheet (columns M..), K10 may move it.
' =====================================================================

Private Const STATS_HDR_ROW As Long = 19
Private Const STATS_COL As Long = 13        ' M: label
Private Const BAR_COL As Long = 15          ' O: bar starts here
Private Const BAR_MAX_WIDTH As Double = 150

' Hit rate between 0 and 1 (0 when there are no accesses)
Public Function HitRate(ByVal hits As Long, ByVal misses As Long) As Double
    If hits + misses = 0 Then
        HitRate = 0
    Else
        HitRate = hits / (hits + misses)
    End If
End Function

' Text such as "87.5 %"
Public Function RateText(ByVal hits As Long, ByVal misses As Long) As String
    RateText = Format(HitRate(hits, misses) * 100, "0.0") & " %"
End Function

' Bar color for a rate (0..1); grey if there are no accesses
Private Function RateColor(ByVal hits As Long, ByVal misses As Long) As Long
    Dim pct As Double
    If hits + misses = 0 Then
        RateColor = CLR_HIGHZ
        Exit Function
    End If
    pct = HitRate(hits, misses) * 100
    If pct >= 80 Then
        RateColor = CLR_HIT
    ElseIf pct >= 50 Then
        RateColor = CLR_LOAD
    Else
        RateColor = CLR_MISS
    End If
End Function

Private Function GetStatsSheet() As Worksheet
    On Error Resume Next
    Set GetStatsSheet = ThisWorkbook.Worksheets(SH_MEM)
    On Error GoTo 0
    If GetStatsSheet Is Nothing Then
        Set GetStatsSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetStatsSheet.Name = SH_MEM
    End If
End Function

Private Sub EnsureStatsTable(ByVal ws As Worksheet)
    ' Column widths are always applied so the text is never cut off
    ws.Columns(STATS_COL).ColumnWidth = 12
    ws.Columns(STATS_COL + 1).ColumnWidth = 9
    If CStr(ws.Cells(STATS_HDR_ROW, STATS_COL).value) = "Hit rate" Then Exit Sub
    With ws.Range(ws.Cells(STATS_HDR_ROW, STATS_COL), ws.Cells(STATS_HDR_ROW + 4, STATS_COL + 1))
        .NumberFormat = "@"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(STATS_HDR_ROW, STATS_COL).value = "Hit rate"
    ws.Cells(STATS_HDR_ROW, STATS_COL + 1).value = "%"
    ws.Range(ws.Cells(STATS_HDR_ROW, STATS_COL), ws.Cells(STATS_HDR_ROW, STATS_COL + 1)).Font.Bold = True
    ws.Cells(STATS_HDR_ROW + 1, STATS_COL).value = "L1I"
    ws.Cells(STATS_HDR_ROW + 2, STATS_COL).value = "L1D"
    ws.Cells(STATS_HDR_ROW + 3, STATS_COL).value = "L2"
    ws.Cells(STATS_HDR_ROW + 4, STATS_COL).value = "Page Faults"
End Sub

' Creates the bar shape on first use (name = bar_<key>)
Private Function GetBar(ByVal ws As Worksheet, ByVal key As String, ByVal rowNum As Long) As Shape
    Dim nm As String, c As Range
    nm = "bar_" & key
    Set c = ws.Cells(rowNum, BAR_COL)
    On Error Resume Next
    Set GetBar = ws.Shapes(nm)
    On Error GoTo 0
    If GetBar Is Nothing Then
        Set GetBar = ws.Shapes.AddShape(msoShapeRectangle, c.Left + 2, c.Top + 2, 1, c.Height - 4)
        GetBar.Name = nm
        GetBar.line.Visible = msoFalse
    End If
    GetBar.Left = c.Left + 2
    GetBar.Top = c.Top + 2
End Function

' Draws one level: text in the cell and the colored bar
Private Sub ShowRateBar(ByVal key As String, ByVal rowNum As Long, ByVal hits As Long, ByVal misses As Long)
    Dim ws As Worksheet, b As Shape, w As Double
    Set ws = GetStatsSheet()
    EnsureStatsTable ws
    ws.Cells(rowNum, STATS_COL + 1).value = RateText(hits, misses)
    Set b = GetBar(ws, key, rowNum)
    w = BAR_MAX_WIDTH * HitRate(hits, misses)
    If w < 1 Then w = 1
    b.Width = w
    b.fill.ForeColor.RGB = RateColor(hits, misses)
End Sub

' Redraws the three bars and the Page Faults counter (called by RefreshAll)
Public Sub RefreshStats()
    Dim ws As Worksheet
    ShowRateBar "l1i", STATS_HDR_ROW + 1, Stats.L1IHit, Stats.L1IMiss
    ShowRateBar "l1d", STATS_HDR_ROW + 2, Stats.L1DHit, Stats.L1DMiss
    ShowRateBar "l2", STATS_HDR_ROW + 3, Stats.L2Hit, Stats.l2Miss
    Set ws = GetStatsSheet()
    ws.Cells(STATS_HDR_ROW + 4, STATS_COL + 1).value = CStr(Stats.PageFaults)
End Sub

' Reset: every counter back to 0 and the view redrawn
Public Sub ResetStats()
    Dim emptyStats As TStats
    Stats = emptyStats
    RefreshStats
End Sub

' ---------- TESTS ----------
Public Sub Test_Stats()
    Dim fails As Integer, ws As Worksheet
    ResetStats
    Set ws = GetStatsSheet()

    ' 1) HitRate math
    If HitRate(7, 1) <> 0.875 Then fails = fails + 1: Debug.Print "  FAIL  1: HitRate(7,1) must be 0.875"
    If HitRate(0, 0) <> 0 Then fails = fails + 1: Debug.Print "  FAIL  1: HitRate(0,0) must be 0 (no division by zero)"
    If RateText(7, 1) <> Format(87.5, "0.0") & " %" Then fails = fails + 1: Debug.Print "  FAIL  1: RateText(7,1) = " & RateText(7, 1)

    ' 2) Reset: everything at 0 and grey
    If CStr(ws.Cells(STATS_HDR_ROW + 1, STATS_COL + 1).value) <> Format(0, "0.0") & " %" Then fails = fails + 1: Debug.Print "  FAIL  2: L1I must show 0.0 % after reset"
    If CStr(ws.Cells(STATS_HDR_ROW + 4, STATS_COL + 1).value) <> "0" Then fails = fails + 1: Debug.Print "  FAIL  2: Page Faults must show 0 after reset"
    If ws.Shapes("bar_l1i").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  2: bar without accesses must be grey"

    ' 3) Green (>= 80 %): 7 hits, 1 miss = 87.5 %
    Stats.L1IHit = 7: Stats.L1IMiss = 1
    ' 4) Amber (50-79 %): 3 hits, 3 misses = 50.0 %
    Stats.L1DHit = 3: Stats.L1DMiss = 3
    ' 5) Red (< 50 %): 1 hit, 4 misses = 20.0 %
    Stats.L2Hit = 1: Stats.l2Miss = 4
    Stats.PageFaults = 6
    RefreshStats

    If CStr(ws.Cells(STATS_HDR_ROW + 1, STATS_COL + 1).value) <> Format(87.5, "0.0") & " %" Then fails = fails + 1: Debug.Print "  FAIL  3: L1I text"
    If ws.Shapes("bar_l1i").fill.ForeColor.RGB <> CLR_HIT Then fails = fails + 1: Debug.Print "  FAIL  3: 87.5 % must be green"
    If Abs(ws.Shapes("bar_l1i").Width - 150 * 0.875) > 0.5 Then fails = fails + 1: Debug.Print "  FAIL  3: L1I bar width"
    If CStr(ws.Cells(STATS_HDR_ROW + 2, STATS_COL + 1).value) <> Format(50, "0.0") & " %" Then fails = fails + 1: Debug.Print "  FAIL  4: L1D text"
    If ws.Shapes("bar_l1d").fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  4: 50 % must be amber"
    If CStr(ws.Cells(STATS_HDR_ROW + 3, STATS_COL + 1).value) <> Format(20, "0.0") & " %" Then fails = fails + 1: Debug.Print "  FAIL  5: L2 text"
    If ws.Shapes("bar_l2").fill.ForeColor.RGB <> CLR_MISS Then fails = fails + 1: Debug.Print "  FAIL  5: 20 % must be red"
    If CStr(ws.Cells(STATS_HDR_ROW + 4, STATS_COL + 1).value) <> "6" Then fails = fails + 1: Debug.Print "  FAIL  5: Page Faults counter must show 6"

    ' 6) Boundary: 4 hits, 1 miss = 80 % is green; 49 % is red
    Stats.L1IHit = 4: Stats.L1IMiss = 1
    Stats.L1DHit = 49: Stats.L1DMiss = 51
    RefreshStats
    If ws.Shapes("bar_l1i").fill.ForeColor.RGB <> CLR_HIT Then fails = fails + 1: Debug.Print "  FAIL  6: 80 % must be green"
    If ws.Shapes("bar_l1d").fill.ForeColor.RGB <> CLR_MISS Then fails = fails + 1: Debug.Print "  FAIL  6: 49 % must be red"

    ' 7) Reset puts everything back to 0
    ResetStats
    If Stats.L1IHit <> 0 Or Stats.l2Miss <> 0 Or Stats.PageFaults <> 0 Then fails = fails + 1: Debug.Print "  FAIL  7: Reset must zero the counters"
    If CStr(ws.Cells(STATS_HDR_ROW + 4, STATS_COL + 1).value) <> "0" Then fails = fails + 1: Debug.Print "  FAIL  7: Page Faults view must be 0"

    Debug.Print IIf(fails = 0, "Test_Stats: ALL OK", "Test_Stats: " & fails & " FAIL(S)")
End Sub

' For the PR screenshot: real traffic through the hierarchy, then the bars.
' Afterwards run Test_Stats to leave everything clean.
Public Sub Demo_Stats()
    Dim i As Integer, v As Integer
    For i = 0 To 63: RAM(i) = i: Next i
    InvalidateAllCaches
    ResetStats
    For i = 0 To 15
        Clock = Clock + 1
        v = CacheReadL1I(i)              ' 2 lines: 2 misses, 14 hits
    Next i
    For i = 0 To 39 Step 8
        Clock = Clock + 1
        v = CacheReadL1D(i)              ' 5 different lines: misses
        CacheWriteL1D i + 1, 9           ' STORE: hit in L1D
    Next i
    Stats.PageFaults = 3
    RefreshStats
    Debug.Print "Demo_Stats: L1I " & RateText(Stats.L1IHit, Stats.L1IMiss) & _
        " | L1D " & RateText(Stats.L1DHit, Stats.L1DMiss) & _
        " | L2 " & RateText(Stats.L2Hit, Stats.l2Miss) & _
        " | Page Faults " & Stats.PageFaults
End Sub

