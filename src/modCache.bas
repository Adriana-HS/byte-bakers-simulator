Attribute VB_Name = "modCache"
Option Explicit
' =====================================================================
' modCache - Cache hierarchy (K6: L1I; K7: L1D; K8: L2)
'   L1I: 8 B (1 line), direct mapped, instruction fetch only.
'   L1D: 8 B (1 line), direct mapped, LOAD/STORE only.
'   L2 : 32 B (4 lines of 8 B), fully associative, LRU, unified.
'   Works with PHYSICAL addresses (after the MMU).
'   tag = paddr \ LINE_SIZE
'
' WRITE POLICY: write-through + no-write-allocate (L1D and L2)
'   - STORE hit : the line is updated AND RAM is written.
'   - STORE miss: the line is NOT brought into the cache (no-write-allocate),
'                 RAM is written anyway.
'   - RAM is ALWAYS up to date. Writes also count hit/miss in L1D and L2.
'
' The I/O page (15) never gets here: MemRead/MemWrite bypass the cache for it.
' =====================================================================

Private Const CACHE_HDR_ROW As Long = 19
Private Const L1I_ROW As Long = 20
Private Const L1D_ROW As Long = 21
Private Const L2_ROW0 As Long = 22   ' L2 lines: rows 22..25

Private Const SHP_L1I As String = "l1i_box"
Private Const SHP_L1D As String = "l1d_box"
Private Const SHP_L2 As String = "l2_line"   ' + index (0..3)

' ---------- K6: L1I ----------

Public Function CacheReadL1I(ByVal paddr As Integer) As Integer
    Dim tag As Integer
    CheckPaddr paddr
    tag = paddr \ LINE_SIZE
    If L1I.Valid And L1I.tag = tag Then
        Stats.L1IHit = Stats.L1IHit + 1
        FlashCache SHP_L1I, True
        ShowResultAt L1I_ROW, L1I, "HIT"
    Else
        Stats.L1IMiss = Stats.L1IMiss + 1
        FlashCache SHP_L1I, False
        FillFromL2 L1I, tag
        ShowLineAt L1I_ROW, L1I, "MISS"
    End If
    L1I.LastUse = Clock
    CacheReadL1I = L1I.Data(paddr Mod LINE_SIZE)
End Function

' ---------- K7: L1D ----------

Public Function CacheReadL1D(ByVal paddr As Integer) As Integer
    Dim tag As Integer
    CheckPaddr paddr
    tag = paddr \ LINE_SIZE
    If L1D.Valid And L1D.tag = tag Then
        Stats.L1DHit = Stats.L1DHit + 1
        FlashCache SHP_L1D, True
        ShowResultAt L1D_ROW, L1D, "HIT"
    Else
        Stats.L1DMiss = Stats.L1DMiss + 1
        FlashCache SHP_L1D, False
        FillFromL2 L1D, tag
        ShowLineAt L1D_ROW, L1D, "MISS"
    End If
    L1D.LastUse = Clock
    CacheReadL1D = L1D.Data(paddr Mod LINE_SIZE)
End Function

' STORE: write-through + no-write-allocate.
'  - HIT : updates the byte in the L1D line.
'  - MISS: the line is NOT loaded.
'  - Both cases count hit/miss in L1D, then L2 is handled, and RAM is ALWAYS written.
Public Sub CacheWriteL1D(ByVal paddr As Integer, ByVal value As Integer)
    Dim tag As Integer
    CheckPaddr paddr
    value = value And BYTE_MAX
    tag = paddr \ LINE_SIZE
    If L1D.Valid And L1D.tag = tag Then
        Stats.L1DHit = Stats.L1DHit + 1
        L1D.Data(paddr Mod LINE_SIZE) = value
        L1D.LastUse = Clock
        FlashCache SHP_L1D, True
        ShowLineAt L1D_ROW, L1D, "HIT"
    Else
        Stats.L1DMiss = Stats.L1DMiss + 1
        FlashCache SHP_L1D, False
        ShowResultAt L1D_ROW, L1D, "MISS"
    End If
    WriteThroughToL2 paddr, value
    RAM(paddr) = value
End Sub

' ---------- K8: L2 ----------

' Write-through to L2 (no-write-allocate): on a hit the L2 line is updated
' and its LastUse refreshed; on a miss nothing is loaded. Both are counted.
Private Sub WriteThroughToL2(ByVal paddr As Integer, ByVal value As Integer)
    Dim tag As Integer, idx As Integer
    tag = paddr \ LINE_SIZE
    idx = L2Find(tag)
    If idx >= 0 Then
        Stats.L2Hit = Stats.L2Hit + 1
        L2(idx).Data(paddr Mod LINE_SIZE) = value
        L2(idx).LastUse = Clock
        FlashCache SHP_L2 & idx, True
        ShowL2Row idx, "HIT"
    Else
        Stats.l2Miss = Stats.l2Miss + 1
        ' no-write-allocate: nothing is loaded; no L2 line to animate
    End If
End Sub

' Reading path: look in L2 first, go down to RAM only on an L2 miss.
Private Sub FillFromL2(ByRef l1 As TCacheLine, ByVal tag As Integer)
    Dim idx As Integer, i As Integer
    idx = L2Find(tag)
    If idx >= 0 Then
        Stats.L2Hit = Stats.L2Hit + 1
        FlashCache SHP_L2 & idx, True
        ShowL2Row idx, "HIT"
    Else
        Stats.l2Miss = Stats.l2Miss + 1
        idx = L2Victim()
        FlashCache SHP_L2 & idx, False        ' highlight the victim before replacing
        L2(idx).Valid = True
        L2(idx).tag = tag
        For i = 0 To LINE_SIZE - 1
            L2(idx).Data(i) = RAM(tag * LINE_SIZE + i)
        Next i
        ShowL2Row idx, "MISS"
    End If
    L2(idx).LastUse = Clock
    l1.Valid = True
    l1.tag = tag
    For i = 0 To LINE_SIZE - 1
        l1.Data(i) = L2(idx).Data(i)
    Next i
    UI_AnimateLineFill idx
End Sub

' Index of the L2 line holding this tag, or -1
Private Function L2Find(ByVal tag As Integer) As Integer
    Dim i As Integer
    For i = 0 To L2_LINES - 1
        If L2(i).Valid And L2(i).tag = tag Then
            L2Find = i
            Exit Function
        End If
    Next i
    L2Find = -1
End Function

' Victim: first invalid line, otherwise the lowest LastUse (LRU)
Private Function L2Victim() As Integer
    Dim i As Integer, best As Integer
    best = 0
    For i = 0 To L2_LINES - 1
        If Not L2(i).Valid Then
            L2Victim = i
            Exit Function
        End If
        If L2(i).LastUse < L2(best).LastUse Then best = i
    Next i
    L2Victim = best
End Function

Public Sub InvalidateAllCaches()
    Dim i As Integer
    L1I.Valid = False
    L1D.Valid = False
    For i = 0 To L2_LINES - 1
        L2(i).Valid = False
    Next i
End Sub

' Called by HandlePageFault when a frame is evicted: every line that
' belongs to that frame is invalidated in L1I, L1D and L2.
Public Sub InvalidateFrame(ByVal pfn As Integer)
    Dim firstTag As Integer, lastTag As Integer, i As Integer
    firstTag = pfn * PAGE_SIZE \ LINE_SIZE
    lastTag = firstTag + (PAGE_SIZE \ LINE_SIZE) - 1
    If L1I.Valid Then
        If L1I.tag >= firstTag And L1I.tag <= lastTag Then L1I.Valid = False
    End If
    If L1D.Valid Then
        If L1D.tag >= firstTag And L1D.tag <= lastTag Then L1D.Valid = False
    End If
    For i = 0 To L2_LINES - 1
        If L2(i).Valid Then
            If L2(i).tag >= firstTag And L2(i).tag <= lastTag Then L2(i).Valid = False
        End If
    Next i
End Sub

Private Sub CheckPaddr(ByVal paddr As Integer)
    If paddr < 0 Or paddr > PADDR_MAX Then
        Err.Raise vbObjectError + 550, "Cache", "Physical address out of range (0-" & PADDR_MAX & "): " & paddr
    End If
End Sub

' ---------- Animation (optional, done by modUI if the shape exists) ----------

Private Sub FlashCache(ByVal shapeName As String, ByVal isHit As Boolean)
    If ShapeExists(shapeName) Then UI_CacheFlash shapeName, isHit
End Sub

Private Function ShapeExists(ByVal nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(SH_SIM).Shapes(nm)
    If s Is Nothing Then Set s = ThisWorkbook.Worksheets(SH_MEM).Shapes(nm)
    On Error GoTo 0
    ShapeExists = Not (s Is Nothing)
End Function

' ---------- Cache line view on the Memory sheet ----------

Private Function Hex2(ByVal v As Integer) As String
    Hex2 = Right$("0" & Hex$(v And BYTE_MAX), 2)
End Function

Private Function GetCacheSheet() As Worksheet
    On Error Resume Next
    Set GetCacheSheet = ThisWorkbook.Worksheets(SH_MEM)
    On Error GoTo 0
    If GetCacheSheet Is Nothing Then
        Set GetCacheSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetCacheSheet.Name = SH_MEM
    End If
End Function

' Draws the headers the first time: Cache | Result | Tag | B0..B7
Private Sub EnsureCacheTable(ByVal ws As Worksheet)
    Dim i As Integer
    If CStr(ws.Cells(CACHE_HDR_ROW, 1).value) = "Cache" Then Exit Sub
    With ws.Range(ws.Cells(CACHE_HDR_ROW, 1), ws.Cells(CACHE_HDR_ROW + 6, 11))
        .NumberFormat = "@"
        .Font.Name = "Consolas"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(CACHE_HDR_ROW, 1).value = "Cache"
    ws.Cells(CACHE_HDR_ROW, 2).value = "Result"
    ws.Cells(CACHE_HDR_ROW, 3).value = "Tag"
    For i = 0 To LINE_SIZE - 1
        ws.Cells(CACHE_HDR_ROW, 4 + i).value = "B" & i
    Next i
    ws.Range(ws.Cells(CACHE_HDR_ROW, 1), ws.Cells(CACHE_HDR_ROW, 11)).Font.Bold = True
End Sub

Private Function RowLabel(ByVal rowNum As Long) As String
    If rowNum = L1D_ROW Then
        RowLabel = "L1D"
    ElseIf rowNum >= L2_ROW0 Then
        RowLabel = "L2-" & (rowNum - L2_ROW0)
    Else
        RowLabel = "L1I"
    End If
End Function

' Redraws a whole cache row: result, tag and the 8 bytes in hexadecimal
Private Sub ShowLineAt(ByVal rowNum As Long, ByRef line As TCacheLine, ByVal result As String)
    Dim ws As Worksheet, i As Integer
    Dim cellsRow(0 To 0, 0 To 9) As String
    Set ws = GetCacheSheet()
    EnsureCacheTable ws
    ws.Cells(rowNum, 1).value = RowLabel(rowNum)
    cellsRow(0, 0) = result
    If line.Valid Then
        cellsRow(0, 1) = CStr(line.tag)
        For i = 0 To LINE_SIZE - 1
            cellsRow(0, 2 + i) = Hex2(line.Data(i))
        Next i
    Else
        cellsRow(0, 1) = "-"
        For i = 0 To LINE_SIZE - 1
            cellsRow(0, 2 + i) = "--"
        Next i
    End If
    ws.Range(ws.Cells(rowNum, 2), ws.Cells(rowNum, 11)).value = cellsRow
    Select Case result
        Case "HIT": ws.Cells(rowNum, 2).Interior.Color = CLR_HIT
        Case "MISS": ws.Cells(rowNum, 2).Interior.Color = CLR_MISS
        Case Else: ws.Cells(rowNum, 2).Interior.Pattern = xlNone
    End Select
End Sub

' The line does not change (read hit / write miss): only the result cell is updated
Private Sub ShowResultAt(ByVal rowNum As Long, ByRef line As TCacheLine, ByVal result As String)
    Dim ws As Worksheet
    Set ws = GetCacheSheet()
    If CStr(ws.Cells(CACHE_HDR_ROW, 1).value) <> "Cache" Then
        ShowLineAt rowNum, line, result
        Exit Sub
    End If
    ws.Cells(rowNum, 1).value = RowLabel(rowNum)
    ws.Cells(rowNum, 2).value = result
    If result = "HIT" Then
        ws.Cells(rowNum, 2).Interior.Color = CLR_HIT
    Else
        ws.Cells(rowNum, 2).Interior.Color = CLR_MISS
    End If
End Sub

' One L2 line (rows 22..25)
Private Sub ShowL2Row(ByVal idx As Integer, ByVal result As String)
    ShowLineAt L2_ROW0 + idx, L2(idx), result
End Sub

Private Sub ShowL1ILine(ByVal result As String)
    ShowLineAt L1I_ROW, L1I, result
End Sub

Private Sub ShowL1DLine(ByVal result As String)
    ShowLineAt L1D_ROW, L1D, result
End Sub

' Resets the four L2 rows of the sheet to "-"
Private Sub ShowAllL2Rows(ByVal result As String)
    Dim i As Integer
    For i = 0 To L2_LINES - 1
        ShowL2Row i, result
    Next i
End Sub

' ---------- TESTS ----------
' Run with F5 (cursor inside the Sub) and check the Immediate window (Ctrl+G)

' Clears every cache line, counters and the sheet rows
Private Sub ResetCachesForTest()
    Dim emptyStats As TStats, emptyLine As TCacheLine, i As Integer
    L1I = emptyLine
    L1D = emptyLine
    For i = 0 To L2_LINES - 1
        L2(i) = emptyLine
    Next i
    Stats = emptyStats
End Sub

Public Sub Test_CacheL1I()
    Dim fails As Integer, i As Integer, v As Integer, ok As Boolean
    Dim ws As Worksheet
    Set ws = GetCacheSheet()

    For i = 0 To 31: RAM(i) = 100 + i: Next i
    ResetCachesForTest

    ' 1) First access: MISS, loads line 0 (bytes 0..7)
    v = CacheReadL1I(0)
    If v <> 100 Then fails = fails + 1: Debug.Print "  FAIL  1: CacheReadL1I(0) = " & v & " (expected 100)"
    If Stats.L1IMiss <> 1 Or Stats.L1IHit <> 0 Then fails = fails + 1: Debug.Print "  FAIL  1: expected 1 miss and 0 hits"
    If Not L1I.Valid Or L1I.tag <> 0 Then fails = fails + 1: Debug.Print "  FAIL  1: the line should be valid with tag 0"
    If CStr(ws.Cells(L1I_ROW, 2).value) <> "MISS" Or CStr(ws.Cells(L1I_ROW, 3).value) <> "0" _
       Or CStr(ws.Cells(L1I_ROW, 4).value) <> "64" Or CStr(ws.Cells(L1I_ROW, 11).value) <> "6B" Then
        fails = fails + 1: Debug.Print "  FAIL  1: the sheet should show MISS, tag 0, bytes 64..6B"
    End If

    ' 2) The other 7 bytes of the line: 7 HITS
    For i = 1 To 7
        v = CacheReadL1I(i)
        If v <> 100 + i Then fails = fails + 1: Debug.Print "  FAIL  2: CacheReadL1I(" & i & ") = " & v & " (expected " & 100 + i & ")"
    Next i
    If Stats.L1IHit <> 7 Or Stats.L1IMiss <> 1 Then fails = fails + 1: Debug.Print "  FAIL  2: 8 consecutive reads must give 1 miss + 7 hits (" & Stats.L1IMiss & " miss, " & Stats.L1IHit & " hits)"
    If CStr(ws.Cells(L1I_ROW, 2).value) <> "HIT" Then fails = fails + 1: Debug.Print "  FAIL  2: the sheet should show HIT"

    ' 3) Line change: MISS (tag 1)
    v = CacheReadL1I(8)
    If v <> 108 Then fails = fails + 1: Debug.Print "  FAIL  3: CacheReadL1I(8) = " & v & " (expected 108)"
    If Stats.L1IMiss <> 2 Then fails = fails + 1: Debug.Print "  FAIL  3: expected 2 misses"
    If L1I.tag <> 1 Then fails = fails + 1: Debug.Print "  FAIL  3: the tag should be 1"
    ok = True
    For i = 0 To 7
        If L1I.Data(i) <> 108 + i Then ok = False
    Next i
    If Not ok Then fails = fails + 1: Debug.Print "  FAIL  3: the line does not hold bytes 108..115"
    If CStr(ws.Cells(L1I_ROW, 2).value) <> "MISS" Or CStr(ws.Cells(L1I_ROW, 3).value) <> "1" _
       Or CStr(ws.Cells(L1I_ROW, 4).value) <> "6C" Or CStr(ws.Cells(L1I_ROW, 11).value) <> "73" Then
        fails = fails + 1: Debug.Print "  FAIL  3: the sheet should show MISS, tag 1, bytes 6C..73"
    End If

    ' 4) A HIT does not go down to RAM
    RAM(10) = 250
    v = CacheReadL1I(10)
    If v <> 110 Then fails = fails + 1: Debug.Print "  FAIL  4: a HIT must be answered by the cache (110), got " & v
    If Stats.L1IHit <> 8 Or Stats.L1IMiss <> 2 Then fails = fails + 1: Debug.Print "  FAIL  4: expected 8 hits and 2 misses"

    ' 5) Going back to line 0: L1I MISS, but L2 still has line 1 (old data 110),
    '    so line 1 is served by L2 (that is the point of a cache hierarchy)
    v = CacheReadL1I(0)
    If v <> 100 Or Stats.L1IMiss <> 3 Then fails = fails + 1: Debug.Print "  FAIL  5: going back to line 0 must be a MISS and return 100"
    v = CacheReadL1I(10)
    If v <> 110 Or Stats.L1IMiss <> 4 Then fails = fails + 1: Debug.Print "  FAIL  5: L1I MISS on line 1 must be served by L2 with the OLD value 110, got " & v
    v = CacheReadL1I(10)
    If v <> 110 Or Stats.L1IHit <> 9 Then fails = fails + 1: Debug.Print "  FAIL  5: the second access must be a HIT returning 110"

    ' 6) Invalid line: even with a matching tag it is a MISS
    L1I.Valid = False
    v = CacheReadL1I(10)
    If v <> 110 Or Stats.L1IMiss <> 5 Or Stats.L1IHit <> 9 Then fails = fails + 1: Debug.Print "  FAIL  6: an invalid line must give a MISS (miss=" & Stats.L1IMiss & ", hit=" & Stats.L1IHit & ")"

    ' 7) Physical address out of range
    On Error Resume Next
    v = CacheReadL1I(256)
    If Err.Number <> vbObjectError + 550 Then fails = fails + 1: Debug.Print "  FAIL  7: CacheReadL1I(256) did not raise error 550"
    Err.Clear
    v = CacheReadL1I(-1)
    If Err.Number <> vbObjectError + 550 Then fails = fails + 1: Debug.Print "  FAIL  7: CacheReadL1I(-1) did not raise error 550"
    Err.Clear
    On Error GoTo 0

    ' Cleanup
    For i = 0 To 31: RAM(i) = 0: Next i
    ResetCachesForTest
    ShowL1ILine "-"
    ShowAllL2Rows "-"
    If CStr(ws.Cells(L1I_ROW, 3).value) <> "-" Then fails = fails + 1: Debug.Print "  FAIL  cleanup: the L1I row was not reset to '-'"

    Debug.Print IIf(fails = 0, "Test_CacheL1I: ALL OK", "Test_CacheL1I: " & fails & " FAIL(S)")
End Sub

Public Sub Test_CacheL1D()
    Dim fails As Integer, i As Integer, v As Integer
    Dim ws As Worksheet
    Set ws = GetCacheSheet()

    For i = 0 To 47: RAM(i) = 100 + i: Next i
    ResetCachesForTest

    ' 1) First LOAD: MISS, loads line 0
    v = CacheReadL1D(0)
    If v <> 100 Then fails = fails + 1: Debug.Print "  FAIL  1: CacheReadL1D(0) = " & v & " (expected 100)"
    If Stats.L1DMiss <> 1 Or Stats.L1DHit <> 0 Then fails = fails + 1: Debug.Print "  FAIL  1: expected 1 miss and 0 hits"
    If Not L1D.Valid Or L1D.tag <> 0 Then fails = fails + 1: Debug.Print "  FAIL  1: the line should be valid with tag 0"
    If CStr(ws.Cells(L1D_ROW, 1).value) <> "L1D" Or CStr(ws.Cells(L1D_ROW, 2).value) <> "MISS" _
       Or CStr(ws.Cells(L1D_ROW, 3).value) <> "0" Or CStr(ws.Cells(L1D_ROW, 4).value) <> "64" _
       Or CStr(ws.Cells(L1D_ROW, 11).value) <> "6B" Then
        fails = fails + 1: Debug.Print "  FAIL  1: row 21 should show L1D, MISS, tag 0, bytes 64..6B"
    End If

    ' 2) The other 7 bytes: 7 HITS
    For i = 1 To 7
        v = CacheReadL1D(i)
        If v <> 100 + i Then fails = fails + 1: Debug.Print "  FAIL  2: CacheReadL1D(" & i & ") = " & v & " (expected " & 100 + i & ")"
    Next i
    If Stats.L1DHit <> 7 Or Stats.L1DMiss <> 1 Then fails = fails + 1: Debug.Print "  FAIL  2: 8 reads of one line must give 1 miss + 7 hits (" & Stats.L1DMiss & " miss, " & Stats.L1DHit & " hits)"
    If CStr(ws.Cells(L1D_ROW, 2).value) <> "HIT" Then fails = fails + 1: Debug.Print "  FAIL  2: the sheet should show HIT"

    ' 3) STORE hit: updates the L1D line AND RAM
    CacheWriteL1D 3, 200
    If L1D.Data(3) <> 200 Then fails = fails + 1: Debug.Print "  FAIL  3: the L1D line should hold 200 at byte 3"
    If RAM(3) <> 200 Then fails = fails + 1: Debug.Print "  FAIL  3: write-through: RAM(3) should be 200"
    If Stats.L1DHit <> 8 Or Stats.L1DMiss <> 1 Then fails = fails + 1: Debug.Print "  FAIL  3: a STORE hit must count as a hit (8 hits, 1 miss)"
    If CStr(ws.Cells(L1D_ROW, 2).value) <> "HIT" Or CStr(ws.Cells(L1D_ROW, 7).value) <> "C8" Then fails = fails + 1: Debug.Print "  FAIL  3: the sheet should show HIT and C8 in B3"
    v = CacheReadL1D(3)
    If v <> 200 Or Stats.L1DHit <> 9 Then fails = fails + 1: Debug.Print "  FAIL  3: a LOAD after the STORE hit must be a HIT returning 200"

    ' 4) STORE miss: the line is NOT loaded in L1D, RAM is written
    CacheWriteL1D 20, 77
    If RAM(20) <> 77 Then fails = fails + 1: Debug.Print "  FAIL  4: write-through: RAM(20) should be 77"
    If Not L1D.Valid Or L1D.tag <> 0 Then fails = fails + 1: Debug.Print "  FAIL  4: no-write-allocate: L1D must still hold line 0"
    If Stats.L1DMiss <> 2 Or Stats.L1DHit <> 9 Then fails = fails + 1: Debug.Print "  FAIL  4: a STORE miss must count as a miss (2 misses, 9 hits)"
    If CStr(ws.Cells(L1D_ROW, 2).value) <> "MISS" Or CStr(ws.Cells(L1D_ROW, 3).value) <> "0" Then fails = fails + 1: Debug.Print "  FAIL  4: the sheet should show MISS and keep tag 0"

    ' 5) LOAD after a STORE miss: MISS, returns the new value
    v = CacheReadL1D(20)
    If v <> 77 Then fails = fails + 1: Debug.Print "  FAIL  5: CacheReadL1D(20) = " & v & " (expected 77)"
    If Stats.L1DMiss <> 3 Or L1D.tag <> 2 Then fails = fails + 1: Debug.Print "  FAIL  5: expected a MISS that loads tag 2"

    ' 6) STORE miss with an INVALID line: it stays invalid
    L1D.Valid = False
    CacheWriteL1D 40, 5
    If L1D.Valid Then fails = fails + 1: Debug.Print "  FAIL  6: a STORE miss must not make the line valid"
    If RAM(40) <> 5 Or Stats.L1DMiss <> 4 Then fails = fails + 1: Debug.Print "  FAIL  6: RAM(40) should be 5 and the misses 4"
    v = CacheReadL1D(40)
    If v <> 5 Or Stats.L1DMiss <> 5 Then fails = fails + 1: Debug.Print "  FAIL  6: the LOAD must be a MISS returning 5"

    ' 7) Values are stored as one byte (300 -> 44)
    CacheWriteL1D 0, 300
    If RAM(0) <> 44 Then fails = fails + 1: Debug.Print "  FAIL  7: RAM(0) should be 44 (300 And 255), got " & RAM(0)
    If Stats.L1DMiss <> 6 Then fails = fails + 1: Debug.Print "  FAIL  7: expected 6 misses"

    ' 8) L1D counters are independent from L1I
    If Stats.L1IHit <> 0 Or Stats.L1IMiss <> 0 Then fails = fails + 1: Debug.Print "  FAIL  8: L1I counters must stay at 0"

    ' 9) Physical address out of range (read and write)
    On Error Resume Next
    v = CacheReadL1D(256)
    If Err.Number <> vbObjectError + 550 Then fails = fails + 1: Debug.Print "  FAIL  9: CacheReadL1D(256) did not raise error 550"
    Err.Clear
    CacheWriteL1D -1, 1
    If Err.Number <> vbObjectError + 550 Then fails = fails + 1: Debug.Print "  FAIL  9: CacheWriteL1D(-1) did not raise error 550"
    Err.Clear
    On Error GoTo 0

    ' Cleanup
    For i = 0 To 47: RAM(i) = 0: Next i
    ResetCachesForTest
    ShowL1DLine "-"
    ShowAllL2Rows "-"
    If CStr(ws.Cells(L1D_ROW, 3).value) <> "-" Then fails = fails + 1: Debug.Print "  FAIL  cleanup: the L1D row was not reset to '-'"

    Debug.Print IIf(fails = 0, "Test_CacheL1D: ALL OK", "Test_CacheL1D: " & fails & " FAIL(S)")
End Sub

Public Sub Test_CacheL2()
    Dim fails As Integer, i As Integer, v As Integer
    Dim ws As Worksheet
    Set ws = GetCacheSheet()

    For i = 0 To 63: RAM(i) = 100 + i: Next i
    ResetCachesForTest

    ' 1) First LOAD: L1D miss AND L2 miss (line 0 enters L2 line 0)
    Clock = 1
    v = CacheReadL1D(0)
    If v <> 100 Then fails = fails + 1: Debug.Print "  FAIL  1: CacheReadL1D(0) = " & v & " (expected 100)"
    If Stats.l2Miss <> 1 Or Stats.L2Hit <> 0 Then fails = fails + 1: Debug.Print "  FAIL  1: expected 1 L2 miss and 0 L2 hits"
    If Not L2(0).Valid Or L2(0).tag <> 0 Then fails = fails + 1: Debug.Print "  FAIL  1: L2 line 0 should hold tag 0"
    If CStr(ws.Cells(L2_ROW0, 1).value) <> "L2-0" Or CStr(ws.Cells(L2_ROW0, 2).value) <> "MISS" Then fails = fails + 1: Debug.Print "  FAIL  1: row 22 should show L2-0 and MISS"

    ' 2) L1I reads the same line: L1I miss, but L2 HIT (unified)
    Clock = 2
    v = CacheReadL1I(1)
    If v <> 101 Then fails = fails + 1: Debug.Print "  FAIL  2: CacheReadL1I(1) = " & v & " (expected 101)"
    If Stats.L1IMiss <> 1 Or Stats.L2Hit <> 1 Or Stats.l2Miss <> 1 Then fails = fails + 1: Debug.Print "  FAIL  2: expected L1I miss, L2 hit=1, L2 miss=1"
    If CStr(ws.Cells(L2_ROW0, 2).value) <> "HIT" Then fails = fails + 1: Debug.Print "  FAIL  2: row 22 should show HIT"

    ' 3) Fill the 4 L2 lines with tags 0..3 (reading lines 1, 2, 3 through L1D)
    Clock = 3: v = CacheReadL1D(8)
    Clock = 4: v = CacheReadL1D(16)
    Clock = 5: v = CacheReadL1D(24)
    If Stats.l2Miss <> 4 Then fails = fails + 1: Debug.Print "  FAIL  3: expected 4 L2 misses, got " & Stats.l2Miss
    For i = 0 To 3
        If Not L2(i).Valid Or L2(i).tag <> i Then fails = fails + 1: Debug.Print "  FAIL  3: L2 line " & i & " should hold tag " & i
    Next i

    ' 4) A 5th line (tag 4): LRU victim is the line with the lowest LastUse.
    '    L2 line 0 was used at Clock 2, lines 1..3 at 3, 4, 5 -> victim = line 0
    Clock = 6
    v = CacheReadL1D(32)
    If v <> 132 Then fails = fails + 1: Debug.Print "  FAIL  4: CacheReadL1D(32) = " & v & " (expected 132)"
    If L2(0).tag <> 4 Then fails = fails + 1: Debug.Print "  FAIL  4: LRU must replace L2 line 0 (tag 4), got tag " & L2(0).tag
    If L2(1).tag <> 1 Or L2(2).tag <> 2 Or L2(3).tag <> 3 Then fails = fails + 1: Debug.Print "  FAIL  4: lines 1..3 must stay untouched"
    If Stats.l2Miss <> 5 Then fails = fails + 1: Debug.Print "  FAIL  4: expected 5 L2 misses"

    ' 5) LRU keeps a recently used line: touch tag 1 (L2 line 1), then bring tag 5.
    '    Lowest LastUse is now L2 line 2 (Clock 4) -> victim = line 2
    Clock = 7
    L1D.Valid = False
    v = CacheReadL1D(8)                  ' tag 1: L1D miss, L2 HIT, refreshes LastUse of line 1
    If Stats.L2Hit <> 2 Then fails = fails + 1: Debug.Print "  FAIL  5: expected 2 L2 hits"
    Clock = 8
    v = CacheReadL1D(40)                 ' tag 5: L2 miss
    If L2(2).tag <> 5 Then fails = fails + 1: Debug.Print "  FAIL  5: LRU must replace L2 line 2 (tag 5), got tag " & L2(2).tag
    If L2(1).tag <> 1 Then fails = fails + 1: Debug.Print "  FAIL  5: the recently used line 1 must be kept"

    ' 6) STORE hit in L2: updates the L2 line and RAM (write-through), counts L2 hit
    Clock = 9
    L1D.Valid = False                    ' L1D has no matching line: L1D miss, but tag 1 is in L2
    CacheWriteL1D 9, 222
    If L2(1).Data(1) <> 222 Then fails = fails + 1: Debug.Print "  FAIL  6: the L2 line must hold 222 at byte 1"
    If RAM(9) <> 222 Then fails = fails + 1: Debug.Print "  FAIL  6: RAM(9) must be 222"
    If Stats.L2Hit <> 3 Then fails = fails + 1: Debug.Print "  FAIL  6: a STORE hit in L2 must count (3 L2 hits)"

    ' 7) STORE miss in L2: the line is NOT loaded, RAM is written, counts L2 miss
    Dim l2Miss As Long
    l2Miss = Stats.l2Miss
    CacheWriteL1D 56, 9                  ' tag 7 is not in L2
    If L2Find(7) <> -1 Then fails = fails + 1: Debug.Print "  FAIL  7: no-write-allocate: tag 7 must NOT enter L2"
    If RAM(56) <> 9 Then fails = fails + 1: Debug.Print "  FAIL  7: RAM(56) must be 9"
    If Stats.l2Miss <> l2Miss + 1 Then fails = fails + 1: Debug.Print "  FAIL  7: a STORE miss in L2 must count as an L2 miss"

    ' 8) InvalidateFrame: frame 0 covers tags 0..3. Valid lines in that range die, others stay
    InvalidateFrame 0
    If L2(1).Valid Or L2(3).Valid Then fails = fails + 1: Debug.Print "  FAIL  8: L2 lines with tags 1 and 3 must be invalid"
    If Not L2(0).Valid Or Not L2(2).Valid Then fails = fails + 1: Debug.Print "  FAIL  8: lines with tags 4 and 5 must stay valid"
    L1I.Valid = True: L1I.tag = 2
    L1D.Valid = True: L1D.tag = 3
    InvalidateFrame 0
    If L1I.Valid Or L1D.Valid Then fails = fails + 1: Debug.Print "  FAIL  8: L1I (tag 2) and L1D (tag 3) must be invalid"
    L1I.Valid = True: L1I.tag = 4
    InvalidateFrame 0
    If Not L1I.Valid Then fails = fails + 1: Debug.Print "  FAIL  8: L1I with tag 4 (frame 1) must stay valid"

    ' 9) InvalidateAllCaches
    InvalidateAllCaches
    If L1I.Valid Or L1D.Valid Or L2(0).Valid Or L2(1).Valid Or L2(2).Valid Or L2(3).Valid Then fails = fails + 1: Debug.Print "  FAIL  9: everything must be invalid"

    ' Cleanup
    For i = 0 To 63: RAM(i) = 0: Next i
    ResetCachesForTest
    Clock = 0
    ShowL1ILine "-"
    ShowL1DLine "-"
    ShowAllL2Rows "-"

    Debug.Print IIf(fails = 0, "Test_CacheL2: ALL OK", "Test_CacheL2: " & fails & " FAIL(S)")
End Sub

' For the PR screenshot: fills the 4 L2 lines and then forces an LRU replacement.
' Afterwards run Test_CacheL2 to leave everything clean.
Public Sub Demo_CacheL2()
    Dim i As Integer, v As Integer, before As Long
    For i = 0 To 63: RAM(i) = &H40 + i: Next i
    ResetCachesForTest
    Debug.Print "Demo_CacheL2: 4 lines, fully associative, LRU"
    For i = 0 To 4
        Clock = i + 1
        before = Stats.L2Hit
        v = CacheReadL1D(i * LINE_SIZE)
        Debug.Print "  LOAD paddr " & (i * LINE_SIZE) & " (tag " & i & ") -> L2 " & IIf(Stats.L2Hit > before, "HIT ", "MISS") & "   data = " & Hex2(v)
    Next i
    Debug.Print "  Tag 4 replaced the least recently used line: L2 line 0 now holds tag " & L2(0).tag
    Clock = 6
    before = Stats.L2Hit
    L1D.Valid = False
    v = CacheReadL1D(8)
    Debug.Print "  LOAD paddr 8 (tag 1) -> L2 " & IIf(Stats.L2Hit > before, "HIT ", "MISS") & "   (L1D was invalid)"
    Debug.Print "  L2: hits = " & Stats.L2Hit & ", misses = " & Stats.l2Miss
End Sub

' For the PR screenshot of the previous levels (kept from K7)
Public Sub Demo_CacheL1D()
    Dim i As Integer, v As Integer, before As Long
    For i = 0 To 31: RAM(i) = &H20 + i: Next i
    ResetCachesForTest
    Debug.Print "Demo_CacheL1D: write-through + no-write-allocate"

    before = Stats.L1DHit
    v = CacheReadL1D(0)
    Debug.Print "  LOAD  paddr 0  -> " & IIf(Stats.L1DHit > before, "HIT ", "MISS") & "   data = " & Hex2(v)

    before = Stats.L1DHit
    v = CacheReadL1D(1)
    Debug.Print "  LOAD  paddr 1  -> " & IIf(Stats.L1DHit > before, "HIT ", "MISS") & "   data = " & Hex2(v)

    before = Stats.L1DHit
    CacheWriteL1D 2, &HAA
    Debug.Print "  STORE paddr 2  -> " & IIf(Stats.L1DHit > before, "HIT ", "MISS") & "   RAM = " & Hex2(RAM(2)) & ", line B2 = " & Hex2(L1D.Data(2))

    before = Stats.L1DHit
    CacheWriteL1D 12, &HBB
    Debug.Print "  STORE paddr 12 -> " & IIf(Stats.L1DHit > before, "HIT ", "MISS") & "   RAM = " & Hex2(RAM(12)) & ", line NOT loaded (tag still " & L1D.tag & ")"

    before = Stats.L1DHit
    v = CacheReadL1D(12)
    Debug.Print "  LOAD  paddr 12 -> " & IIf(Stats.L1DHit > before, "HIT ", "MISS") & "   data = " & Hex2(v)

    Debug.Print "  L1D: hits = " & Stats.L1DHit & ", misses = " & Stats.L1DMiss
End Sub

' For the PR screenshot of L1I (kept from K6)
Public Sub Demo_CacheL1I()
    Dim i As Integer, hitsBefore As Long, v As Integer
    For i = 0 To 15: RAM(i) = &H10 + i: Next i
    ResetCachesForTest
    Debug.Print "Demo_CacheL1I: 8 consecutive bytes from physical address 0, then byte 8"
    For i = 0 To 8
        hitsBefore = Stats.L1IHit
        v = CacheReadL1I(i)
        Debug.Print "  paddr " & i & " -> " & IIf(Stats.L1IHit > hitsBefore, "HIT ", "MISS") & "   data = " & Hex2(v)
    Next i
    Debug.Print "  L1I: hits = " & Stats.L1IHit & ", misses = " & Stats.L1IMiss
End Sub

