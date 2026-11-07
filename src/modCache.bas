Attribute VB_Name = "modCache"
Option Explicit
' =====================================================================
' modCache - Cache hierarchy (K6: L1I; K7: L1D; K8: L2)
'   L1I: 8 B (1 line), direct mapped, instruction fetch only.
'   L1D: 8 B (1 line), direct mapped, LOAD/STORE only.
'   Works with PHYSICAL addresses (after the MMU).
'   tag = paddr \ LINE_SIZE
'
' WRITE POLICY (L1D): write-through + no-write-allocate
'   - STORE hit : the line is updated AND RAM is written.
'   - STORE miss: the line is NOT brought into the cache (no-write-allocate),
'                 RAM is written anyway.
'   - RAM is ALWAYS up to date. Writes also count hit/miss in L1D.
'
' The I/O page (15) never gets here: MemRead/MemWrite bypass the cache for it.
' Until L2 exists (K8), a new line is read straight from RAM.
' =====================================================================

Private Const CACHE_HDR_ROW As Long = 19
Private Const L1I_ROW As Long = 20
Private Const L1D_ROW As Long = 21

Private Const SHP_L1I As String = "l1i_box"
Private Const SHP_L1D As String = "l1d_box"

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
'  - Both cases count hit/miss in L1D, and RAM is ALWAYS written.
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

' K7 PLACEHOLDER: there is no L2 yet.
' K8 completes it: update the L2 line on a hit, count Stats.L2Hit/L2Miss
' for the STORE and animate l2_lineN.
Private Sub WriteThroughToL2(ByVal paddr As Integer, ByVal value As Integer)
    ' Intentionally empty until K8.
End Sub

' ---------- Line fill ----------

' K6 PLACEHOLDER: there is no L2 yet, so the line is read straight from RAM.
' K8 replaces the body: look in L2 first and go down to RAM only on an L2 miss.
Private Sub FillFromL2(ByRef l1 As TCacheLine, ByVal tag As Integer)
    Dim i As Integer
    For i = 0 To LINE_SIZE - 1
        l1.Data(i) = RAM(tag * LINE_SIZE + i)
    Next i
    l1.Valid = True
    l1.tag = tag
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
    If rowNum = L1D_ROW Then RowLabel = "L1D" Else RowLabel = "L1I"
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

Private Sub ShowL1ILine(ByVal result As String)
    ShowLineAt L1I_ROW, L1I, result
End Sub

Private Sub ShowL1DLine(ByVal result As String)
    ShowLineAt L1D_ROW, L1D, result
End Sub

' ---------- TESTS ----------
' Run with F5 (cursor inside the Sub) and check the Immediate window (Ctrl+G)
Public Sub Test_CacheL1I()
    Dim fails As Integer, i As Integer, v As Integer, ok As Boolean
    Dim ws As Worksheet, emptyStats As TStats, emptyLine As TCacheLine
    Set ws = GetCacheSheet()

    For i = 0 To 31: RAM(i) = 100 + i: Next i
    L1I = emptyLine
    Stats = emptyStats

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

    ' 5) A MISS reads again
    v = CacheReadL1I(0)
    If v <> 100 Or Stats.L1IMiss <> 3 Then fails = fails + 1: Debug.Print "  FAIL  5: going back to line 0 must be a MISS and return 100"
    v = CacheReadL1I(10)
    If v <> 250 Or Stats.L1IMiss <> 4 Then fails = fails + 1: Debug.Print "  FAIL  5: the refreshed line 1 must return 250 (MISS)"
    v = CacheReadL1I(10)
    If v <> 250 Or Stats.L1IHit <> 9 Then fails = fails + 1: Debug.Print "  FAIL  5: the second access must be a HIT returning 250"

    ' 6) Invalid line: even with a matching tag it is a MISS
    L1I.Valid = False
    v = CacheReadL1I(10)
    If v <> 250 Or Stats.L1IMiss <> 5 Or Stats.L1IHit <> 9 Then fails = fails + 1: Debug.Print "  FAIL  6: an invalid line must give a MISS (miss=" & Stats.L1IMiss & ", hit=" & Stats.L1IHit & ")"

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
    L1I = emptyLine
    Stats = emptyStats
    ShowL1ILine "-"
    If CStr(ws.Cells(L1I_ROW, 3).value) <> "-" Then fails = fails + 1: Debug.Print "  FAIL  cleanup: the L1I row was not reset to '-'"

    Debug.Print IIf(fails = 0, "Test_CacheL1I: ALL OK", "Test_CacheL1I: " & fails & " FAIL(S)")
End Sub

Public Sub Test_CacheL1D()
    Dim fails As Integer, i As Integer, v As Integer
    Dim ws As Worksheet, emptyStats As TStats, emptyLine As TCacheLine
    Set ws = GetCacheSheet()

    For i = 0 To 47: RAM(i) = 100 + i: Next i
    L1D = emptyLine
    L1I = emptyLine
    Stats = emptyStats

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

    ' 4) STORE miss: the line is NOT loaded, RAM is written
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
    L1D = emptyLine
    L1I = emptyLine
    Stats = emptyStats
    ShowL1DLine "-"
    If CStr(ws.Cells(L1D_ROW, 3).value) <> "-" Then fails = fails + 1: Debug.Print "  FAIL  cleanup: the L1D row was not reset to '-'"

    Debug.Print IIf(fails = 0, "Test_CacheL1D: ALL OK", "Test_CacheL1D: " & fails & " FAIL(S)")
End Sub

' For the PR screenshot: LOAD miss, LOAD hit, STORE hit, STORE miss, LOAD after the miss.
' Afterwards run Test_CacheL1D to leave everything clean.
Public Sub Demo_CacheL1D()
    Dim i As Integer, v As Integer, before As Long
    Dim emptyStats As TStats, emptyLine As TCacheLine
    For i = 0 To 31: RAM(i) = &H20 + i: Next i
    L1D = emptyLine
    Stats = emptyStats
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

