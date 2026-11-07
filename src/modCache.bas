Attribute VB_Name = "modCache"
Option Explicit
' =====================================================================
' modCache - Cache hierarchy (K6: L1I; K7: L1D; K8: L2)
'   L1I: 8 B (1 line), direct mapped, instruction fetch only.
'   Works with PHYSICAL addresses (after the MMU).
'   tag = paddr \ LINE_SIZE
' The I/O page (15) never gets here: MemRead bypasses the cache for it.
' Until L2 exists (K8), a new line is read straight from RAM.
' =====================================================================

' Area of the Memory sheet that shows the cache state
' (below the page table and the swap grid; K10 will move it).
Private Const CACHE_HDR_ROW As Long = 19
Private Const L1I_ROW As Long = 20

' Shape of the K10 panel that is animated on every L1I access.
' If it does not exist yet, the animation is simply skipped.
Private Const SHP_L1I As String = "l1i_box"

' ---------- K6: L1I ----------

' Reads one instruction byte through L1I.
'  - HIT : same tag and valid line -> the cache answers, it does NOT go to L2/RAM.
'  - MISS: fetches the whole 8-byte line (4 instructions) and answers.
' An instruction takes 2 bytes, so 4 consecutive instructions are
' 8 reads: 1 miss + 7 hits.
Public Function CacheReadL1I(ByVal paddr As Integer) As Integer
    Dim tag As Integer
    CheckPaddr paddr
    tag = paddr \ LINE_SIZE
    If L1I.Valid And L1I.tag = tag Then
        Stats.L1IHit = Stats.L1IHit + 1
        FlashCache SHP_L1I, True
        ShowL1IHit
    Else
        Stats.L1IMiss = Stats.L1IMiss + 1
        FlashCache SHP_L1I, False
        FillFromL2 L1I, tag
        ShowL1ILine "MISS"
    End If
    L1I.LastUse = Clock
    CacheReadL1I = L1I.Data(paddr Mod LINE_SIZE)
End Function

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
    ws.Cells(L1I_ROW, 1).value = "L1I"
End Sub

' Redraws the whole L1I row: result, tag and the 8 bytes in hexadecimal
Private Sub ShowL1ILine(ByVal result As String)
    Dim ws As Worksheet, i As Integer
    Dim cellsRow(0 To 0, 0 To 9) As String
    Set ws = GetCacheSheet()
    EnsureCacheTable ws
    cellsRow(0, 0) = result
    If L1I.Valid Then
        cellsRow(0, 1) = CStr(L1I.tag)
        For i = 0 To LINE_SIZE - 1
            cellsRow(0, 2 + i) = Hex2(L1I.Data(i))
        Next i
    Else
        cellsRow(0, 1) = "-"
        For i = 0 To LINE_SIZE - 1
            cellsRow(0, 2 + i) = "--"
        Next i
    End If
    ws.Range(ws.Cells(L1I_ROW, 2), ws.Cells(L1I_ROW, 11)).value = cellsRow
    Select Case result
        Case "HIT": ws.Cells(L1I_ROW, 2).Interior.Color = CLR_HIT
        Case "MISS": ws.Cells(L1I_ROW, 2).Interior.Color = CLR_MISS
        Case Else: ws.Cells(L1I_ROW, 2).Interior.Pattern = xlNone
    End Select
End Sub

' On a hit the line does not change: only the result cell is updated
Private Sub ShowL1IHit()
    Dim ws As Worksheet
    Set ws = GetCacheSheet()
    If CStr(ws.Cells(CACHE_HDR_ROW, 1).value) <> "Cache" Then
        ShowL1ILine "HIT"
        Exit Sub
    End If
    ws.Cells(L1I_ROW, 2).value = "HIT"
    ws.Cells(L1I_ROW, 2).Interior.Color = CLR_HIT
End Sub

' ---------- TESTS ----------
' Run with F5 (cursor inside the Sub) and check the Immediate window (Ctrl+G)
Public Sub Test_CacheL1I()
    Dim fails As Integer, i As Integer, v As Integer, ok As Boolean
    Dim ws As Worksheet, emptyStats As TStats, emptyLine As TCacheLine
    Set ws = GetCacheSheet()

    ' Starting point: RAM(0..31) = 100+i, L1I invalid, counters at 0
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

    ' 2) The other 7 bytes of the line (4 instructions = 8 reads): 7 HITS
    For i = 1 To 7
        v = CacheReadL1I(i)
        If v <> 100 + i Then fails = fails + 1: Debug.Print "  FAIL  2: CacheReadL1I(" & i & ") = " & v & " (expected " & 100 + i & ")"
    Next i
    If Stats.L1IHit <> 7 Or Stats.L1IMiss <> 1 Then fails = fails + 1: Debug.Print "  FAIL  2: 8 consecutive reads must give 1 miss + 7 hits (" & Stats.L1IMiss & " miss, " & Stats.L1IHit & " hits)"
    If CStr(ws.Cells(L1I_ROW, 2).value) <> "HIT" Then fails = fails + 1: Debug.Print "  FAIL  2: the sheet should show HIT"

    ' 3) Line change: MISS, L1I has a single line (tag 1)
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

    ' 4) A HIT does not go down to RAM: RAM changes, the cache still returns the old data
    RAM(10) = 250
    v = CacheReadL1I(10)
    If v <> 110 Then fails = fails + 1: Debug.Print "  FAIL  4: a HIT must be answered by the cache (110), got " & v
    If Stats.L1IHit <> 8 Or Stats.L1IMiss <> 2 Then fails = fails + 1: Debug.Print "  FAIL  4: expected 8 hits and 2 misses"

    ' 5) A MISS does read again: the new line already has the updated data
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

    ' 7) Physical address out of range: clear error
    On Error Resume Next
    v = CacheReadL1I(256)
    If Err.Number <> vbObjectError + 550 Then fails = fails + 1: Debug.Print "  FAIL  7: CacheReadL1I(256) did not raise error 550"
    Err.Clear
    v = CacheReadL1I(-1)
    If Err.Number <> vbObjectError + 550 Then fails = fails + 1: Debug.Print "  FAIL  7: CacheReadL1I(-1) did not raise error 550"
    Err.Clear
    On Error GoTo 0

    ' Cleanup: RAM(0..31) at 0, L1I invalid, counters at 0 and the row set to "-"
    For i = 0 To 31: RAM(i) = 0: Next i
    L1I = emptyLine
    Stats = emptyStats
    ShowL1ILine "-"
    If CStr(ws.Cells(L1I_ROW, 3).value) <> "-" Then fails = fails + 1: Debug.Print "  FAIL  cleanup: the L1I row was not reset to '-'"

    Debug.Print IIf(fails = 0, "Test_CacheL1I: ALL OK", "Test_CacheL1I: " & fails & " FAIL(S)")
End Sub

' For the PR screenshot: 4 consecutive instructions (8 bytes) and then byte 8.
' Leaves the L1I row visible on the Memory sheet (A19:K20).
' Afterwards run Test_CacheL1I to leave everything clean.
Public Sub Demo_CacheL1I()
    Dim i As Integer, hitsBefore As Long, v As Integer
    Dim emptyStats As TStats, emptyLine As TCacheLine
    For i = 0 To 15: RAM(i) = &H10 + i: Next i
    L1I = emptyLine
    Stats = emptyStats
    Debug.Print "Demo_CacheL1I: 8 consecutive bytes from physical address 0, then byte 8"
    For i = 0 To 8
        hitsBefore = Stats.L1IHit
        v = CacheReadL1I(i)
        Debug.Print "  paddr " & i & " -> " & IIf(Stats.L1IHit > hitsBefore, "HIT ", "MISS") & "   data = " & Hex2(v)
    Next i
    Debug.Print "  L1I: hits = " & Stats.L1IHit & ", misses = " & Stats.L1IMiss
End Sub

