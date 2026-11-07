Attribute VB_Name = "modPanelMemory"
Option Explicit
' =====================================================================
' modPanelMemory - Memory panel drawn with shapes (K10)
' Piece 1: static layout. 8 frames (frm_0..7), L1I, L1D and the four L2
' lines, each with a label chip (lbl_<name>) where UI_CacheFlash writes
' HIT / MISS. The panel lives on the Memory sheet below row 28, so the
' existing tables (page table, swap grid, cache rows, hit rate bars)
' stay where they are.
' =====================================================================

Private Const PANEL_TOP_ROW As Long = 28
Private Const BOX_W As Single = 78
Private Const STEP_X As Single = 84
Private Const FRAME_H As Single = 50
Private Const CACHE_H As Single = 36
Private Const CHIP_H As Single = 16

' Draws (or redraws) the whole panel. Safe to call many times:
' UI_AddBox deletes a shape with the same name before creating it.
Public Sub BuildMemoryPanel()
    Dim ws As Worksheet, i As Integer
    Dim leftX As Single, topY As Single, cacheY As Single
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    leftX = ws.Cells(1, 1).Left + 8
    topY = ws.Rows(PANEL_TOP_ROW).Top
    cacheY = topY + FRAME_H + 16

    ' RAM frames 0..7 (frame 7 belongs to the I/O page, always pinned)
    For i = 0 To NUM_FRAMES - 1
        If i = IO_FRAME Then
            UI_AddBox ws, PFX_FRM & i, leftX + i * STEP_X, topY, BOX_W, FRAME_H, _
                "F" & i & vbLf & "I/O (P" & IO_VPAGE & ")"
        Else
            UI_AddBox ws, PFX_FRM & i, leftX + i * STEP_X, topY, BOX_W, FRAME_H, _
                "F" & i & vbLf & "free"
        End If
    Next i

    ' L1I and L1D
    AddCacheBox ws, "l1i_box", leftX, cacheY, "L1I"
    AddCacheBox ws, "l1d_box", leftX + STEP_X, cacheY, "L1D"

    ' L2: four lines, a small gap after L1D
    For i = 0 To L2_LINES - 1
        AddCacheBox ws, "l2_line" & i, leftX + (2 + i) * STEP_X + 16, cacheY, "L2-" & i
    Next i
End Sub

' A cache box plus its label chip underneath (lbl_<name>)
Private Sub AddCacheBox(ByVal ws As Worksheet, ByVal nm As String, _
        ByVal l As Single, ByVal t As Single, ByVal txt As String)
    Dim chip As Shape
    UI_AddBox ws, nm, l, t, BOX_W, CACHE_H, txt
    Set chip = UI_AddBox(ws, "lbl_" & nm, l, t + CACHE_H + 2, BOX_W, CHIP_H, "")
    chip.TextFrame2.TextRange.Font.size = 9
End Sub

' ---------- TESTS ----------

Private Function HasShape(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(nm)
    On Error GoTo 0
    HasShape = Not (s Is Nothing)
End Function

Public Sub Test_BuildMemoryPanel()
    Dim fails As Integer, i As Integer, ws As Worksheet, v As Integer
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' 1) Every shape exists
    For i = 0 To NUM_FRAMES - 1
        If Not HasShape(ws, PFX_FRM & i) Then fails = fails + 1: Debug.Print "  FAIL  1: missing shape " & PFX_FRM & i
    Next i
    If Not HasShape(ws, "l1i_box") Or Not HasShape(ws, "lbl_l1i_box") Then fails = fails + 1: Debug.Print "  FAIL  1: missing l1i_box or its label"
    If Not HasShape(ws, "l1d_box") Or Not HasShape(ws, "lbl_l1d_box") Then fails = fails + 1: Debug.Print "  FAIL  1: missing l1d_box or its label"
    For i = 0 To L2_LINES - 1
        If Not HasShape(ws, "l2_line" & i) Then fails = fails + 1: Debug.Print "  FAIL  1: missing l2_line" & i
        If Not HasShape(ws, "lbl_l2_line" & i) Then fails = fails + 1: Debug.Print "  FAIL  1: missing lbl_l2_line" & i
    Next i

    ' 2) Real traffic: the cache animation must paint the boxes
    InvalidateAllCaches
    ResetStats
    v = CacheReadL1I(0)                  ' L1I miss and L2 miss
    If ws.Shapes("l1i_box").fill.ForeColor.RGB <> CLR_MISS Then fails = fails + 1: Debug.Print "  FAIL  2: l1i_box must be red after a miss"
    If ws.Shapes("lbl_l1i_box").TextFrame2.TextRange.Text <> "MISS" Then fails = fails + 1: Debug.Print "  FAIL  2: lbl_l1i_box must say MISS"
    If ws.Shapes("l2_line0").fill.ForeColor.RGB <> CLR_MISS Then fails = fails + 1: Debug.Print "  FAIL  2: l2_line0 must be red (L2 miss)"
    v = CacheReadL1I(1)                  ' same line: L1I hit
    If ws.Shapes("l1i_box").fill.ForeColor.RGB <> CLR_HIT Then fails = fails + 1: Debug.Print "  FAIL  2: l1i_box must be green after a hit"
    If ws.Shapes("lbl_l1i_box").TextFrame2.TextRange.Text <> "HIT" Then fails = fails + 1: Debug.Print "  FAIL  2: lbl_l1i_box must say HIT"

    ' Cleanup: caches empty, counters at 0, panel redrawn in grey
    InvalidateAllCaches
    ResetStats
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_BuildMemoryPanel: ALL OK", "Test_BuildMemoryPanel: " & fails & " FAIL(S)")
End Sub

