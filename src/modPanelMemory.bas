Attribute VB_Name = "modPanelMemory"
Option Explicit
' =====================================================================
' modPanelMemory - Memory panel drawn with shapes (K10)
' Piece 1: static layout. 8 frames (frm_0..7), L1I, L1D and the four L2
' lines, each with a label chip (lbl_<name>) where UI_CacheFlash writes
' HIT / MISS. The panel lives on the Memory sheet below row 28, so the
' existing tables (page table, swap grid, cache rows, hit rate bars)
' stay where they are.
' Piece 2: RefreshMemoryPanel. Shows which page sits in each frame (from
' FrameOwner) with its first bytes, and repaints the cache boxes grey.
' Piece 3: ShowTranslation. A strip (xlt_box, row 37) that shows
' vaddr = VPN | offset -> PFN | offset = paddr, or PAGE FAULT.
' Piece 4: UI_PageFaultFlash (page table row blinks red) and
' UI_AnimateSwapIn (a block travels from the swap grid to the frame).
' Piece 5: UI_AnimateLineFill (a block travels from an L2 line to L1I/L1D).
' Piece 6a: replacement policy selector (pol_fifo / pol_lru, row 40).
' Piece 6b: click on a page table row (A2:E17) -> its frame gets a thick
' outline and the translation strip shows that page (OnPageTableClick).
' =====================================================================

Private Const PANEL_TOP_ROW As Long = 28
Private Const XLT_ROW As Long = 37
Private Const XLT_NAME As String = "xlt_box"
Private Const POL_ROW As Long = 40
Private Const POL_FIFO_NAME As String = "pol_fifo"
Private Const POL_LRU_NAME As String = "pol_lru"
Private Const POL_W As Single = 120
Private Const POL_H As Single = 24
Private Const TOKEN_NAME As String = "tok_swap"
Private Const FILL_TOKEN_NAME As String = "tok_fill"
Private Const TOKEN_W As Single = 36
Private Const TOKEN_H As Single = 18
Private Const TOKEN_STEPS As Integer = 12
Private Const PT_FIRST_ROW As Long = 2      ' page table: VPN 0 is on row 2
Private Const SWAP_FIRST_ROW As Long = 2    ' swap grid: P0 is on row 2
Private Const SWAP_FIRST_COL As Long = 9    ' swap grid: first byte is column I
Private Const BOX_W As Single = 78
Private Const STEP_X As Single = 84
Private Const FRAME_H As Single = 50
Private Const CACHE_H As Single = 36
Private Const CHIP_H As Single = 16
Private Const PREVIEW_BYTES As Integer = 3
Private Const HL_WEIGHT As Single = 4       ' outline weight of a highlighted frame

' Piece 6b: which frame is highlighted (-1 = none) and its original outline
Private hlPfn As Integer
Private hlVisible As Long
Private hlWeight As Single
Private hlColor As Long
Private hlActive As Boolean

' Draws (or redraws) the whole panel. Safe to call many times:
' UI_AddBox deletes a shape with the same name before creating it.
Public Sub BuildMemoryPanel()
    Dim ws As Worksheet, i As Integer
    Dim leftX As Single, topY As Single, cacheY As Single
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    leftX = ws.Cells(1, 1).Left + 8
    topY = ws.Rows(PANEL_TOP_ROW).Top
    cacheY = topY + FRAME_H + 16

    ' The frames are recreated, so any old highlight is gone
    hlActive = False

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

    ' Translation strip under the caches
    UI_AddBox ws, XLT_NAME, leftX, ws.Rows(XLT_ROW).Top, NUM_FRAMES * STEP_X - 6, 36, _
        "Translation: waiting for a memory access"

    ' Replacement policy selector under the strip
    BuildPolicySelector ws, leftX
End Sub

' A cache box plus its label chip underneath (lbl_<name>)
Private Sub AddCacheBox(ByVal ws As Worksheet, ByVal nm As String, _
        ByVal l As Single, ByVal t As Single, ByVal txt As String)
    Dim chip As Shape
    UI_AddBox ws, nm, l, t, BOX_W, CACHE_H, txt
    Set chip = UI_AddBox(ws, "lbl_" & nm, l, t + CACHE_H + 2, BOX_W, CHIP_H, "")
    chip.TextFrame2.TextRange.Font.size = 9
End Sub

' ---------- Piece 6a: replacement policy selector ----------

' Two buttons. The one that matches ReplacePolicy is amber, the other grey.
Private Sub BuildPolicySelector(ByVal ws As Worksheet, ByVal leftX As Single)
    Dim t As Single, b As Shape
    t = ws.Rows(POL_ROW).Top
    Set b = UI_AddBox(ws, POL_FIFO_NAME, leftX, t, POL_W, POL_H, "Replace: FIFO")
    b.OnAction = "SetPolicyFIFO"
    Set b = UI_AddBox(ws, POL_LRU_NAME, leftX + POL_W + 8, t, POL_W, POL_H, "Replace: LRU")
    b.OnAction = "SetPolicyLRU"
    PaintPolicySelector ws
End Sub

Private Sub PaintPolicySelector(ByVal ws As Worksheet)
    If Not HasShape(ws, POL_FIFO_NAME) Then Exit Sub
    If Not HasShape(ws, POL_LRU_NAME) Then Exit Sub
    UI_SetFill POL_FIFO_NAME, IIf(ReplacePolicy = POLICY_FIFO, CLR_LOAD, CLR_HIGHZ)
    UI_SetFill POL_LRU_NAME, IIf(ReplacePolicy = POLICY_LRU, CLR_LOAD, CLR_HIGHZ)
End Sub

' Button handlers (assigned with OnAction)
Public Sub SetPolicyFIFO()
    ReplacePolicy = POLICY_FIFO
    PaintPolicySelector ThisWorkbook.Worksheets(SH_MEM)
End Sub

Public Sub SetPolicyLRU()
    ReplacePolicy = POLICY_LRU
    PaintPolicySelector ThisWorkbook.Worksheets(SH_MEM)
End Sub

' ---------- Piece 6b: click on a page table row ----------

' Called by Worksheet_SelectionChange of the Memory sheet.
' If the selection is one row inside the page table (A2:E17):
'  - the translation strip shows that page (offset 0), green or PAGE FAULT
'  - if the page is in RAM, its frame gets a thick outline
' Any other selection is ignored (the current highlight stays).
Public Sub OnPageTableClick(ByVal target As Range)
    Dim ws As Worksheet, vpn As Integer
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If target.Rows.Count <> 1 Then Exit Sub
    If Intersect(target, ws.Range("A2:E17")) Is Nothing Then Exit Sub

    vpn = target.row - PT_FIRST_ROW
    If vpn < 0 Or vpn >= NUM_VPAGES Then Exit Sub

    If Not HasShape(ws, PFX_FRM & "0") Then BuildMemoryPanel
    ClearFrameHighlight ws
    ShowTranslation vpn, 0
    If PageTable(vpn).Present Then HighlightFrame ws, PageTable(vpn).pfn
End Sub

' Thick outline on frame <pfn>. Remembers the original outline to restore it.
Private Sub HighlightFrame(ByVal ws As Worksheet, ByVal pfn As Integer)
    Dim shp As Shape
    If pfn < 0 Or pfn >= NUM_FRAMES Then Exit Sub
    If Not HasShape(ws, PFX_FRM & pfn) Then Exit Sub
    Set shp = ws.Shapes(PFX_FRM & pfn)
    hlVisible = shp.line.Visible
    hlWeight = shp.line.weight
    hlColor = shp.line.ForeColor.RGB
    shp.line.Visible = msoTrue
    shp.line.weight = HL_WEIGHT
    shp.line.ForeColor.RGB = CLR_TEXT
    hlPfn = pfn
    hlActive = True
End Sub

' Gives the highlighted frame its original outline back
Private Sub ClearFrameHighlight(ByVal ws As Worksheet)
    Dim shp As Shape
    If Not hlActive Then Exit Sub
    hlActive = False
    If Not HasShape(ws, PFX_FRM & hlPfn) Then Exit Sub
    Set shp = ws.Shapes(PFX_FRM & hlPfn)
    shp.line.weight = hlWeight
    shp.line.ForeColor.RGB = hlColor
    shp.line.Visible = hlVisible
End Sub

' Repaints the panel from the current machine state:
'  - frames 0..6: page that owns the frame (FrameOwner) + first bytes of RAM
'  - frame 7 keeps its fixed I/O text
'  - cache boxes go back to grey and their HIT/MISS chips are cleared
'  - the policy selector matches ReplacePolicy
Public Sub RefreshMemoryPanel()
    Dim ws As Worksheet, i As Integer, k As Integer
    Dim owner As Integer, base As Integer, preview As String
    Set ws = ThisWorkbook.Worksheets(SH_MEM)

    ' If the panel was never drawn, draw it first
    If Not HasShape(ws, PFX_FRM & "0") Then BuildMemoryPanel

    For i = 0 To NUM_FRAMES - 1
        If i <> IO_FRAME Then
            owner = FrameOwner(i)
            If owner < 0 Then
                UI_SetText PFX_FRM & i, "F" & i & vbLf & "free"
                UI_SetFill PFX_FRM & i, CLR_HIGHZ
            Else
                base = i * PAGE_SIZE
                preview = ""
                For k = 0 To PREVIEW_BYTES - 1
                    If k > 0 Then preview = preview & " "
                    preview = preview & Right$("0" & Hex$(RAM(base + k)), 2)
                Next k
                UI_SetText PFX_FRM & i, "F" & i & vbLf & "P" & owner & vbLf & preview
                UI_SetFill PFX_FRM & i, CLR_LOAD
            End If
        End If
    Next i

    ResetCacheBox "l1i_box", ws
    ResetCacheBox "l1d_box", ws
    For i = 0 To L2_LINES - 1
        ResetCacheBox "l2_line" & i, ws
    Next i

    PaintPolicySelector ws
End Sub

' One cache box back to grey, its label chip empty
Private Sub ResetCacheBox(ByVal nm As String, ByVal ws As Worksheet)
    If HasShape(ws, nm) Then UI_SetFill nm, CLR_HIGHZ
    If HasShape(ws, "lbl_" & nm) Then UI_SetText "lbl_" & nm, ""
End Sub

' Shows the address translation of the access in progress.
' Present page : "vaddr 0x0A5 = VPN 5 | off 5  ->  PFN 2  ->  paddr 0x045 = PFN 2 | off 5"
' Missing page : red strip with PAGE FAULT.
Public Sub ShowTranslation(ByVal vpn As Integer, ByVal offset As Integer)
    Dim ws As Worksheet, vaddr As Integer, paddr As Integer, txt As String
    Set ws = ThisWorkbook.Worksheets(SH_MEM)

    ' If the strip was never drawn, draw the panel first
    If Not HasShape(ws, XLT_NAME) Then BuildMemoryPanel

    vaddr = vpn * PAGE_SIZE + offset
    txt = "vaddr 0x" & Right$("000" & Hex$(vaddr), 3) & " = " & _
          "VPN " & vpn & " (" & ToBin(vpn, 4) & ") | off " & offset & " (" & ToBin(offset, OFFSET_BITS) & ")"

    If PageTable(vpn).Present Then
        paddr = PageTable(vpn).pfn * PAGE_SIZE + offset
        txt = txt & vbLf & "-> PFN " & PageTable(vpn).pfn & " | off " & offset & _
              "  =  paddr 0x" & Right$("00" & Hex$(paddr), 2)
        UI_SetFill XLT_NAME, CLR_HIT
    Else
        txt = txt & vbLf & "-> PAGE FAULT (P" & vpn & " is not in RAM)"
        UI_SetFill XLT_NAME, CLR_FAULT
    End If
    UI_SetText XLT_NAME, txt
End Sub

' Binary text of v with a fixed number of bits
Private Function ToBin(ByVal v As Integer, ByVal bits As Integer) As String
    Dim i As Integer, s As String
    For i = bits - 1 To 0 Step -1
        s = s & CStr((v \ (2 ^ i)) Mod 2)
    Next i
    ToBin = s
End Function

' The page table row of vpn blinks dark red twice, then gets its colors back.
' Skipped when animations are off (AnimDelay <= 0).
Public Sub UI_PageFaultFlash(ByVal vpn As Integer)
    Dim ws As Worksheet, c As Integer, n As Integer, r As Long
    Dim oldPat(1 To 5) As Long, oldClr(1 To 5) As Long, oldFnt(1 To 5) As Long
    If AnimDelay <= 0 Then Exit Sub
    If vpn < 0 Or vpn >= NUM_VPAGES Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    r = PT_FIRST_ROW + vpn

    ' Remember the original look of the 5 cells (A..E)
    For c = 1 To 5
        oldPat(c) = ws.Cells(r, c).Interior.Pattern
        oldClr(c) = ws.Cells(r, c).Interior.Color
        oldFnt(c) = ws.Cells(r, c).Font.Color
    Next c

    For n = 1 To 2
        For c = 1 To 5
            ws.Cells(r, c).Interior.Color = CLR_FAULT
            ws.Cells(r, c).Font.Color = CLR_TEXT
        Next c
        UI_Wait 150
        For c = 1 To 5
            RestoreCell ws.Cells(r, c), oldPat(c), oldClr(c), oldFnt(c)
        Next c
        UI_Wait 100
    Next n
End Sub

Private Sub RestoreCell(ByVal cel As Range, ByVal pat As Long, ByVal clr As Long, ByVal fnt As Long)
    If pat = xlNone Then
        cel.Interior.Pattern = xlNone
    Else
        cel.Interior.Color = clr
    End If
    cel.Font.Color = fnt
End Sub

' A block "P<vpn>" travels from its row in the swap grid to frame <pfn>.
' When it arrives the frame turns coral and shows the page. The bytes
' are filled in later by RefreshMemoryPanel.
' Skipped when animations are off (AnimDelay <= 0).
Public Sub UI_AnimateSwapIn(ByVal vpn As Integer, ByVal pfn As Integer)
    Dim ws As Worksheet, tok As Shape, frm As Shape, k As Integer
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
    If AnimDelay <= 0 Then Exit Sub
    If vpn < 0 Or vpn >= NUM_VPAGES Then Exit Sub
    If pfn < 0 Or pfn >= NUM_FRAMES Then Exit Sub
    If pfn = IO_FRAME Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If Not HasShape(ws, PFX_FRM & pfn) Then BuildMemoryPanel
    Set frm = ws.Shapes(PFX_FRM & pfn)

    x0 = ws.Cells(SWAP_FIRST_ROW + vpn, SWAP_FIRST_COL).Left
    y0 = ws.Cells(SWAP_FIRST_ROW + vpn, SWAP_FIRST_COL).Top
    x1 = frm.Left + (frm.Width - TOKEN_W) / 2
    y1 = frm.Top + (frm.Height - TOKEN_H) / 2

    Set tok = UI_AddBox(ws, TOKEN_NAME, x0, y0, TOKEN_W, TOKEN_H, "P" & vpn, CLR_ACTIVE)
    For k = 1 To TOKEN_STEPS
        tok.Left = x0 + (x1 - x0) * k / TOKEN_STEPS
        tok.Top = y0 + (y1 - y0) * k / TOKEN_STEPS
        DoEvents
        UI_Wait 20
    Next k
    tok.Delete

    UI_SetText PFX_FRM & pfn, "F" & pfn & vbLf & "P" & vpn
    UI_SetFill PFX_FRM & pfn, CLR_ACTIVE
End Sub

' A block "line" travels from L2 line <idx> to the L1 box that receives it
' (target = "l1i_box" or "l1d_box"). Called by FillFromL2 in modCache.
' The boxes keep the colors that FlashCache gave them; only the token moves.
' Skipped when animations are off (AnimDelay <= 0) or a shape is missing.
Public Sub UI_AnimateLineFill(ByVal idx As Integer, ByVal target As String)
    Dim ws As Worksheet, tok As Shape, src As Shape, dst As Shape, k As Integer
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
    If AnimDelay <= 0 Then Exit Sub
    If idx < 0 Or idx >= L2_LINES Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If Not HasShape(ws, "l2_line" & idx) Then Exit Sub
    If Not HasShape(ws, target) Then Exit Sub
    Set src = ws.Shapes("l2_line" & idx)
    Set dst = ws.Shapes(target)

    x0 = src.Left + (src.Width - TOKEN_W) / 2
    y0 = src.Top + (src.Height - TOKEN_H) / 2
    x1 = dst.Left + (dst.Width - TOKEN_W) / 2
    y1 = dst.Top + (dst.Height - TOKEN_H) / 2

    Set tok = UI_AddBox(ws, FILL_TOKEN_NAME, x0, y0, TOKEN_W, TOKEN_H, "line", CLR_ACTIVE)
    For k = 1 To TOKEN_STEPS
        tok.Left = x0 + (x1 - x0) * k / TOKEN_STEPS
        tok.Top = y0 + (y1 - y0) * k / TOKEN_STEPS
        DoEvents
        UI_Wait 20
    Next k
    tok.Delete
End Sub

' Visual demo for screenshots: a fault on page 3 loading into frame 2.
' Does not change the machine state.
Public Sub Demo_PageFaultAnim()
    Dim oldDelay As Long
    oldDelay = AnimDelay
    If AnimDelay <= 0 Then AnimDelay = 300   ' animations on only while the demo runs
    BuildMemoryPanel
    ShowTranslation 3, 5
    UI_PageFaultFlash 3
    UI_AnimateSwapIn 3, 2
    AnimDelay = oldDelay
End Sub

' Visual demo: a LOAD (L1D miss, line goes L2 -> L1D) and a fetch of the
' same line (L1I miss, L2 hit, line goes L2 -> L1I). Reads only, RAM is not
' changed. Afterwards run Test_CacheL1I and then BuildMemoryPanel to clean.
Public Sub Demo_LineFillAnim()
    Dim oldDelay As Long, v As Integer
    oldDelay = AnimDelay
    If AnimDelay <= 0 Then AnimDelay = 300   ' animations on only while the demo runs
    BuildMemoryPanel
    InvalidateAllCaches
    ResetStats
    v = CacheReadL1D(0)     ' L2 miss, line comes up to L1D
    v = CacheReadL1I(1)     ' same line, L2 hit, line comes up to L1I
    AnimDelay = oldDelay
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
    If Not HasShape(ws, XLT_NAME) Then fails = fails + 1: Debug.Print "  FAIL  1: missing " & XLT_NAME
    If Not HasShape(ws, POL_FIFO_NAME) Or Not HasShape(ws, POL_LRU_NAME) Then fails = fails + 1: Debug.Print "  FAIL  1: missing policy selector"

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

Public Sub Test_RefreshMemoryPanel()
    Dim fails As Integer, ws As Worksheet
    Dim oldOwner As Integer, oldB0 As Integer, oldB1 As Integer, oldB2 As Integer
    Dim t As String
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' Save the state this test changes
    oldOwner = FrameOwner(0)
    oldB0 = RAM(0): oldB1 = RAM(1): oldB2 = RAM(2)

    ' 1) Frame 0 owned by page 3 with known bytes
    FrameOwner(0) = 3
    RAM(0) = &H1A: RAM(1) = &HFF: RAM(2) = &H5
    RefreshMemoryPanel
    t = ws.Shapes(PFX_FRM & "0").TextFrame2.TextRange.Text
    If InStr(t, "P3") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: frm_0 must show P3, got: " & Replace(t, vbLf, "/")
    If InStr(t, "1A FF 05") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: frm_0 must show bytes 1A FF 05, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(PFX_FRM & "0").fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  1: occupied frame must use CLR_LOAD"

    ' 2) Frame 0 free again
    FrameOwner(0) = -1
    RefreshMemoryPanel
    t = ws.Shapes(PFX_FRM & "0").TextFrame2.TextRange.Text
    If InStr(t, "free") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: frm_0 must say free, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(PFX_FRM & "0").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  2: free frame must be grey"

    ' 3) Frame 7 keeps its I/O text
    t = ws.Shapes(PFX_FRM & IO_FRAME).TextFrame2.TextRange.Text
    If InStr(t, "I/O") = 0 Then fails = fails + 1: Debug.Print "  FAIL  3: frm_7 must keep I/O text"

    ' 4) Cache boxes painted by an animation go back to grey and the chip is cleared
    UI_SetFill "l1i_box", CLR_MISS
    UI_SetText "lbl_l1i_box", "MISS"
    UI_SetFill "l2_line2", CLR_HIT
    UI_SetText "lbl_l2_line2", "HIT"
    RefreshMemoryPanel
    If ws.Shapes("l1i_box").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  4: l1i_box must be grey after refresh"
    If ws.Shapes("lbl_l1i_box").TextFrame2.TextRange.Text <> "" Then fails = fails + 1: Debug.Print "  FAIL  4: lbl_l1i_box must be empty after refresh"
    If ws.Shapes("l2_line2").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  4: l2_line2 must be grey after refresh"
    If ws.Shapes("lbl_l2_line2").TextFrame2.TextRange.Text <> "" Then fails = fails + 1: Debug.Print "  FAIL  4: lbl_l2_line2 must be empty after refresh"

    ' Cleanup: restore the state and leave the panel as the machine has it
    FrameOwner(0) = oldOwner
    RAM(0) = oldB0: RAM(1) = oldB1: RAM(2) = oldB2
    RefreshMemoryPanel

    Debug.Print IIf(fails = 0, "Test_RefreshMemoryPanel: ALL OK", "Test_RefreshMemoryPanel: " & fails & " FAIL(S)")
End Sub

Public Sub Test_ShowTranslation()
    Dim fails As Integer, ws As Worksheet, t As String
    Dim oldPte As TPageEntry
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' Save the page table entry this test changes
    oldPte = PageTable(5)

    ' 1) Page 5 present in frame 2: vaddr 0x0A5 (VPN 5, offset 5) -> paddr 0x45
    PageTable(5).Present = True
    PageTable(5).pfn = 2
    ShowTranslation 5, 5
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.Text
    If InStr(t, "0x0A5") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show vaddr 0x0A5, got: " & Replace(t, vbLf, "/")
    If InStr(t, "VPN 5") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show VPN 5, got: " & Replace(t, vbLf, "/")
    If InStr(t, "0101") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show VPN in binary 0101, got: " & Replace(t, vbLf, "/")
    If InStr(t, "PFN 2") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show PFN 2, got: " & Replace(t, vbLf, "/")
    If InStr(t, "paddr 0x45") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show paddr 0x45, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(XLT_NAME).fill.ForeColor.RGB <> CLR_HIT Then fails = fails + 1: Debug.Print "  FAIL  1: present page must be green"

    ' 2) Page 5 not present: PAGE FAULT in red
    PageTable(5).Present = False
    ShowTranslation 5, 5
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.Text
    If InStr(t, "PAGE FAULT") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: must say PAGE FAULT, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(XLT_NAME).fill.ForeColor.RGB <> CLR_FAULT Then fails = fails + 1: Debug.Print "  FAIL  2: fault must be CLR_FAULT"

    ' Cleanup: restore the entry and redraw the panel
    PageTable(5) = oldPte
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_ShowTranslation: ALL OK", "Test_ShowTranslation: " & fails & " FAIL(S)")
End Sub

Public Sub Test_PageFaultAnim()
    Dim fails As Integer, ws As Worksheet, t As String, c As Integer
    Dim before(1 To 5) As Long, beforePat(1 To 5) As Long
    Dim oldDelay As Long
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' The test needs the animations on (they are skipped when AnimDelay <= 0)
    oldDelay = AnimDelay
    If AnimDelay <= 0 Then AnimDelay = 1

    ' 1) The flash gives the page table row its original colors back
    For c = 1 To 5
        before(c) = ws.Cells(PT_FIRST_ROW + 3, c).Interior.Color
        beforePat(c) = ws.Cells(PT_FIRST_ROW + 3, c).Interior.Pattern
    Next c
    UI_PageFaultFlash 3
    For c = 1 To 5
        If ws.Cells(PT_FIRST_ROW + 3, c).Interior.Pattern <> beforePat(c) Then fails = fails + 1: Debug.Print "  FAIL  1: row of P3 changed its fill pattern in column " & c
        If beforePat(c) <> xlNone And ws.Cells(PT_FIRST_ROW + 3, c).Interior.Color <> before(c) Then fails = fails + 1: Debug.Print "  FAIL  1: row of P3 did not get its color back in column " & c
    Next c

    ' 2) The swap-in leaves frame 2 coral with "P5" and no token behind
    UI_AnimateSwapIn 5, 2
    t = ws.Shapes(PFX_FRM & "2").TextFrame2.TextRange.Text
    If InStr(t, "P5") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: frm_2 must show P5, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(PFX_FRM & "2").fill.ForeColor.RGB <> CLR_ACTIVE Then fails = fails + 1: Debug.Print "  FAIL  2: frm_2 must be coral (CLR_ACTIVE)"
    If HasShape(ws, TOKEN_NAME) Then fails = fails + 1: Debug.Print "  FAIL  2: the token must be deleted after the animation"

    ' 3) Frame 7 (I/O) is never touched
    t = ws.Shapes(PFX_FRM & IO_FRAME).TextFrame2.TextRange.Text
    UI_AnimateSwapIn 5, IO_FRAME
    If ws.Shapes(PFX_FRM & IO_FRAME).TextFrame2.TextRange.Text <> t Then fails = fails + 1: Debug.Print "  FAIL  3: frm_7 must not change"

    ' Cleanup
    AnimDelay = oldDelay
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_PageFaultAnim: ALL OK", "Test_PageFaultAnim: " & fails & " FAIL(S)")
End Sub

Public Sub Test_LineFillAnim()
    Dim fails As Integer, ws As Worksheet, v As Integer, i As Integer
    Dim oldDelay As Long, oldClr As Long
    Dim oldRam(0 To 15) As Integer
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' The test needs the animations on (they are skipped when AnimDelay <= 0)
    oldDelay = AnimDelay
    If AnimDelay <= 0 Then AnimDelay = 1
    For i = 0 To 15: oldRam(i) = RAM(i): Next i

    ' 1) Direct call: the token is deleted and the boxes keep their color
    oldClr = ws.Shapes("l1d_box").fill.ForeColor.RGB
    UI_AnimateLineFill 2, "l1d_box"
    If HasShape(ws, FILL_TOKEN_NAME) Then fails = fails + 1: Debug.Print "  FAIL  1: the token must be deleted after the animation"
    If ws.Shapes("l1d_box").fill.ForeColor.RGB <> oldClr Then fails = fails + 1: Debug.Print "  FAIL  1: the animation must not change the color of l1d_box"

    ' 2) Bad arguments are ignored without error
    UI_AnimateLineFill 9, "l1d_box"
    UI_AnimateLineFill 0, "no_such_box"
    If HasShape(ws, FILL_TOKEN_NAME) Then fails = fails + 1: Debug.Print "  FAIL  2: bad arguments must not leave a token"

    ' 3) Real traffic: a miss in L1D and another in L1I go through the animation
    '    (modCache passes the box name) and leave no token and correct data
    For i = 0 To 15: RAM(i) = 100 + i: Next i
    InvalidateAllCaches
    ResetStats
    v = CacheReadL1D(0)
    If v <> 100 Then fails = fails + 1: Debug.Print "  FAIL  3: CacheReadL1D(0) = " & v & " (expected 100)"
    v = CacheReadL1I(9)
    If v <> 109 Then fails = fails + 1: Debug.Print "  FAIL  3: CacheReadL1I(9) = " & v & " (expected 109)"
    If HasShape(ws, FILL_TOKEN_NAME) Then fails = fails + 1: Debug.Print "  FAIL  3: no token may be left after real reads"
    If Stats.L1DMiss <> 1 Or Stats.L1IMiss <> 1 Then fails = fails + 1: Debug.Print "  FAIL  3: expected 1 L1D miss and 1 L1I miss"

    ' 4) With the animations off nothing is created
    AnimDelay = 0
    UI_AnimateLineFill 1, "l1i_box"
    If HasShape(ws, FILL_TOKEN_NAME) Then fails = fails + 1: Debug.Print "  FAIL  4: with AnimDelay = 0 nothing may be drawn"

    ' Cleanup
    For i = 0 To 15: RAM(i) = oldRam(i): Next i
    InvalidateAllCaches
    ResetStats
    AnimDelay = oldDelay
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_LineFillAnim: ALL OK", "Test_LineFillAnim: " & fails & " FAIL(S)")
End Sub

Public Sub Test_PolicySelector()
    Dim fails As Integer, ws As Worksheet, oldPolicy As Integer
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    oldPolicy = ReplacePolicy
    BuildMemoryPanel

    ' 1) Both buttons exist and call the right macros
    If Not HasShape(ws, POL_FIFO_NAME) Or Not HasShape(ws, POL_LRU_NAME) Then
        fails = fails + 1: Debug.Print "  FAIL  1: missing pol_fifo or pol_lru"
    Else
        If InStr(ws.Shapes(POL_FIFO_NAME).OnAction, "SetPolicyFIFO") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: pol_fifo must call SetPolicyFIFO"
        If InStr(ws.Shapes(POL_LRU_NAME).OnAction, "SetPolicyLRU") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: pol_lru must call SetPolicyLRU"
    End If

    ' 2) LRU: the policy changes and the amber button moves
    SetPolicyLRU
    If ReplacePolicy <> POLICY_LRU Then fails = fails + 1: Debug.Print "  FAIL  2: ReplacePolicy must be POLICY_LRU"
    If ws.Shapes(POL_LRU_NAME).fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  2: pol_lru must be amber"
    If ws.Shapes(POL_FIFO_NAME).fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  2: pol_fifo must be grey"

    ' 3) FIFO: back again
    SetPolicyFIFO
    If ReplacePolicy <> POLICY_FIFO Then fails = fails + 1: Debug.Print "  FAIL  3: ReplacePolicy must be POLICY_FIFO"
    If ws.Shapes(POL_FIFO_NAME).fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  3: pol_fifo must be amber"
    If ws.Shapes(POL_LRU_NAME).fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  3: pol_lru must be grey"

    ' 4) A change made from outside (not by the buttons) shows up after a refresh
    ReplacePolicy = POLICY_LRU
    RefreshMemoryPanel
    If ws.Shapes(POL_LRU_NAME).fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  4: refresh must repaint pol_lru as amber"
    If ws.Shapes(POL_FIFO_NAME).fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  4: refresh must repaint pol_fifo as grey"

    ' Cleanup: the policy is the user's choice, so put it back as it was
    ReplacePolicy = oldPolicy
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_PolicySelector: ALL OK", "Test_PolicySelector: " & fails & " FAIL(S)")
End Sub

Public Sub Test_PageTableClick()
    Dim fails As Integer, ws As Worksheet, t As String
    Dim oldPte5 As TPageEntry, oldPte6 As TPageEntry, oldPte15 As TPageEntry
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' Save the page table entries this test changes
    oldPte5 = PageTable(5)
    oldPte6 = PageTable(6)
    oldPte15 = PageTable(15)

    ' 1) Page 5 present in frame 2: click on its row (row 7) outlines frm_2
    PageTable(5).Present = True
    PageTable(5).pfn = 2
    OnPageTableClick ws.Range("A7")
    If ws.Shapes(PFX_FRM & "2").line.weight <> HL_WEIGHT Then fails = fails + 1: Debug.Print "  FAIL  1: frm_2 must have the thick outline"
    If ws.Shapes(PFX_FRM & "3").line.weight = HL_WEIGHT Then fails = fails + 1: Debug.Print "  FAIL  1: frm_3 must not be outlined"
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.Text
    If InStr(t, "VPN 5") = 0 Or InStr(t, "PFN 2") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: strip must show VPN 5 -> PFN 2, got: " & Replace(t, vbLf, "/")

    ' 2) A click outside the page table does nothing
    OnPageTableClick ws.Range("H3")
    If ws.Shapes(PFX_FRM & "2").line.weight <> HL_WEIGHT Then fails = fails + 1: Debug.Print "  FAIL  2: a click outside A2:E17 must keep the outline"

    ' 3) Page 6 not present: the outline of frm_2 goes away and the strip says PAGE FAULT
    PageTable(6).Present = False
    OnPageTableClick ws.Range("C8")
    If ws.Shapes(PFX_FRM & "2").line.weight = HL_WEIGHT Then fails = fails + 1: Debug.Print "  FAIL  3: frm_2 must lose its outline"
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.Text
    If InStr(t, "PAGE FAULT") = 0 Then fails = fails + 1: Debug.Print "  FAIL  3: strip must say PAGE FAULT, got: " & Replace(t, vbLf, "/")

    ' 4) Page 15 (I/O) outlines frame 7. The test sets the entry itself,
    '    so it does not depend on the state of the machine.
    PageTable(15).Present = True
    PageTable(15).pfn = IO_FRAME
    OnPageTableClick ws.Range("A17")
    If ws.Shapes(PFX_FRM & IO_FRAME).line.weight <> HL_WEIGHT Then fails = fails + 1: Debug.Print "  FAIL  4: frm_7 must have the thick outline for P15"

    ' Cleanup: restore the entries and redraw the panel (this also clears the outline)
    PageTable(5) = oldPte5
    PageTable(6) = oldPte6
    PageTable(15) = oldPte15
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_PageTableClick: ALL OK", "Test_PageTableClick: " & fails & " FAIL(S)")
End Sub

