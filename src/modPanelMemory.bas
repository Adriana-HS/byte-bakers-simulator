Attribute VB_Name = "modPanelMemory"
Option Explicit
' =====================================================================
' modPanelMemory - Memory panel drawn with shapes (K10)
'
' CODE OF THE "Memory" SHEET (lives only in the .xlsm, not in the repo).
' When the Excel is integrated, copy these two handlers into the sheet
' object "Memory" (a sheet .cls is not re-attached by the module import):
'
'   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'       OnPageTableClick Target
'   End Sub
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       OnMemoryCellEdit Target
'   End Sub
'
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
' Piece 6c: HEX / BIN / DEC view of the bytes shown in the frames
' (fmt_hex / fmt_bin / fmt_dec, row 42).
' Piece 7: live RAM editor. A 4x8 grid of cells (A47:H50) shows the 32
' bytes of the frame of the clicked page. Typing a hex value in a cell
' writes RAM, marks the page dirty, invalidates the caches of that frame
' and refreshes the panel (OnMemoryCellEdit, called by Worksheet_Change).
' Piece 8: BuildMiniMemory. A summary of the 8 frames inside the MEMORY
' frame (frm_MEM) of the Simulator sheet (shapes mini_*). It is repainted
' by RefreshMemoryPanel, so RefreshAll keeps it up to date.
' =====================================================================

Private Const PANEL_TOP_ROW As Long = 28
Private Const XLT_ROW As Long = 37
Private Const XLT_NAME As String = "xlt_box"
Private Const POL_ROW As Long = 40
Private Const POL_FIFO_NAME As String = "pol_fifo"
Private Const POL_LRU_NAME As String = "pol_lru"
Private Const POL_W As Single = 120
Private Const POL_H As Single = 24
Private Const FMT_ROW As Long = 42
Private Const FMT_HEX_NAME As String = "fmt_hex"
Private Const FMT_BIN_NAME As String = "fmt_bin"
Private Const FMT_DEC_NAME As String = "fmt_dec"
Private Const FMT_W As Single = 76
Private Const FMT_H As Single = 24
Private Const FMT_HEX As Integer = 0
Private Const FMT_BIN As Integer = 1
Private Const FMT_DEC As Integer = 2
Private Const EDIT_TITLE_ROW As Long = 45
Private Const EDIT_LABEL_ROW As Long = 46
Private Const EDIT_FIRST_ROW As Long = 47   ' grid: bytes 0..7 on this row
Private Const EDIT_COLS As Long = 8         ' grid: 8 bytes per row, columns A..H
Private Const EDIT_RANGE As String = "A47:H50"
Private Const TOKEN_NAME As String = "tok_swap"
Private Const FILL_TOKEN_NAME As String = "tok_fill"
Private Const TOKEN_W As Single = 36
Private Const TOKEN_H As Single = 18
Private Const TOKEN_STEPS As Integer = 12
Private Const PT_FIRST_ROW As Long = 2      ' page table: VPN 0 is on row 2
Private Const SWAP_FIRST_ROW As Long = 2    ' swap grid: P0 is on row 2
Private Const SWAP_FIRST_COL As Long = 9    ' swap grid: first byte is column I
Private Const BOX_W As Single = 86
Private Const STEP_X As Single = 92
Private Const FRAME_H As Single = 54
Private Const CACHE_H As Single = 36
Private Const CHIP_H As Single = 16
Private Const PREVIEW_BYTES As Integer = 3
Private Const HL_WEIGHT As Single = 4       ' outline weight of a highlighted frame

' Piece 8: mini memory summary inside frm_MEM of the Simulator sheet.
' frm_MEM is at left 680, top 50, 640 x 344 (modDashboard).
Private Const MINI_PREFIX As String = "mini_"
Private Const MINI_TITLE_NAME As String = "mini_title"
Private Const MINI_INFO_NAME As String = "mini_info"
Private Const MINI_X As Single = 696        ' 16 pt inside the frame
Private Const MINI_TITLE_Y As Single = 66
Private Const MINI_TOP As Single = 92
Private Const MINI_COLS As Integer = 4
Private Const MINI_BOX_W As Single = 140
Private Const MINI_BOX_H As Single = 70
Private Const MINI_GAP As Single = 16
Private Const MINI_INFO_Y As Single = 260
Private Const MINI_INFO_H As Single = 36

' Piece 6b: which frame is highlighted (-1 = none) and its original outline
Private hlPfn As Integer
Private hlVisible As Long
Private hlWeight As Single
Private hlColor As Long
Private hlActive As Boolean

' Piece 6c: number format of the bytes in the frames (FMT_HEX by default)
Private viewFmt As Integer

' Frame shown in the RAM editor. 0 = none, otherwise pfn + 1
' (so the default value 0 of a fresh module means "no frame").
Private editSel As Integer

' Last accepted RAM-editor write. Painted coral again after every grid refill.
' Replaced by the next accepted edit. Dropped when the grid is cleared.
Private lastEditOn As Boolean
Private lastEditPfn As Integer
Private lastEditIdx As Integer

' Draws (or redraws) the whole panel. Safe to call many times:
' UI_AddBox deletes a shape with the same name before creating it.
Public Sub BuildMemoryPanel()
    Dim ws As Worksheet, i As Integer, frm As Shape
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
            Set frm = UI_AddBox(ws, PFX_FRM & i, leftX + i * STEP_X, topY, BOX_W, FRAME_H, _
                "F" & i & vbLf & "I/O (P" & IO_VPAGE & ")")
        Else
            Set frm = UI_AddBox(ws, PFX_FRM & i, leftX + i * STEP_X, topY, BOX_W, FRAME_H, _
                "F" & i & vbLf & "free")
        End If
        With frm.TextFrame2
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 2: .MarginBottom = 2
            .WordWrap = msoFalse
        End With
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

    ' Number format selector under the policy selector
    BuildFormatSelector ws, leftX

    ' RAM editor grid: no frame selected yet
    editSel = 0
    RefreshEditGrid ws
    RefreshPageTableView
End Sub

' A cache box plus its label chip underneath (lbl_<name>)
Private Sub AddCacheBox(ByVal ws As Worksheet, ByVal nm As String, _
        ByVal l As Single, ByVal t As Single, ByVal txt As String)
    Dim chip As Shape
    UI_AddBox ws, nm, l, t, BOX_W, CACHE_H, txt
    Set chip = UI_AddBox(ws, "lbl_" & nm, l, t + CACHE_H + 2, BOX_W, CHIP_H, "")
    chip.TextFrame2.TextRange.Font.size = 9
End Sub

' ---------- Piece 8: mini memory summary on the Simulator sheet ----------

' Called by BuildDashboard (modDashboard) with the Simulator sheet.
' Draws a title, the 8 frames in 2 rows of 4 and an info line (policy and
' page faults) inside frm_MEM. The shapes are named mini_* so they never
' clash with frm_0..7 of the Memory sheet (UI_SetFill looks in Simulator
' first). Frame 7 is fixed text, the others are filled by RefreshMiniMemory.
Public Sub BuildMiniMemory(ByVal ws As Worksheet)
    Dim i As Integer, l As Single, t As Single, s As Shape

    Set s = UI_AddBox(ws, MINI_TITLE_NAME, MINI_X, MINI_TITLE_Y, 608, 20, _
        "Physical memory: 8 frames x 32 B (full panel in the Memory sheet)")
    s.fill.Visible = msoFalse
    s.line.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 0
        .WordWrap = msoFalse
        .TextRange.Font.fill.ForeColor.RGB = CLR_TEXT
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    For i = 0 To NUM_FRAMES - 1
        l = MINI_X + (i Mod MINI_COLS) * (MINI_BOX_W + MINI_GAP)
        t = MINI_TOP + (i \ MINI_COLS) * (MINI_BOX_H + 12)
        If i = IO_FRAME Then
            UI_AddBox ws, MINI_PREFIX & i, l, t, MINI_BOX_W, MINI_BOX_H, _
                "F" & i & vbLf & "I/O (P" & IO_VPAGE & ")"
        Else
            UI_AddBox ws, MINI_PREFIX & i, l, t, MINI_BOX_W, MINI_BOX_H, _
                "F" & i & vbLf & "free"
        End If
    Next i

    UI_AddBox ws, MINI_INFO_NAME, MINI_X, MINI_INFO_Y, 608, MINI_INFO_H, ""
    RefreshMiniMemory
End Sub

' Repaints the mini summary from the machine state. Does nothing if the
' Simulator sheet or the mini shapes do not exist yet.
Private Sub RefreshMiniMemory()
    Dim ws As Worksheet, i As Integer, k As Integer
    Dim owner As Integer, base As Integer, preview As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If Not HasShape(ws, MINI_PREFIX & "0") Then Exit Sub

    For i = 0 To NUM_FRAMES - 1
        If i <> IO_FRAME Then
            owner = FrameOwner(i)
            If owner < 0 Then
                UI_SetText MINI_PREFIX & i, "F" & i & vbLf & "free"
                UI_SetFill MINI_PREFIX & i, CLR_HIGHZ
            Else
                base = i * PAGE_SIZE
                preview = ""
                For k = 0 To BytesToShow() - 1
                    If k > 0 Then preview = preview & " "
                    preview = preview & FormatByte(RAM(base + k))
                Next k
                UI_SetText MINI_PREFIX & i, "F" & i & "  P" & owner & vbLf & preview
                UI_SetFill MINI_PREFIX & i, CLR_LOAD
            End If
        End If
    Next i

    If HasShape(ws, MINI_INFO_NAME) Then
        UI_SetText MINI_INFO_NAME, "Policy: " & IIf(ReplacePolicy = POLICY_LRU, "LRU", "FIFO") & _
            "  |  Page faults: " & Stats.PageFaults
    End If
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

' ---------- Piece 6c: HEX / BIN / DEC view ----------

' Three buttons. The one that matches viewFmt is amber, the others grey.
Private Sub BuildFormatSelector(ByVal ws As Worksheet, ByVal leftX As Single)
    Dim t As Single, b As Shape
    t = ws.Rows(FMT_ROW).Top
    Set b = UI_AddBox(ws, FMT_HEX_NAME, leftX, t, FMT_W, FMT_H, "HEX")
    b.OnAction = "SetViewHex"
    Set b = UI_AddBox(ws, FMT_BIN_NAME, leftX + FMT_W + 8, t, FMT_W, FMT_H, "BIN")
    b.OnAction = "SetViewBin"
    Set b = UI_AddBox(ws, FMT_DEC_NAME, leftX + 2 * (FMT_W + 8), t, FMT_W, FMT_H, "DEC")
    b.OnAction = "SetViewDec"
    PaintFormatSelector ws
End Sub

Private Sub PaintFormatSelector(ByVal ws As Worksheet)
    If Not HasShape(ws, FMT_HEX_NAME) Then Exit Sub
    If Not HasShape(ws, FMT_BIN_NAME) Then Exit Sub
    If Not HasShape(ws, FMT_DEC_NAME) Then Exit Sub
    UI_SetFill FMT_HEX_NAME, IIf(viewFmt = FMT_HEX, CLR_LOAD, CLR_HIGHZ)
    UI_SetFill FMT_BIN_NAME, IIf(viewFmt = FMT_BIN, CLR_LOAD, CLR_HIGHZ)
    UI_SetFill FMT_DEC_NAME, IIf(viewFmt = FMT_DEC, CLR_LOAD, CLR_HIGHZ)
End Sub

' Button handlers (assigned with OnAction). Each one repaints the frames.
Public Sub SetViewHex()
    viewFmt = FMT_HEX
    RefreshMemoryPanel
End Sub

Public Sub SetViewBin()
    viewFmt = FMT_BIN
    RefreshMemoryPanel
End Sub

Public Sub SetViewDec()
    viewFmt = FMT_DEC
    RefreshMemoryPanel
End Sub

' One byte as text in the current view: HEX "1A", BIN "00011010", DEC "26"
Private Function FormatByte(ByVal v As Integer) As String
    Select Case viewFmt
        Case FMT_BIN
            FormatByte = ToBin(v And BYTE_MAX, DATA_BITS)
        Case FMT_DEC
            FormatByte = CStr(v And BYTE_MAX)
        Case Else
            FormatByte = Right$("0" & Hex$(v And BYTE_MAX), 2)
    End Select
End Function

' How many bytes fit in a frame for the current view (BIN and DEC are wider)
Private Function BytesToShow() As Integer
    If viewFmt = FMT_HEX Then
        BytesToShow = PREVIEW_BYTES
    Else
        BytesToShow = 2
    End If
End Function

' ---------- Piece 6b: click on a page table row ----------

' Called by Worksheet_SelectionChange of the Memory sheet.
' If the selection is one row inside the page table (A2:E17):
'  - the translation strip shows that page (offset 0), green or PAGE FAULT
'  - if the page is in RAM, its frame gets a thick outline and its 32
'    bytes are loaded in the RAM editor grid (piece 7)
'  - if the page is not in RAM, the editor grid is cleared
' Any other selection is ignored (the current highlight stays).
Public Sub OnPageTableClick(ByVal Target As Range)
    Dim ws As Worksheet, vpn As Integer
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If Target.Rows.Count <> 1 Then Exit Sub
    If Intersect(Target, ws.Range("A2:E17")) Is Nothing Then Exit Sub

    vpn = Target.row - PT_FIRST_ROW
    If vpn < 0 Or vpn >= NUM_VPAGES Then Exit Sub

    If Not HasShape(ws, PFX_FRM & "0") Then BuildMemoryPanel
    ClearFrameHighlight ws
    ShowTranslation vpn, 0
    If PageTable(vpn).Present Then
        HighlightFrame ws, PageTable(vpn).pfn
        editSel = PageTable(vpn).pfn + 1
    Else
        editSel = 0
    End If
    RefreshEditGrid ws
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

' ---------- Piece 7: live RAM editor ----------

' Fills (or clears) the grid A47:H50 with the 32 bytes of the selected
' frame, always in hex, two digits per cell. Writes with events off so
' Worksheet_Change does not fire. The last accepted edit is painted coral
' again here, because this refill replaces the previous cell fill.
Private Sub RefreshEditGrid(ByVal ws As Worksheet)
    Dim oldEv As Boolean, idx As Integer, base As Integer, owner As Integer
    oldEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fail

    ws.Cells(EDIT_TITLE_ROW, 1).value = "RAM editor (hex, 2 digits per cell)"
    ws.Cells(EDIT_TITLE_ROW, 1).Font.bold = True
    With ws.Range(EDIT_RANGE)
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .Font.Name = "Consolas"
        .Borders.LineStyle = xlContinuous
    End With
    ClearEditorChrome ws

    If editSel = 0 Then
        lastEditOn = False
        ws.Range(EDIT_RANGE).ClearContents
        ws.Cells(EDIT_LABEL_ROW, 1).value = "Click a page in the table (A2:E17) to load its frame here"
    Else
        owner = FrameOwner(editSel - 1)
        If owner < 0 Then
            ws.Range(EDIT_RANGE).ClearContents
            ws.Cells(EDIT_LABEL_ROW, 1).value = "Frame F" & (editSel - 1) & " is free"
        Else
            base = (editSel - 1) * PAGE_SIZE
            ws.Cells(EDIT_LABEL_ROW, 1).value = "Frame F" & (editSel - 1) & " (P" & owner & "), bytes 0..31. Type a hex value to change RAM"
            For idx = 0 To PAGE_SIZE - 1
                ws.Cells(EDIT_FIRST_ROW + idx \ EDIT_COLS, 1 + idx Mod EDIT_COLS).value = _
                    Right$("0" & Hex$(RAM(base + idx) And BYTE_MAX), 2)
            Next idx
        End If
    End If
    PaintLastEditedByte ws
    Application.EnableEvents = oldEv
    Exit Sub
fail:
    Application.EnableEvents = oldEv
    Debug.Print "RefreshEditGrid error: " & Err.description
End Sub

' Drop any previous coral fill on A47:H50. Does not change the border.
Private Sub ClearEditorChrome(ByVal ws As Worksheet)
    Dim cel As Range
    For Each cel In ws.Range(EDIT_RANGE).Cells
        cel.Interior.Pattern = xlNone
        cel.Font.ColorIndex = xlAutomatic
    Next cel
End Sub

' Coral on the last accepted byte, only while that frame is still in the grid.
Private Sub PaintLastEditedByte(ByVal ws As Worksheet)
    Dim cel As Range
    Dim pfn As Integer
    If Not lastEditOn Then Exit Sub
    If editSel = 0 Then Exit Sub
    pfn = editSel - 1
    If pfn <> lastEditPfn Then Exit Sub
    If FrameOwner(pfn) < 0 Then Exit Sub
    If lastEditIdx < 0 Or lastEditIdx >= PAGE_SIZE Then Exit Sub
    Set cel = ws.Cells(EDIT_FIRST_ROW + lastEditIdx \ EDIT_COLS, 1 + (lastEditIdx Mod EDIT_COLS))
    cel.Interior.Color = CLR_ACTIVE
    cel.Font.Color = UI_InkFor(CLR_ACTIVE)
End Sub

' True if txt is 1 or 2 hex digits
Private Function IsHexByte(ByVal txt As String) As Boolean
    Dim i As Integer, ch As String
    If Len(txt) < 1 Or Len(txt) > 2 Then Exit Function
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If InStr("0123456789ABCDEFabcdef", ch) = 0 Then Exit Function
    Next i
    IsHexByte = True
End Function

' Called by Worksheet_Change of the Memory sheet. For every changed cell
' inside the editor grid: writes the byte to RAM, marks its page dirty (so
' the edit is saved to swap if the page is evicted), invalidates the cache
' lines of that frame and refreshes the panel. A value that is not 1-2 hex
' digits is rejected and the cell gets the real RAM value back.
Public Sub OnMemoryCellEdit(ByVal Target As Range)
    Dim ws As Worksheet, hit As Range, cel As Range
    Dim pfn As Integer, idx As Integer, addr As Integer, owner As Integer
    Dim txt As String, oldEv As Boolean, changed As Boolean
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If editSel = 0 Then Exit Sub
    Set hit = Intersect(Target, ws.Range(EDIT_RANGE))
    If hit Is Nothing Then Exit Sub

    pfn = editSel - 1
    oldEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fail

    For Each cel In hit.Cells
        idx = (cel.row - EDIT_FIRST_ROW) * EDIT_COLS + (cel.Column - 1)
        addr = pfn * PAGE_SIZE + idx
        txt = Trim$(CStr(cel.value))
        If IsHexByte(txt) Then
            RAM(addr) = CInt(CLng("&H" & txt))
            InvalidateFrame pfn
            owner = FrameOwner(pfn)
            If owner >= 0 Then
                PageTable(owner).Dirty = True
            End If
            lastEditPfn = pfn
            lastEditIdx = idx
            lastEditOn = True
            changed = True
        Else
            Debug.Print "OnMemoryCellEdit: '" & txt & "' is not a hex byte, value restored"
        End If
        cel.value = Right$("0" & Hex$(RAM(addr) And BYTE_MAX), 2)
    Next cel

    Application.EnableEvents = oldEv
    If changed Then RefreshMemoryPanel
    Exit Sub
fail:
    Application.EnableEvents = oldEv
    Debug.Print "OnMemoryCellEdit error: " & Err.description
End Sub

' Repaints the panel from the current machine state:
'  - frames 0..6: page that owns the frame (FrameOwner) + first bytes of RAM
'    in the current view (HEX / BIN / DEC)
'  - frame 7 keeps its fixed I/O text
'  - cache boxes go back to grey and their HIT/MISS chips are cleared
'  - the policy and format selectors match the current choices
'  - the RAM editor grid shows the current bytes of its frame
'  - the mini summary on the Simulator sheet (piece 8) is repainted
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
                For k = 0 To BytesToShow() - 1
                    If k > 0 Then preview = preview & " "
                    preview = preview & FormatByte(RAM(base + k))
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
    PaintFormatSelector ws
    RefreshEditGrid ws
    RefreshMiniMemory
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
Public Sub UI_AnimateLineFill(ByVal idx As Integer, ByVal Target As String)
    Dim ws As Worksheet, tok As Shape, src As Shape, dst As Shape, k As Integer
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
    If AnimDelay <= 0 Then Exit Sub
    If idx < 0 Or idx >= L2_LINES Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If Not HasShape(ws, "l2_line" & idx) Then Exit Sub
    If Not HasShape(ws, Target) Then Exit Sub
    Set src = ws.Shapes("l2_line" & idx)
    Set dst = ws.Shapes(Target)

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
    If Not HasShape(ws, FMT_HEX_NAME) Or Not HasShape(ws, FMT_BIN_NAME) Or Not HasShape(ws, FMT_DEC_NAME) Then fails = fails + 1: Debug.Print "  FAIL  1: missing format selector"

    ' 2) Real traffic: the cache animation must paint the boxes
    InvalidateAllCaches
    ResetStats
    v = CacheReadL1I(0)                  ' L1I miss and L2 miss
    If ws.Shapes("l1i_box").fill.ForeColor.RGB <> CLR_MISS Then fails = fails + 1: Debug.Print "  FAIL  2: l1i_box must be red after a miss"
    If ws.Shapes("lbl_l1i_box").TextFrame2.TextRange.text <> "MISS" Then fails = fails + 1: Debug.Print "  FAIL  2: lbl_l1i_box must say MISS"
    If ws.Shapes("l2_line0").fill.ForeColor.RGB <> CLR_MISS Then fails = fails + 1: Debug.Print "  FAIL  2: l2_line0 must be red (L2 miss)"
    v = CacheReadL1I(1)                  ' same line: L1I hit
    If ws.Shapes("l1i_box").fill.ForeColor.RGB <> CLR_HIT Then fails = fails + 1: Debug.Print "  FAIL  2: l1i_box must be green after a hit"
    If ws.Shapes("lbl_l1i_box").TextFrame2.TextRange.text <> "HIT" Then fails = fails + 1: Debug.Print "  FAIL  2: lbl_l1i_box must say HIT"

    ' Cleanup: caches empty, counters at 0, panel redrawn in grey
    InvalidateAllCaches
    ResetStats
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_BuildMemoryPanel: ALL OK", "Test_BuildMemoryPanel: " & fails & " FAIL(S)")
End Sub

Public Sub Test_RefreshMemoryPanel()
    Dim fails As Integer, ws As Worksheet
    Dim oldOwner As Integer, oldB0 As Integer, oldB1 As Integer, oldB2 As Integer
    Dim oldFmt As Integer
    Dim t As String
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' Save the state this test changes. The test needs the HEX view.
    oldOwner = FrameOwner(0)
    oldB0 = RAM(0): oldB1 = RAM(1): oldB2 = RAM(2)
    oldFmt = viewFmt
    viewFmt = FMT_HEX

    ' 1) Frame 0 owned by page 3 with known bytes
    FrameOwner(0) = 3
    RAM(0) = &H1A: RAM(1) = &HFF: RAM(2) = &H5
    RefreshMemoryPanel
    t = ws.Shapes(PFX_FRM & "0").TextFrame2.TextRange.text
    If InStr(t, "P3") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: frm_0 must show P3, got: " & Replace(t, vbLf, "/")
    If InStr(t, "1A FF 05") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: frm_0 must show bytes 1A FF 05, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(PFX_FRM & "0").fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  1: occupied frame must use CLR_LOAD"

    ' 2) Frame 0 free again
    FrameOwner(0) = -1
    RefreshMemoryPanel
    t = ws.Shapes(PFX_FRM & "0").TextFrame2.TextRange.text
    If InStr(t, "free") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: frm_0 must say free, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(PFX_FRM & "0").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  2: free frame must be grey"

    ' 3) Frame 7 keeps its I/O text
    t = ws.Shapes(PFX_FRM & IO_FRAME).TextFrame2.TextRange.text
    If InStr(t, "I/O") = 0 Then fails = fails + 1: Debug.Print "  FAIL  3: frm_7 must keep I/O text"

    ' 4) Cache boxes painted by an animation go back to grey and the chip is cleared
    UI_SetFill "l1i_box", CLR_MISS
    UI_SetText "lbl_l1i_box", "MISS"
    UI_SetFill "l2_line2", CLR_HIT
    UI_SetText "lbl_l2_line2", "HIT"
    RefreshMemoryPanel
    If ws.Shapes("l1i_box").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  4: l1i_box must be grey after refresh"
    If ws.Shapes("lbl_l1i_box").TextFrame2.TextRange.text <> "" Then fails = fails + 1: Debug.Print "  FAIL  4: lbl_l1i_box must be empty after refresh"
    If ws.Shapes("l2_line2").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  4: l2_line2 must be grey after refresh"
    If ws.Shapes("lbl_l2_line2").TextFrame2.TextRange.text <> "" Then fails = fails + 1: Debug.Print "  FAIL  4: lbl_l2_line2 must be empty after refresh"

    ' Cleanup: restore the state and leave the panel as the machine has it
    FrameOwner(0) = oldOwner
    RAM(0) = oldB0: RAM(1) = oldB1: RAM(2) = oldB2
    viewFmt = oldFmt
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
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.text
    If InStr(t, "0x0A5") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show vaddr 0x0A5, got: " & Replace(t, vbLf, "/")
    If InStr(t, "VPN 5") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show VPN 5, got: " & Replace(t, vbLf, "/")
    If InStr(t, "0101") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show VPN in binary 0101, got: " & Replace(t, vbLf, "/")
    If InStr(t, "PFN 2") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show PFN 2, got: " & Replace(t, vbLf, "/")
    If InStr(t, "paddr 0x45") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: must show paddr 0x45, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(XLT_NAME).fill.ForeColor.RGB <> CLR_HIT Then fails = fails + 1: Debug.Print "  FAIL  1: present page must be green"

    ' 2) Page 5 not present: PAGE FAULT in red
    PageTable(5).Present = False
    ShowTranslation 5, 5
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.text
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
    t = ws.Shapes(PFX_FRM & "2").TextFrame2.TextRange.text
    If InStr(t, "P5") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: frm_2 must show P5, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(PFX_FRM & "2").fill.ForeColor.RGB <> CLR_ACTIVE Then fails = fails + 1: Debug.Print "  FAIL  2: frm_2 must be coral (CLR_ACTIVE)"
    If HasShape(ws, TOKEN_NAME) Then fails = fails + 1: Debug.Print "  FAIL  2: the token must be deleted after the animation"

    ' 3) Frame 7 (I/O) is never touched
    t = ws.Shapes(PFX_FRM & IO_FRAME).TextFrame2.TextRange.text
    UI_AnimateSwapIn 5, IO_FRAME
    If ws.Shapes(PFX_FRM & IO_FRAME).TextFrame2.TextRange.text <> t Then fails = fails + 1: Debug.Print "  FAIL  3: frm_7 must not change"

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
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.text
    If InStr(t, "VPN 5") = 0 Or InStr(t, "PFN 2") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: strip must show VPN 5 -> PFN 2, got: " & Replace(t, vbLf, "/")

    ' 2) A click outside the page table does nothing
    OnPageTableClick ws.Range("H3")
    If ws.Shapes(PFX_FRM & "2").line.weight <> HL_WEIGHT Then fails = fails + 1: Debug.Print "  FAIL  2: a click outside A2:E17 must keep the outline"

    ' 3) Page 6 not present: the outline of frm_2 goes away and the strip says PAGE FAULT
    PageTable(6).Present = False
    OnPageTableClick ws.Range("C8")
    If ws.Shapes(PFX_FRM & "2").line.weight = HL_WEIGHT Then fails = fails + 1: Debug.Print "  FAIL  3: frm_2 must lose its outline"
    t = ws.Shapes(XLT_NAME).TextFrame2.TextRange.text
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

Public Sub Test_ViewFormat()
    Dim fails As Integer, ws As Worksheet, t As String
    Dim oldOwner As Integer, oldB0 As Integer, oldB1 As Integer, oldB2 As Integer
    Dim oldFmt As Integer
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' Save the state this test changes
    oldOwner = FrameOwner(0)
    oldB0 = RAM(0): oldB1 = RAM(1): oldB2 = RAM(2)
    oldFmt = viewFmt

    ' Frame 0 owned by page 3 with known bytes: 0x1A = 26 = 00011010, 0xFF = 255
    FrameOwner(0) = 3
    RAM(0) = &H1A: RAM(1) = &HFF: RAM(2) = &H5

    ' 1) Buttons exist and call the right macros
    If Not HasShape(ws, FMT_HEX_NAME) Or Not HasShape(ws, FMT_BIN_NAME) Or Not HasShape(ws, FMT_DEC_NAME) Then
        fails = fails + 1: Debug.Print "  FAIL  1: missing format buttons"
    Else
        If InStr(ws.Shapes(FMT_HEX_NAME).OnAction, "SetViewHex") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: fmt_hex must call SetViewHex"
        If InStr(ws.Shapes(FMT_BIN_NAME).OnAction, "SetViewBin") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: fmt_bin must call SetViewBin"
        If InStr(ws.Shapes(FMT_DEC_NAME).OnAction, "SetViewDec") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: fmt_dec must call SetViewDec"
    End If

    ' 2) HEX: three bytes
    SetViewHex
    t = ws.Shapes(PFX_FRM & "0").TextFrame2.TextRange.text
    If InStr(t, "1A FF 05") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: HEX must show 1A FF 05, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(FMT_HEX_NAME).fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  2: fmt_hex must be amber"
    If ws.Shapes(FMT_BIN_NAME).fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  2: fmt_bin must be grey"

    ' 3) BIN: two bytes, 8 bits each
    SetViewBin
    t = ws.Shapes(PFX_FRM & "0").TextFrame2.TextRange.text
    If InStr(t, "00011010 11111111") = 0 Then fails = fails + 1: Debug.Print "  FAIL  3: BIN must show 00011010 11111111, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(FMT_BIN_NAME).fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  3: fmt_bin must be amber"
    If ws.Shapes(FMT_HEX_NAME).fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  3: fmt_hex must be grey"

    ' 4) DEC: two bytes
    SetViewDec
    t = ws.Shapes(PFX_FRM & "0").TextFrame2.TextRange.text
    If InStr(t, "26 255") = 0 Then fails = fails + 1: Debug.Print "  FAIL  4: DEC must show 26 255, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(FMT_DEC_NAME).fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  4: fmt_dec must be amber"

    ' 5) Frame 7 (I/O) keeps its text in every view
    t = ws.Shapes(PFX_FRM & IO_FRAME).TextFrame2.TextRange.text
    If InStr(t, "I/O") = 0 Then fails = fails + 1: Debug.Print "  FAIL  5: frm_7 must keep I/O text"

    ' Cleanup: restore the state and the view, and redraw the panel
    FrameOwner(0) = oldOwner
    RAM(0) = oldB0: RAM(1) = oldB1: RAM(2) = oldB2
    viewFmt = oldFmt
    BuildMemoryPanel
    RefreshMemoryPanel

    Debug.Print IIf(fails = 0, "Test_ViewFormat: ALL OK", "Test_ViewFormat: " & fails & " FAIL(S)")
End Sub

Public Sub Test_RamEdit()
    Dim fails As Integer, ws As Worksheet, v As Integer
    Dim oldPte5 As TPageEntry, oldPte6 As TPageEntry
    Dim oldOwner2 As Integer, oldRam69 As Integer
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    BuildMemoryPanel

    ' Save the state this test changes. Page 5 sits in frame 2, byte 5 of
    ' that frame is RAM(69) (frame 2 starts at 64).
    oldPte5 = PageTable(5)
    oldPte6 = PageTable(6)
    oldOwner2 = FrameOwner(2)
    oldRam69 = RAM(69)
    PageTable(5).Present = True
    PageTable(5).pfn = 2
    PageTable(5).Dirty = False
    FrameOwner(2) = 5
    RAM(69) = &H1A

    ' 1) Click on the row of P5: the grid shows RAM(69) = 1A in F47 (byte 5)
    OnPageTableClick ws.Range("A7")
    If CStr(ws.Range("F47").value) <> "1A" Then fails = fails + 1: Debug.Print "  FAIL  1: F47 must show 1A, got: '" & CStr(ws.Range("F47").value) & "'"
    If InStr(CStr(ws.Cells(EDIT_LABEL_ROW, 1).value), "F2") = 0 Then fails = fails + 1: Debug.Print "  FAIL  1: label must mention F2"

    ' 2) Edit: type "ff" in F47. RAM changes, page is dirty, cell is normalized.
    '    A cached copy of the old line must not survive (RAM(69) is in line 8..15 of frame 2).
    InvalidateAllCaches
    ResetStats
    v = CacheReadL1D(69)
    If v <> &H1A Then fails = fails + 1: Debug.Print "  FAIL  2: before the edit CacheReadL1D(69) = " & v & " (expected 26)"
    Application.EnableEvents = False
    ws.Range("F47").value = "ff"
    Application.EnableEvents = True
    OnMemoryCellEdit ws.Range("F47")
    If RAM(69) <> 255 Then fails = fails + 1: Debug.Print "  FAIL  2: RAM(69) = " & RAM(69) & " (expected 255)"
    If CStr(ws.Range("F47").value) <> "FF" Then fails = fails + 1: Debug.Print "  FAIL  2: F47 must be normalized to FF"
    If Not PageTable(5).Dirty Then fails = fails + 1: Debug.Print "  FAIL  2: page 5 must be dirty after the edit"
    v = CacheReadL1D(69)
    If v <> 255 Then fails = fails + 1: Debug.Print "  FAIL  2: after the edit CacheReadL1D(69) = " & v & " (expected 255, stale cache?)"

    ' 3) Invalid value: rejected, the cell gets the real value back
    Application.EnableEvents = False
    ws.Range("F47").value = "ZZ"
    Application.EnableEvents = True
    OnMemoryCellEdit ws.Range("F47")
    If RAM(69) <> 255 Then fails = fails + 1: Debug.Print "  FAIL  3: RAM(69) must stay 255 after an invalid value"
    If CStr(ws.Range("F47").value) <> "FF" Then fails = fails + 1: Debug.Print "  FAIL  3: F47 must go back to FF"

    ' 4) One digit is valid: "7" -> 07
    Application.EnableEvents = False
    ws.Range("F47").value = "7"
    Application.EnableEvents = True
    OnMemoryCellEdit ws.Range("F47")
    If RAM(69) <> 7 Then fails = fails + 1: Debug.Print "  FAIL  4: RAM(69) = " & RAM(69) & " (expected 7)"
    If CStr(ws.Range("F47").value) <> "07" Then fails = fails + 1: Debug.Print "  FAIL  4: F47 must show 07"

    ' 5) The frame preview of the panel follows the edit (byte 0 of frame 2 is not
    '    the edited one, so only check the panel still shows P5)
    If InStr(ws.Shapes(PFX_FRM & "2").TextFrame2.TextRange.text, "P5") = 0 Then fails = fails + 1: Debug.Print "  FAIL  5: frm_2 must still show P5 after the refresh"

    ' 6) Click on a page that is not in RAM: the grid is cleared
    PageTable(6).Present = False
    OnPageTableClick ws.Range("A8")
    If CStr(ws.Range("F47").value) <> "" Then fails = fails + 1: Debug.Print "  FAIL  6: grid must be empty for a page that is not in RAM"

    ' 7) With no frame selected an edit does nothing
    Application.EnableEvents = False
    ws.Range("F47").value = "AA"
    Application.EnableEvents = True
    OnMemoryCellEdit ws.Range("F47")
    If RAM(69) <> 7 Then fails = fails + 1: Debug.Print "  FAIL  7: RAM(69) must not change when no frame is selected"

    ' Cleanup: restore everything and redraw the panel (this also clears the grid)
    RAM(69) = oldRam69
    FrameOwner(2) = oldOwner2
    PageTable(5) = oldPte5
    PageTable(6) = oldPte6
    InvalidateAllCaches
    ResetStats
    BuildMemoryPanel

    Debug.Print IIf(fails = 0, "Test_RamEdit: ALL OK", "Test_RamEdit: " & fails & " FAIL(S)")
End Sub

Public Sub Test_MiniMemory()
    Dim fails As Integer, i As Integer, ws As Worksheet, t As String
    Dim oldOwner As Integer, oldB0 As Integer, oldB1 As Integer, oldB2 As Integer
    Dim oldFmt As Integer, oldPolicy As Integer
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    BuildMemoryPanel
    BuildMiniMemory ws

    ' Save the state this test changes. The test starts in the HEX view.
    oldOwner = FrameOwner(0)
    oldB0 = RAM(0): oldB1 = RAM(1): oldB2 = RAM(2)
    oldFmt = viewFmt
    oldPolicy = ReplacePolicy
    viewFmt = FMT_HEX

    ' 1) Every mini shape exists
    For i = 0 To NUM_FRAMES - 1
        If Not HasShape(ws, MINI_PREFIX & i) Then fails = fails + 1: Debug.Print "  FAIL  1: missing shape " & MINI_PREFIX & i
    Next i
    If Not HasShape(ws, MINI_TITLE_NAME) Then fails = fails + 1: Debug.Print "  FAIL  1: missing " & MINI_TITLE_NAME
    If Not HasShape(ws, MINI_INFO_NAME) Then fails = fails + 1: Debug.Print "  FAIL  1: missing " & MINI_INFO_NAME

    ' 2) Frame 0 owned by page 3: the refresh of the Memory panel repaints the mini
    FrameOwner(0) = 3
    RAM(0) = &H1A: RAM(1) = &HFF: RAM(2) = &H5
    RefreshMemoryPanel
    t = ws.Shapes(MINI_PREFIX & "0").TextFrame2.TextRange.text
    If InStr(t, "P3") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: mini_0 must show P3, got: " & Replace(t, vbLf, "/")
    If InStr(t, "1A FF 05") = 0 Then fails = fails + 1: Debug.Print "  FAIL  2: mini_0 must show 1A FF 05, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(MINI_PREFIX & "0").fill.ForeColor.RGB <> CLR_LOAD Then fails = fails + 1: Debug.Print "  FAIL  2: occupied mini frame must use CLR_LOAD"

    ' 3) BIN view
    viewFmt = FMT_BIN
    RefreshMemoryPanel
    t = ws.Shapes(MINI_PREFIX & "0").TextFrame2.TextRange.text
    If InStr(t, "00011010 11111111") = 0 Then fails = fails + 1: Debug.Print "  FAIL  3: mini_0 must show BIN bytes, got: " & Replace(t, vbLf, "/")
    viewFmt = FMT_HEX

    ' 4) Frame 0 free again
    FrameOwner(0) = -1
    RefreshMemoryPanel
    t = ws.Shapes(MINI_PREFIX & "0").TextFrame2.TextRange.text
    If InStr(t, "free") = 0 Then fails = fails + 1: Debug.Print "  FAIL  4: mini_0 must say free, got: " & Replace(t, vbLf, "/")
    If ws.Shapes(MINI_PREFIX & "0").fill.ForeColor.RGB <> CLR_HIGHZ Then fails = fails + 1: Debug.Print "  FAIL  4: free mini frame must be grey"

    ' 5) Frame 7 keeps its I/O text
    t = ws.Shapes(MINI_PREFIX & IO_FRAME).TextFrame2.TextRange.text
    If InStr(t, "I/O") = 0 Then fails = fails + 1: Debug.Print "  FAIL  5: mini_7 must keep I/O text"

    ' 6) Info line follows the policy
    ReplacePolicy = POLICY_LRU
    RefreshMemoryPanel
    t = ws.Shapes(MINI_INFO_NAME).TextFrame2.TextRange.text
    If InStr(t, "LRU") = 0 Then fails = fails + 1: Debug.Print "  FAIL  6: mini_info must say LRU, got: " & t
    ReplacePolicy = POLICY_FIFO
    RefreshMemoryPanel
    t = ws.Shapes(MINI_INFO_NAME).TextFrame2.TextRange.text
    If InStr(t, "FIFO") = 0 Then fails = fails + 1: Debug.Print "  FAIL  6: mini_info must say FIFO, got: " & t

    ' 7) The mini names do not clash with the frames of the Memory sheet
    If Not HasShape(ThisWorkbook.Worksheets(SH_MEM), PFX_FRM & "0") Then fails = fails + 1: Debug.Print "  FAIL  7: frm_0 of the Memory sheet must still exist"
    If HasShape(ws, PFX_FRM & "0") Then fails = fails + 1: Debug.Print "  FAIL  7: frm_0 must not exist on the Simulator sheet"

    ' Cleanup: restore the state and redraw
    FrameOwner(0) = oldOwner
    RAM(0) = oldB0: RAM(1) = oldB1: RAM(2) = oldB2
    viewFmt = oldFmt
    ReplacePolicy = oldPolicy
    BuildMemoryPanel
    RefreshMemoryPanel

    Debug.Print IIf(fails = 0, "Test_MiniMemory: ALL OK", "Test_MiniMemory: " & fails & " FAIL(S)")
End Sub

' ---------- Sheet handlers (they live on the Memory sheet, not in src) ----------

' Inserts Worksheet_SelectionChange and Worksheet_Change into the Memory
' sheet module. Safe to run twice. Needs Trust access to the VBA project
' object model. Uses the sheet CodeName, not a fixed SheetN name.
Public Sub InstallMemoryHandlers()
    Dim ws As Worksheet
    Dim comp As Object
    Dim cm As Object
    Dim src As String
    Dim added As Integer
    On Error GoTo EH

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    On Error GoTo EH
    If ws Is Nothing Then
        MsgBox "The Memory sheet (" & SH_MEM & ") was not found.", vbExclamation, "Byte Bakers"
        Exit Sub
    End If

    Set comp = ThisWorkbook.VBProject.VBComponents(ws.CodeName)
    Set cm = comp.CodeModule
    src = ModuleSource(cm)
    If Not HandlerPresent(src, "Worksheet_SelectionChange") Then
        cm.AddFromString SelectionHandlerText()
        added = added + 1
    ElseIf InStr(1, src, "OnPageTableClick", vbTextCompare) = 0 Then
        MsgBox "Worksheet_SelectionChange already exists on the Memory sheet but does not call OnPageTableClick. No second handler was added.", vbExclamation, "Byte Bakers"
        Exit Sub
    End If

    src = ModuleSource(cm)
    If Not HandlerPresent(src, "Worksheet_Change") Then
        cm.AddFromString ChangeHandlerText()
        added = added + 1
    ElseIf InStr(1, src, "OnMemoryCellEdit", vbTextCompare) = 0 Then
        MsgBox "Worksheet_Change already exists on the Memory sheet but does not call OnMemoryCellEdit. No second handler was added.", vbExclamation, "Byte Bakers"
        Exit Sub
    End If

    If added = 0 Then
        MsgBox "Memory handlers were already installed.", vbInformation, "Byte Bakers"
    Else
        MsgBox "Memory handlers installed (" & added & "). Save the workbook (Ctrl+S).", vbInformation, "Byte Bakers"
    End If
    Exit Sub
EH:
    If Err.Number = 1004 Then
        MsgBox "Excel blocked access to the VBA project." & vbCrLf & vbCrLf & _
            "Turn on: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model.", _
            vbExclamation, "Byte Bakers"
    Else
        MsgBox "InstallMemoryHandlers: " & Err.description, vbCritical, "Byte Bakers"
    End If
End Sub

Private Function ModuleSource(ByVal cm As Object) As String
    If cm.CountOfLines = 0 Then
        ModuleSource = ""
    Else
        ModuleSource = cm.lines(1, cm.CountOfLines)
    End If
End Function

Private Function HandlerPresent(ByVal src As String, ByVal subName As String) As Boolean
    HandlerPresent = (InStr(1, src, "Sub " & subName, vbTextCompare) > 0)
End Function

Private Function SelectionHandlerText() As String
    SelectionHandlerText = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
        "    OnPageTableClick Target" & vbCrLf & _
        "End Sub" & vbCrLf
End Function

Private Function ChangeHandlerText() As String
    ChangeHandlerText = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
        "    OnMemoryCellEdit Target" & vbCrLf & _
        "End Sub" & vbCrLf
End Function

' Checks the page-table style and the coral edited byte, then restores
' the machine state this test changed.
Public Function Test_PanelStyle() As Long
    Dim fails As Long
    Dim ws As Worksheet
    Dim old0 As TPageEntry, old1 As TPageEntry, old3 As TPageEntry
    Dim old5 As TPageEntry, old15 As TPageEntry
    Dim oldOwn2 As Integer, oldRam As Integer
    Dim oldL1I As TCacheLine, oldL1D As TCacheLine
    Dim oldL2(0 To 3) As TCacheLine
    Dim oldStats As TStats
    Dim oldEv As Boolean
    Dim i As Integer
    Dim cel As Range

    fails = 0
    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    oldEv = Application.EnableEvents
    old0 = PageTable(0)
    old1 = PageTable(1)
    old3 = PageTable(3)
    old5 = PageTable(5)
    old15 = PageTable(15)
    oldOwn2 = FrameOwner(2)
    oldRam = RAM(64)
    oldL1I = L1I
    oldL1D = L1D
    For i = 0 To L2_LINES - 1
        oldL2(i) = L2(i)
    Next i
    oldStats = Stats

    RefreshPageTableView
    fails = fails + StyleCheck("header is navy with cream bold text", HeaderLooksRight(ws))

    PageTable(3).Present = False
    RefreshPageTableRow 3
    Set cel = ws.Cells(5, 1)
    fails = fails + StyleCheck("absent page is cream Consolas with ink", _
        cel.Interior.Color = CLR_TEXT And cel.Font.Color = CLR_INK And cel.Font.Name = "Consolas")
    fails = fails + StyleCheck("absent page border is CLR_BORDER", PageBorderOk(cel))
    fails = fails + StyleCheck("column P is 0 when the page is absent", ws.Cells(5, 2).value = 0)

    PageTable(0).Present = True
    PageTable(0).pfn = 0
    RefreshPageTableRow 0
    Set cel = ws.Cells(2, 1)
    fails = fails + StyleCheck("page 0 present is amber", _
        cel.Interior.Color = CLR_LOAD And cel.Font.Color = UI_InkFor(CLR_LOAD))
    fails = fails + StyleCheck("column P is 1 when the page is present", ws.Cells(2, 2).value = 1)

    PageTable(1).Present = True
    PageTable(1).pfn = 1
    RefreshPageTableRow 1
    fails = fails + StyleCheck("data page present is gray", ws.Cells(3, 1).Interior.Color = CLR_HIGHZ)

    PageTable(15).Present = True
    PageTable(15).pfn = IO_FRAME
    RefreshPageTableRow 15
    Set cel = ws.Cells(17, 1)
    fails = fails + StyleCheck("page 15 present is LCD green", _
        cel.Interior.Color = CLR_LCD And cel.Font.Color = UI_InkFor(CLR_LCD))

    PageTable(5).Present = True
    PageTable(5).pfn = 2
    PageTable(5).Dirty = False
    FrameOwner(2) = 5
    RAM(64) = &H1A
    OnPageTableClick ws.Range("A7")
    Application.EnableEvents = False
    ws.Range("A47").value = "2B"
    Application.EnableEvents = True
    OnMemoryCellEdit ws.Range("A47")
    fails = fails + StyleCheck("edited byte is coral", _
        ws.Range("A47").Interior.Color = CLR_ACTIVE And ws.Range("A47").Font.Color = UI_InkFor(CLR_ACTIVE))
    fails = fails + StyleCheck("neighbor cell is not coral", ws.Range("B47").Interior.Pattern = xlNone)
    fails = fails + StyleCheck("edit still writes RAM and marks the page dirty", _
        RAM(64) = &H2B And PageTable(5).Dirty)
    RefreshMemoryPanel
    fails = fails + StyleCheck("edited byte stays coral after RefreshMemoryPanel", _
        ws.Range("A47").Interior.Color = CLR_ACTIVE)

    GoTo Clean
CleanFail:
    fails = fails + 1
    Debug.Print "  FAIL  Test_PanelStyle stopped: " & Err.description
Clean:
    On Error Resume Next
    RAM(64) = oldRam
    FrameOwner(2) = oldOwn2
    PageTable(0) = old0
    PageTable(1) = old1
    PageTable(3) = old3
    PageTable(5) = old5
    PageTable(15) = old15
    L1I = oldL1I
    L1D = oldL1D
    For i = 0 To L2_LINES - 1
        L2(i) = oldL2(i)
    Next i
    Stats = oldStats
    Application.EnableEvents = oldEv
    RefreshPageTableView
    BuildMemoryPanel
    ResetMachine
    RefreshPageTableView
    RefreshMemoryPanel
    On Error GoTo 0
    If fails = 0 Then
        Debug.Print "Test_PanelStyle: ALL OK"
    Else
        Debug.Print "Test_PanelStyle: " & fails & " FAIL(S)"
    End If
    Test_PanelStyle = fails
End Function

Private Function StyleCheck(ByVal what As String, ByVal passed As Boolean) As Long
    If passed Then
        Debug.Print "  OK    " & what
        StyleCheck = 0
    Else
        Debug.Print "  FAIL  " & what
        StyleCheck = 1
    End If
End Function

Private Function HeaderLooksRight(ByVal ws As Worksheet) As Boolean
    Dim c As Integer
    HeaderLooksRight = True
    For c = 1 To 5
        If ws.Cells(1, c).Interior.Color <> CLR_BG Then
            HeaderLooksRight = False
        End If
        If ws.Cells(1, c).Font.Color <> CLR_TEXT Then
            HeaderLooksRight = False
        End If
        If ws.Cells(1, c).Font.bold <> True Then
            HeaderLooksRight = False
        End If
    Next c
End Function

Private Function PageBorderOk(ByVal cel As Range) As Boolean
    PageBorderOk = True
    If cel.Borders(xlEdgeLeft).LineStyle <> xlContinuous Then
        PageBorderOk = False
    End If
    If cel.Borders(xlEdgeLeft).Color <> CLR_BORDER Then
        PageBorderOk = False
    End If
    If cel.Borders(xlEdgeRight).Color <> CLR_BORDER Then
        PageBorderOk = False
    End If
    If cel.Borders(xlEdgeTop).Color <> CLR_BORDER Then
        PageBorderOk = False
    End If
    If cel.Borders(xlEdgeBottom).Color <> CLR_BORDER Then
        PageBorderOk = False
    End If
End Function

