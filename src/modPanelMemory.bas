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
' =====================================================================

Private Const PANEL_TOP_ROW As Long = 28
Private Const XLT_ROW As Long = 37
Private Const XLT_NAME As String = "xlt_box"
Private Const BOX_W As Single = 78
Private Const STEP_X As Single = 84
Private Const FRAME_H As Single = 50
Private Const CACHE_H As Single = 36
Private Const CHIP_H As Single = 16
Private Const PREVIEW_BYTES As Integer = 3

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

    ' Translation strip under the caches
    UI_AddBox ws, XLT_NAME, leftX, ws.Rows(XLT_ROW).Top, NUM_FRAMES * STEP_X - 6, 36, _
        "Translation: waiting for a memory access"
End Sub

' A cache box plus its label chip underneath (lbl_<name>)
Private Sub AddCacheBox(ByVal ws As Worksheet, ByVal nm As String, _
        ByVal l As Single, ByVal t As Single, ByVal txt As String)
    Dim chip As Shape
    UI_AddBox ws, nm, l, t, BOX_W, CACHE_H, txt
    Set chip = UI_AddBox(ws, "lbl_" & nm, l, t + CACHE_H + 2, BOX_W, CHIP_H, "")
    chip.TextFrame2.TextRange.Font.size = 9
End Sub

' Repaints the panel from the current machine state:
'  - frames 0..6: page that owns the frame (FrameOwner) + first bytes of RAM
'  - frame 7 keeps its fixed I/O text
'  - cache boxes go back to grey and their HIT/MISS chips are cleared
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

