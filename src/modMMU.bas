Attribute VB_Name = "modMMU"
Option Explicit
' =====================================================================
' modMMU - Memory management unit
'   K3: page table | K4: translation | K5: Page Fault (FIFO / LRU)
' Virtual address  (9 bits) = [VPN 4 bits | offset 5 bits]
' Physical address (8 bits) = [PFN 3 bits | offset 5 bits]
' Page 15 is I/O: always in frame 7, pinned, never cached.
' =====================================================================

' ---------- K3: PAGE TABLE ----------

' Leaves the page table with every page absent (P=0),
' except page 15, which stays in frame 7 and pinned.
Public Sub InitPageTable()
    Dim i As Integer, emptyPTE As TPageEntry
    For i = 0 To NUM_VPAGES - 1: PageTable(i) = emptyPTE: Next i
    For i = 0 To NUM_FRAMES - 1: FrameOwner(i) = -1: Next i
    With PageTable(IO_VPAGE)
        .Present = True
        .pfn = IO_FRAME
        .Pinned = True
    End With
    FrameOwner(IO_FRAME) = IO_VPAGE
    RefreshPageTableView
End Sub

' Draws the whole VPN / P / PFN / D / R table on the Memory sheet (A1:E17)
Public Sub RefreshPageTableView()
    Dim ws As Worksheet, i As Integer
    Dim hdr As Range
    Set ws = GetMemorySheet()
    Set hdr = ws.Range("A1:E1")
    hdr.value = Array("VPN", "P", "PFN", "D", "R")
    hdr.Font.Name = "Consolas"
    hdr.Font.bold = True
    hdr.Font.Color = CLR_TEXT
    hdr.Interior.Color = CLR_BG
    ApplyPageBorders hdr
    For i = 0 To NUM_VPAGES - 1
        RefreshPageTableRow i
    Next i
End Sub

' Redraws only the row of one page (called on every access).
' Present: type color (page 0 amber, pages 1-14 gray, page 15 LCD green).
' Absent: cream fill and ink text. Column P stays 1 or 0.
Public Sub RefreshPageTableRow(ByVal vpn As Integer)
    Dim ws As Worksheet, r As Long
    Dim band As Range
    Dim fill As Long
    Set ws = GetMemorySheet()
    r = vpn + 2
    ws.Cells(r, 1).value = vpn
    ws.Cells(r, 2).value = IIf(PageTable(vpn).Present, 1, 0)
    ws.Cells(r, 3).value = IIf(PageTable(vpn).Present, CStr(PageTable(vpn).pfn), "-")
    ws.Cells(r, 4).value = IIf(PageTable(vpn).Dirty, 1, 0)
    ws.Cells(r, 5).value = IIf(PageTable(vpn).Ref, 1, 0)
    Set band = ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
    If PageTable(vpn).Present Then
        fill = PresentFill(vpn)
        band.Interior.Color = fill
        band.Font.Color = UI_InkFor(fill)
    Else
        band.Interior.Color = CLR_TEXT
        band.Font.Color = CLR_INK
    End If
    band.Font.Name = "Consolas"
    ApplyPageBorders band
End Sub

' Present row fill by page type. Not used for an absent page.
Private Function PresentFill(ByVal vpn As Integer) As Long
    If vpn = 0 Then
        PresentFill = CLR_LOAD
    ElseIf vpn = IO_VPAGE Then
        PresentFill = CLR_LCD
    Else
        PresentFill = CLR_HIGHZ
    End If
End Function

Private Sub ApplyPageBorders(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .weight = xlThin
        .Color = CLR_BORDER
    End With
End Sub

Private Function GetMemorySheet() As Worksheet
    On Error Resume Next
    Set GetMemorySheet = ThisWorkbook.Worksheets(SH_MEM)
    On Error GoTo 0
    If GetMemorySheet Is Nothing Then
        Set GetMemorySheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetMemorySheet.Name = SH_MEM
    End If
End Function

' ---------- K4: VIRTUAL -> PHYSICAL TRANSLATION ----------

' Translates a virtual address (0..511) to a physical one (0..255).
'  - If the page is absent (P=0) it calls HandlePageFault (K5) first.
'  - Sets Ref=1 on every access and Dirty=1 if it is a write.
'  - Leaves the result in CPU.PADDR and updates the table row.
Public Function Translate(ByVal vaddr As Integer, ByVal isWrite As Boolean) As Integer
    Dim vpn As Integer, offset As Integer, paddr As Integer
    If vaddr < 0 Or vaddr > VADDR_MAX Then
        Err.Raise vbObjectError + 531, "MMU", "Direccion virtual fuera de rango (0-" & VADDR_MAX & "): " & vaddr
    End If
    vpn = vaddr \ PAGE_SIZE
    offset = vaddr Mod PAGE_SIZE
    ShowTranslation vpn, offset                      ' shows the live translation (K10)
    If Not PageTable(vpn).Present Then
        HandlePageFault vpn                          ' K5: loads the page
        If Not PageTable(vpn).Present Then
            Err.Raise vbObjectError + 532, "MMU", "La pagina " & vpn & " sigue ausente despues de HandlePageFault"
        End If
    End If
    With PageTable(vpn)
        .Ref = True
        .LastUse = Clock
        If isWrite Then .Dirty = True
        paddr = .pfn * PAGE_SIZE + offset
    End With
    RefreshPageTableRow vpn
    CPU.paddr = paddr
    Translate = paddr
End Function

' ---------- K5: PAGE FAULT ----------

' Loads page vpn into a RAM frame:
'  1) uses a free frame if there is one; otherwise picks a victim (FIFO or LRU per ReplacePolicy)
'  2) if the victim is dirty (Dirty=1) it is saved to swap
'  3) invalidates the cache lines of the frame (L1I, L1D, L2)
'  4) copies the page from swap and sets P=1
' Page 15 (I/O, Pinned) is never a victim.
Public Sub HandlePageFault(ByVal vpn As Integer)
    Dim pfn As Integer, victim As Integer
    If vpn < 0 Or vpn > NUM_VPAGES - 1 Then
        Err.Raise vbObjectError + 533, "MMU", "Pagina virtual fuera de rango (0-" & NUM_VPAGES - 1 & "): " & vpn
    End If
    If PageTable(vpn).Present Then Exit Sub          ' already loaded: no fault

    Stats.PageFaults = Stats.PageFaults + 1
    UI_PageFaultFlash vpn                            ' red flash (K10)

    pfn = FindFreeFrame()
    If pfn = -1 Then
        victim = ChooseVictim()
        pfn = PageTable(victim).pfn
        If PageTable(victim).Dirty Then SwapOut victim, pfn
        With PageTable(victim)
            .Present = False
            .Dirty = False
            .Ref = False
            .pfn = 0
        End With
        FrameOwner(pfn) = -1
        InvalidateFrame pfn                          ' K8: invalidates L1I, L1D and L2 for that frame
        RefreshPageTableRow victim
    End If

    SwapIn vpn, pfn
    UI_AnimateSwapIn vpn, pfn                        ' swap -> frame block (K10)

    With PageTable(vpn)
        .Present = True
        .pfn = pfn
        .Dirty = False
        .Ref = False
        .LoadedAt = Clock                            ' for FIFO
        .LastUse = Clock                             ' for LRU
    End With
    FrameOwner(pfn) = vpn
    RefreshPageTableRow vpn
End Sub

' First free frame, or -1 if RAM is full
Private Function FindFreeFrame() As Integer
    Dim i As Integer
    For i = 0 To NUM_FRAMES - 1
        If FrameOwner(i) = -1 Then FindFreeFrame = i: Exit Function
    Next i
    FindFreeFrame = -1
End Function

' Picks the victim: FIFO = lowest LoadedAt, LRU = lowest LastUse.
' Only present and NOT pinned pages (page 15 is never evicted).
Private Function ChooseVictim() As Integer
    Dim i As Integer, best As Integer, bestTime As Long, t As Long
    best = -1
    bestTime = 2147483647
    For i = 0 To NUM_VPAGES - 1
        If PageTable(i).Present And Not PageTable(i).Pinned Then
            If ReplacePolicy = POLICY_FIFO Then
                t = PageTable(i).LoadedAt
            Else
                t = PageTable(i).LastUse
            End If
            If t < bestTime Then
                best = i
                bestTime = t
            End If
        End If
    Next i
    If best = -1 Then
        Err.Raise vbObjectError + 530, "MMU", "No hay ninguna pagina desalojable (todas estan fijas)"
    End If
    ChooseVictim = best
End Function

' ---------- TESTS ----------
' Run with F5 (cursor inside the Sub) and watch the Immediate window (Ctrl+G)
Public Sub Test_PageTable()
    Dim i As Integer, fails As Integer, ws As Worksheet

    ' 1) Initial state in memory
    InitPageTable
    For i = 0 To NUM_VPAGES - 2
        If PageTable(i).Present Then fails = fails + 1: Debug.Print "  FAIL  pagina " & i & " deberia estar ausente"
    Next i
    For i = 0 To NUM_FRAMES - 2
        If FrameOwner(i) <> -1 Then fails = fails + 1: Debug.Print "  FAIL  marco " & i & " deberia estar libre"
    Next i
    If Not (PageTable(IO_VPAGE).Present And PageTable(IO_VPAGE).pfn = IO_FRAME) Then fails = fails + 1: Debug.Print "  FAIL  pagina 15 no esta en el marco 7"
    If Not PageTable(IO_VPAGE).Pinned Then fails = fails + 1: Debug.Print "  FAIL  pagina 15 no esta fija"
    If FrameOwner(IO_FRAME) <> IO_VPAGE Then fails = fails + 1: Debug.Print "  FAIL  el marco 7 no pertenece a la pagina 15"

    ' 2) What is shown on the Memory sheet
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If ws.Cells(1, 1).value <> "VPN" Then fails = fails + 1: Debug.Print "  FAIL  falta el encabezado VPN"
    If ws.Cells(2, 2).value <> 0 Or CStr(ws.Cells(2, 3).value) <> "-" Then fails = fails + 1: Debug.Print "  FAIL  fila de la pagina 0 deberia mostrar P=0, PFN=-"
    If ws.Cells(17, 2).value <> 1 Or CStr(ws.Cells(17, 3).value) <> "7" Then fails = fails + 1: Debug.Print "  FAIL  fila de la pagina 15 deberia mostrar P=1, PFN=7"

    ' 3) The view updates when the table changes
    With PageTable(3)
        .Present = True: .pfn = 2: .Dirty = True: .Ref = True
    End With
    FrameOwner(2) = 3
    RefreshPageTableView
    If ws.Cells(5, 2).value <> 1 Or CStr(ws.Cells(5, 3).value) <> "2" Or ws.Cells(5, 4).value <> 1 Or ws.Cells(5, 5).value <> 1 Then
        fails = fails + 1: Debug.Print "  FAIL  la fila de la pagina 3 no se actualizo (esperado P=1 PFN=2 D=1 R=1)"
    End If

    ' 4) InitPageTable leaves everything clean again
    InitPageTable
    If PageTable(3).Present Or FrameOwner(2) <> -1 Or ws.Cells(5, 2).value <> 0 Then
        fails = fails + 1: Debug.Print "  FAIL  InitPageTable no limpio la pagina 3"
    End If

    Debug.Print IIf(fails = 0, "Test_PageTable: ALL OK", "Test_PageTable: " & fails & " FAIL(S)")
End Sub

Public Sub Test_Translate()
    Dim fails As Integer, p As Integer, ws As Worksheet
    InitPageTable
    Set ws = ThisWorkbook.Worksheets(SH_MEM)

    ' Page 13 loaded by hand into frame 4 (so it does not depend on HandlePageFault)
    With PageTable(13)
        .Present = True: .pfn = 4
    End With
    FrameOwner(4) = 13

    ' 1) Translate(&H1A5): VPN=13, offset=5, PADDR = 4*32 + 5 = 133
    p = Translate(&H1A5, False)
    If p <> 133 Then fails = fails + 1: Debug.Print "  FAIL  Translate(&H1A5) = " & p & " (esperado 133)"
    If CPU.paddr <> 133 Then fails = fails + 1: Debug.Print "  FAIL  CPU.PADDR = " & CPU.paddr & " (esperado 133)"
    If Not PageTable(13).Ref Then fails = fails + 1: Debug.Print "  FAIL  la lectura no puso Ref=1"
    If PageTable(13).Dirty Then fails = fails + 1: Debug.Print "  FAIL  la lectura puso Dirty=1"

    ' 2) A write sets Dirty=1
    p = Translate(&H1A5, True)
    If Not PageTable(13).Dirty Then fails = fails + 1: Debug.Print "  FAIL  la escritura no puso Dirty=1"

    ' 3) The table row (VPN 13 = row 15) is updated on every access
    If ws.Cells(15, 2).value <> 1 Or CStr(ws.Cells(15, 3).value) <> "4" Or ws.Cells(15, 4).value <> 1 Or ws.Cells(15, 5).value <> 1 Then
        fails = fails + 1: Debug.Print "  FAIL  la fila de la pagina 13 no muestra P=1 PFN=4 D=1 R=1"
    End If

    ' 4) Page 15 (I/O): &H1E0 -> frame 7, offset 0 -> 224 (&HE0)
    p = Translate(&H1E0, False)
    If p <> 224 Then fails = fails + 1: Debug.Print "  FAIL  Translate(&H1E0) = " & p & " (esperado 224)"

    ' 5) Out-of-range address: must give a clear error
    On Error Resume Next
    p = Translate(512, False)
    If Err.Number <> vbObjectError + 531 Then fails = fails + 1: Debug.Print "  FAIL  Translate(512) no dio el error de rango"
    Err.Clear
    On Error GoTo 0

    InitPageTable
    Debug.Print IIf(fails = 0, "Test_Translate: ALL OK", "Test_Translate: " & fails & " FAIL(S)")
End Sub

' Prepares a clean machine for the Page Fault tests:
' swap with one pattern per page (VMem(p*32) = 100+p), frames 0..6 at 0,
' clean page table, counters at 0 and the requested policy.
Private Sub PreparePFTest(ByVal policy As Integer)
    Dim p As Integer, i As Integer
    ClearVMem
    For p = 0 To NUM_VPAGES - 2: VMem(p * PAGE_SIZE) = 100 + p: Next p
    For i = 0 To IO_FRAME * PAGE_SIZE - 1: RAM(i) = 0: Next i
    InitPageTable
    Stats.PageFaults = 0
    ReplacePolicy = policy
    Clock = 0
End Sub

' Reads the first address of pages 0..n-1 (the clock advances 1 per access)
Private Sub LoadFirstPages(ByVal n As Integer)
    Dim p As Integer
    For p = 0 To n - 1
        Clock = Clock + 1
        Translate p * PAGE_SIZE, False
    Next p
End Sub

Public Sub Test_PageFault()
    Dim fails As Integer, p As Integer, ws As Worksheet
    Dim oldPolicy As Integer, oldClock As Long
    oldPolicy = ReplacePolicy: oldClock = Clock
    Set ws = ThisWorkbook.Worksheets(SH_MEM)

    ' ===== 1) Free frames: 7 pages -> 7 Page Faults, no eviction =====
    PreparePFTest POLICY_FIFO
    LoadFirstPages 7
    If Stats.PageFaults <> 7 Then fails = fails + 1: Debug.Print "  FAIL  1: Page Faults = " & Stats.PageFaults & " (esperado 7)"
    For p = 0 To 6
        If Not PageTable(p).Present Or PageTable(p).pfn <> p Or FrameOwner(p) <> p Then
            fails = fails + 1: Debug.Print "  FAIL  1: pagina " & p & " deberia estar en el marco " & p
        End If
        If RAM(p * PAGE_SIZE) <> 100 + p Then fails = fails + 1: Debug.Print "  FAIL  1: el marco " & p & " no recibio el bloque del swap"
    Next p
    If Not PageTable(IO_VPAGE).Present Or FrameOwner(IO_FRAME) <> IO_VPAGE Then fails = fails + 1: Debug.Print "  FAIL  1: la pagina 15 se movio del marco 7"

    ' ===== 2) FIFO with a DIRTY victim: page 0 goes out even if it was used recently =====
    Clock = Clock + 1: Translate 1, True            ' write to page 0 -> Dirty
    RAM(1) = 77                                     ' the data the program "wrote"
    Clock = Clock + 1: Translate 0, False           ' reading it again does NOT save it in FIFO
    Clock = Clock + 1: Translate 7 * PAGE_SIZE, False ' page 7: RAM full -> Page Fault
    If Stats.PageFaults <> 8 Then fails = fails + 1: Debug.Print "  FAIL  2: Page Faults = " & Stats.PageFaults & " (esperado 8)"
    If PageTable(0).Present Then fails = fails + 1: Debug.Print "  FAIL  2: FIFO deberia desalojar la pagina 0"
    If Not PageTable(7).Present Or PageTable(7).pfn <> 0 Or FrameOwner(0) <> 7 Then fails = fails + 1: Debug.Print "  FAIL  2: la pagina 7 deberia quedar en el marco 0"
    If VMem(1) <> 77 Then fails = fails + 1: Debug.Print "  FAIL  2: la victima sucia no se guardo en swap (VMem(1) = " & VMem(1) & ")"
    If RAM(0) <> 107 Then fails = fails + 1: Debug.Print "  FAIL  2: el marco 0 no recibio la pagina 7 (RAM(0) = " & RAM(0) & ")"
    If PageTable(7).Dirty Then fails = fails + 1: Debug.Print "  FAIL  2: una pagina recien cargada no debe estar sucia"
    If ws.Cells(2, 2).value <> 0 Or ws.Cells(9, 2).value <> 1 Or CStr(ws.Cells(9, 3).value) <> "0" Then
        fails = fails + 1: Debug.Print "  FAIL  2: la tabla visible no muestra P0 ausente y P7 en el marco 0"
    End If

    ' ===== 3) FIFO with a CLEAN victim: nothing is written to swap =====
    RAM(1 * PAGE_SIZE + 5) = 55                     ' change in RAM without marking Dirty
    Clock = Clock + 1: Translate 8 * PAGE_SIZE, False ' page 1 goes out (the oldest now)
    If PageTable(1).Present Then fails = fails + 1: Debug.Print "  FAIL  3: FIFO deberia desalojar la pagina 1"
    If VMem(1 * PAGE_SIZE + 5) <> 0 Then fails = fails + 1: Debug.Print "  FAIL  3: una victima limpia no debe guardarse en swap"
    If Stats.PageFaults <> 9 Then fails = fails + 1: Debug.Print "  FAIL  3: Page Faults = " & Stats.PageFaults & " (esperado 9)"

    ' ===== 4) LRU: page 0, used recently, is saved; page 1 goes out =====
    PreparePFTest POLICY_LRU
    LoadFirstPages 7
    Clock = Clock + 1: Translate 0, False           ' page 0 becomes the most recent
    Clock = Clock + 1: Translate 7 * PAGE_SIZE, False
    If PageTable(1).Present Then fails = fails + 1: Debug.Print "  FAIL  4: LRU deberia desalojar la pagina 1"
    If Not PageTable(0).Present Then fails = fails + 1: Debug.Print "  FAIL  4: LRU no debia desalojar la pagina 0"
    If Not PageTable(7).Present Or PageTable(7).pfn <> 1 Or FrameOwner(1) <> 7 Then fails = fails + 1: Debug.Print "  FAIL  4: la pagina 7 deberia quedar en el marco 1"
    If Stats.PageFaults <> 8 Then fails = fails + 1: Debug.Print "  FAIL  4: Page Faults = " & Stats.PageFaults & " (esperado 8)"

    ' ===== 5) No possible victim (all pinned): clear error and page 15 intact =====
    PreparePFTest POLICY_FIFO
    LoadFirstPages 7
    For p = 0 To 6: PageTable(p).Pinned = True: Next p
    On Error Resume Next
    Translate 7 * PAGE_SIZE, False
    If Err.Number <> vbObjectError + 530 Then fails = fails + 1: Debug.Print "  FAIL  5: no dio el error 530 (sin victima)"
    Err.Clear
    On Error GoTo 0
    If PageTable(7).Present Then fails = fails + 1: Debug.Print "  FAIL  5: la pagina 7 no deberia haberse cargado"
    If Not PageTable(IO_VPAGE).Present Or FrameOwner(IO_FRAME) <> IO_VPAGE Then fails = fails + 1: Debug.Print "  FAIL  5: la pagina 15 fue desalojada"

    ' ===== 6) Page out of range =====
    On Error Resume Next
    HandlePageFault 16
    If Err.Number <> vbObjectError + 533 Then fails = fails + 1: Debug.Print "  FAIL  6: HandlePageFault(16) no dio error"
    Err.Clear
    On Error GoTo 0

    ' Cleanup: leaves the machine as it was
    PreparePFTest oldPolicy
    ClearVMem
    ReplacePolicy = oldPolicy
    Clock = oldClock
    Debug.Print IIf(fails = 0, "Test_PageFault: ALL OK", "Test_PageFault: " & fails & " FAIL(S)")
End Sub
