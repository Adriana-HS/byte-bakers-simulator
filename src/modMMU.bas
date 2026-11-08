Attribute VB_Name = "modMMU"
Option Explicit
' =====================================================================
' modMMU - Unidad de gestion de memoria
'   K3: tabla de paginas | K4: traduccion | K5: Page Fault (FIFO / LRU)
' Direccion virtual (9 bits) = [VPN 4 bits | offset 5 bits]
' Direccion fisica  (8 bits) = [PFN 3 bits | offset 5 bits]
' La pagina 15 es E/S: siempre en el marco 7, fija (Pinned), sin cache.
' =====================================================================

' ---------- K3: TABLA DE PAGINAS ----------

' Deja la tabla de paginas con todas las paginas ausentes (P=0),
' excepto la 15, que queda en el marco 7 y fija.
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

' Dibuja la tabla VPN / P / PFN / D / R completa en la hoja Memory (A1:E17)
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

' Redibuja solo la fila de una pagina (se llama en cada acceso).
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

' ---------- K4: TRADUCCION VIRTUAL -> FISICA ----------

' Traduce una direccion virtual (0..511) a fisica (0..255).
'  - Si la pagina esta ausente (P=0) llama a HandlePageFault (K5) antes.
'  - Marca Ref=1 en cada acceso y Dirty=1 si es escritura.
'  - Deja el resultado en CPU.PADDR y actualiza la fila de la tabla.
Public Function Translate(ByVal vaddr As Integer, ByVal isWrite As Boolean) As Integer
    Dim vpn As Integer, offset As Integer, paddr As Integer
    If vaddr < 0 Or vaddr > VADDR_MAX Then
        Err.Raise vbObjectError + 531, "MMU", "Direccion virtual fuera de rango (0-" & VADDR_MAX & "): " & vaddr
    End If
    vpn = vaddr \ PAGE_SIZE
    offset = vaddr Mod PAGE_SIZE
    ShowTranslation vpn, offset                      ' muestra la traduccion en vivo (K10)
    If Not PageTable(vpn).Present Then
        HandlePageFault vpn                          ' K5: carga la pagina
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

' Carga la pagina vpn en un marco de RAM:
'  1) usa un marco libre si hay; si no, elige victima (FIFO o LRU segun ReplacePolicy)
'  2) si la victima esta sucia (Dirty=1) la guarda en swap
'  3) invalida las lineas de cache del marco (L1I, L1D, L2)
'  4) copia la pagina desde swap y marca P=1
' La pagina 15 (E/S, Pinned) nunca es victima.
Public Sub HandlePageFault(ByVal vpn As Integer)
    Dim pfn As Integer, victim As Integer
    If vpn < 0 Or vpn > NUM_VPAGES - 1 Then
        Err.Raise vbObjectError + 533, "MMU", "Pagina virtual fuera de rango (0-" & NUM_VPAGES - 1 & "): " & vpn
    End If
    If PageTable(vpn).Present Then Exit Sub          ' ya esta cargada: no hay fallo

    Stats.PageFaults = Stats.PageFaults + 1
    UI_PageFaultFlash vpn                            ' entrada roja (K10)

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
        InvalidateFrame pfn                          ' K8: invalida L1I, L1D y L2 de ese marco
        RefreshPageTableRow victim
    End If

    SwapIn vpn, pfn
    UI_AnimateSwapIn vpn, pfn                        ' bloque swap -> marco (K10)

    With PageTable(vpn)
        .Present = True
        .pfn = pfn
        .Dirty = False
        .Ref = False
        .LoadedAt = Clock                            ' para FIFO
        .LastUse = Clock                             ' para LRU
    End With
    FrameOwner(pfn) = vpn
    RefreshPageTableRow vpn
End Sub

' Primer marco libre, o -1 si la RAM esta llena
Private Function FindFreeFrame() As Integer
    Dim i As Integer
    For i = 0 To NUM_FRAMES - 1
        If FrameOwner(i) = -1 Then FindFreeFrame = i: Exit Function
    Next i
    FindFreeFrame = -1
End Function

' Elige la victima: FIFO = menor LoadedAt, LRU = menor LastUse.
' Solo paginas presentes y NO fijas (la 15 nunca se desaloja).
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

' ---------- PRUEBAS ----------
' Ejecutar con F5 (cursor dentro del Sub) y mirar Inmediato (Ctrl+G)
Public Sub Test_PageTable()
    Dim i As Integer, fails As Integer, ws As Worksheet

    ' 1) Estado inicial en memoria
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

    ' 2) Lo que se ve en la hoja Memory
    Set ws = ThisWorkbook.Worksheets(SH_MEM)
    If ws.Cells(1, 1).value <> "VPN" Then fails = fails + 1: Debug.Print "  FAIL  falta el encabezado VPN"
    If ws.Cells(2, 2).value <> 0 Or CStr(ws.Cells(2, 3).value) <> "-" Then fails = fails + 1: Debug.Print "  FAIL  fila de la pagina 0 deberia mostrar P=0, PFN=-"
    If ws.Cells(17, 2).value <> 1 Or CStr(ws.Cells(17, 3).value) <> "7" Then fails = fails + 1: Debug.Print "  FAIL  fila de la pagina 15 deberia mostrar P=1, PFN=7"

    ' 3) La vista se actualiza cuando cambia la tabla
    With PageTable(3)
        .Present = True: .pfn = 2: .Dirty = True: .Ref = True
    End With
    FrameOwner(2) = 3
    RefreshPageTableView
    If ws.Cells(5, 2).value <> 1 Or CStr(ws.Cells(5, 3).value) <> "2" Or ws.Cells(5, 4).value <> 1 Or ws.Cells(5, 5).value <> 1 Then
        fails = fails + 1: Debug.Print "  FAIL  la fila de la pagina 3 no se actualizo (esperado P=1 PFN=2 D=1 R=1)"
    End If

    ' 4) InitPageTable deja todo limpio otra vez
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

    ' Pagina 13 cargada a mano en el marco 4 (asi no depende de HandlePageFault)
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

    ' 2) Una escritura pone Dirty=1
    p = Translate(&H1A5, True)
    If Not PageTable(13).Dirty Then fails = fails + 1: Debug.Print "  FAIL  la escritura no puso Dirty=1"

    ' 3) La fila de la tabla (VPN 13 = fila 15) se actualiza en cada acceso
    If ws.Cells(15, 2).value <> 1 Or CStr(ws.Cells(15, 3).value) <> "4" Or ws.Cells(15, 4).value <> 1 Or ws.Cells(15, 5).value <> 1 Then
        fails = fails + 1: Debug.Print "  FAIL  la fila de la pagina 13 no muestra P=1 PFN=4 D=1 R=1"
    End If

    ' 4) Pagina 15 (E/S): &H1E0 -> marco 7, offset 0 -> 224 (&HE0)
    p = Translate(&H1E0, False)
    If p <> 224 Then fails = fails + 1: Debug.Print "  FAIL  Translate(&H1E0) = " & p & " (esperado 224)"

    ' 5) Direccion fuera de rango: debe dar error claro
    On Error Resume Next
    p = Translate(512, False)
    If Err.Number <> vbObjectError + 531 Then fails = fails + 1: Debug.Print "  FAIL  Translate(512) no dio el error de rango"
    Err.Clear
    On Error GoTo 0

    InitPageTable
    Debug.Print IIf(fails = 0, "Test_Translate: ALL OK", "Test_Translate: " & fails & " FAIL(S)")
End Sub

' Prepara una maquina limpia para las pruebas de Page Fault:
' swap con un patron por pagina (VMem(p*32) = 100+p), marcos 0..6 en 0,
' tabla de paginas limpia, contadores en 0 y la politica pedida.
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

' Lee la primera direccion de las paginas 0..n-1 (el reloj avanza 1 por acceso)
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

    ' ===== 1) Marcos libres: 7 paginas -> 7 Page Faults, sin desalojar =====
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

    ' ===== 2) FIFO con victima SUCIA: sale la pagina 0 aunque se uso hace poco =====
    Clock = Clock + 1: Translate 1, True            ' escritura en pagina 0 -> Dirty
    RAM(1) = 77                                     ' el dato que "escribio" el programa
    Clock = Clock + 1: Translate 0, False           ' leerla de nuevo NO la salva en FIFO
    Clock = Clock + 1: Translate 7 * PAGE_SIZE, False ' pagina 7: RAM llena -> Page Fault
    If Stats.PageFaults <> 8 Then fails = fails + 1: Debug.Print "  FAIL  2: Page Faults = " & Stats.PageFaults & " (esperado 8)"
    If PageTable(0).Present Then fails = fails + 1: Debug.Print "  FAIL  2: FIFO deberia desalojar la pagina 0"
    If Not PageTable(7).Present Or PageTable(7).pfn <> 0 Or FrameOwner(0) <> 7 Then fails = fails + 1: Debug.Print "  FAIL  2: la pagina 7 deberia quedar en el marco 0"
    If VMem(1) <> 77 Then fails = fails + 1: Debug.Print "  FAIL  2: la victima sucia no se guardo en swap (VMem(1) = " & VMem(1) & ")"
    If RAM(0) <> 107 Then fails = fails + 1: Debug.Print "  FAIL  2: el marco 0 no recibio la pagina 7 (RAM(0) = " & RAM(0) & ")"
    If PageTable(7).Dirty Then fails = fails + 1: Debug.Print "  FAIL  2: una pagina recien cargada no debe estar sucia"
    If ws.Cells(2, 2).value <> 0 Or ws.Cells(9, 2).value <> 1 Or CStr(ws.Cells(9, 3).value) <> "0" Then
        fails = fails + 1: Debug.Print "  FAIL  2: la tabla visible no muestra P0 ausente y P7 en el marco 0"
    End If

    ' ===== 3) FIFO con victima LIMPIA: no se escribe en swap =====
    RAM(1 * PAGE_SIZE + 5) = 55                     ' cambio en RAM sin marcar Dirty
    Clock = Clock + 1: Translate 8 * PAGE_SIZE, False ' sale la pagina 1 (la mas vieja ahora)
    If PageTable(1).Present Then fails = fails + 1: Debug.Print "  FAIL  3: FIFO deberia desalojar la pagina 1"
    If VMem(1 * PAGE_SIZE + 5) <> 0 Then fails = fails + 1: Debug.Print "  FAIL  3: una victima limpia no debe guardarse en swap"
    If Stats.PageFaults <> 9 Then fails = fails + 1: Debug.Print "  FAIL  3: Page Faults = " & Stats.PageFaults & " (esperado 9)"

    ' ===== 4) LRU: la pagina 0, usada hace poco, se salva; sale la 1 =====
    PreparePFTest POLICY_LRU
    LoadFirstPages 7
    Clock = Clock + 1: Translate 0, False           ' pagina 0 pasa a ser la mas reciente
    Clock = Clock + 1: Translate 7 * PAGE_SIZE, False
    If PageTable(1).Present Then fails = fails + 1: Debug.Print "  FAIL  4: LRU deberia desalojar la pagina 1"
    If Not PageTable(0).Present Then fails = fails + 1: Debug.Print "  FAIL  4: LRU no debia desalojar la pagina 0"
    If Not PageTable(7).Present Or PageTable(7).pfn <> 1 Or FrameOwner(1) <> 7 Then fails = fails + 1: Debug.Print "  FAIL  4: la pagina 7 deberia quedar en el marco 1"
    If Stats.PageFaults <> 8 Then fails = fails + 1: Debug.Print "  FAIL  4: Page Faults = " & Stats.PageFaults & " (esperado 8)"

    ' ===== 5) Sin victima posible (todas fijas): error claro y la 15 intacta =====
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

    ' ===== 6) Pagina fuera de rango =====
    On Error Resume Next
    HandlePageFault 16
    If Err.Number <> vbObjectError + 533 Then fails = fails + 1: Debug.Print "  FAIL  6: HandlePageFault(16) no dio error"
    Err.Clear
    On Error GoTo 0

    ' Limpieza: deja la maquina como estaba
    PreparePFTest oldPolicy
    ClearVMem
    ReplacePolicy = oldPolicy
    Clock = oldClock
    Debug.Print IIf(fails = 0, "Test_PageFault: ALL OK", "Test_PageFault: " & fails & " FAIL(S)")
End Sub
