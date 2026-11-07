Attribute VB_Name = "modMMU"
Option Explicit
' =====================================================================
' modMMU - Unidad de gestion de memoria (K3 tabla de paginas, K4 traduccion)
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
    Set ws = GetMemorySheet()
    ws.Range("A1:E1").value = Array("VPN", "P", "PFN", "D", "R")
    ws.Range("A1:E1").Font.Bold = True
    For i = 0 To NUM_VPAGES - 1
        RefreshPageTableRow i
    Next i
End Sub

' Redibuja solo la fila de una pagina (se llama en cada acceso)
Public Sub RefreshPageTableRow(ByVal vpn As Integer)
    Dim ws As Worksheet, r As Long
    Set ws = GetMemorySheet()
    r = vpn + 2
    ws.Cells(r, 1).value = vpn
    ws.Cells(r, 2).value = IIf(PageTable(vpn).Present, 1, 0)
    ws.Cells(r, 3).value = IIf(PageTable(vpn).Present, CStr(PageTable(vpn).pfn), "-")
    ws.Cells(r, 4).value = IIf(PageTable(vpn).Dirty, 1, 0)
    ws.Cells(r, 5).value = IIf(PageTable(vpn).Ref, 1, 0)
    If PageTable(vpn).Present Then
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = CLR_HIT
    Else
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Pattern = xlNone
    End If
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

' ---------- PRUEBAS ----------
' Ejecutar con F5 y mirar Inmediato (Ctrl+G)
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

    ' Pagina 13 cargada a mano en el marco 4 (HandlePageFault aun es un stub)
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
