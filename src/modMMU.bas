Attribute VB_Name = "modMMU"
Option Explicit
' =====================================================================
' modMMU - Unidad de gestion de memoria (K3: tabla de paginas)
' Direccion virtual (9 bits) = [VPN 4 bits | offset 5 bits]
' Direccion fisica  (8 bits) = [PFN 3 bits | offset 5 bits]
' La pagina 15 es E/S: siempre en el marco 7, fija (Pinned), sin cache.
' =====================================================================

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

' Dibuja la tabla VPN / P / PFN / D / R en la hoja Memory (A1:E17)
Public Sub RefreshPageTableView()
    Dim ws As Worksheet, i As Integer, r As Long
    Set ws = GetMemorySheet()
    ws.Range("A1:E1").value = Array("VPN", "P", "PFN", "D", "R")
    ws.Range("A1:E1").Font.Bold = True
    For i = 0 To NUM_VPAGES - 1
        r = i + 2
        ws.Cells(r, 1).value = i
        ws.Cells(r, 2).value = IIf(PageTable(i).Present, 1, 0)
        ws.Cells(r, 3).value = IIf(PageTable(i).Present, CStr(PageTable(i).pfn), "-")
        ws.Cells(r, 4).value = IIf(PageTable(i).Dirty, 1, 0)
        ws.Cells(r, 5).value = IIf(PageTable(i).Ref, 1, 0)
        If PageTable(i).Present Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Color = CLR_HIT
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Interior.Pattern = xlNone
        End If
    Next i
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

' ---------- PRUEBA ----------
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
