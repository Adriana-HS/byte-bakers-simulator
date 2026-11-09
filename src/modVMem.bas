Attribute VB_Name = "modVMem"
Option Explicit
' =====================================================================
' modVMem - Virtual memory / swap area (K2)
'   VMem(0..511) = 16 pages x 32 bytes (the "disk" of the machine).
'   SwapIn  : swap -> RAM frame   (when a Page Fault is resolved)
'   SwapOut : RAM frame -> swap   (when a dirty page is evicted)
' The Memory sheet shows the 16 pages in a 16 x 32 grid in
' hexadecimal, starting at cell H1 (the page table uses A1:E17).
' =====================================================================

' First column of the grid: H = 8 (labels P0..P15), data in I..AN
Private Const GRID_LABEL_COL As Long = 8
Private Const GRID_DATA_COL As Long = 9

' ---------- Operations on virtual memory ----------

Public Sub ClearVMem()
    Dim i As Integer
    For i = 0 To VMEM_SIZE - 1: VMem(i) = 0: Next i
    RefreshVMemGrid
End Sub

' Copies bytes() into VMem starting at startAddr (the loader uses this)
Public Sub VMemWriteBlock(ByVal startAddr As Integer, bytes() As Integer)
    Dim i As Integer, n As Integer
    n = UBound(bytes) - LBound(bytes) + 1
    If startAddr < 0 Or startAddr + n > VMEM_SIZE Then
        Err.Raise vbObjectError + 541, "VMem", "El bloque no cabe en la memoria virtual (inicio " & startAddr & ", " & n & " bytes)"
    End If
    For i = LBound(bytes) To UBound(bytes)
        VMem(startAddr + i - LBound(bytes)) = bytes(i) And BYTE_MAX
    Next i
    RefreshVMemGrid
End Sub

' swap -> RAM frame: copies the 32 bytes of page vpn into frame pfn
Public Sub SwapIn(ByVal vpn As Integer, ByVal pfn As Integer)
    Dim i As Integer
    CheckSwapArgs vpn, pfn
    For i = 0 To PAGE_SIZE - 1
        RAM(pfn * PAGE_SIZE + i) = VMem(vpn * PAGE_SIZE + i)
    Next i
End Sub

' RAM frame -> swap: saves the 32 bytes of frame pfn into page vpn
Public Sub SwapOut(ByVal vpn As Integer, ByVal pfn As Integer)
    Dim i As Integer
    CheckSwapArgs vpn, pfn
    For i = 0 To PAGE_SIZE - 1
        VMem(vpn * PAGE_SIZE + i) = RAM(pfn * PAGE_SIZE + i)
    Next i
    RefreshVMemRow vpn
End Sub

Private Sub CheckSwapArgs(ByVal vpn As Integer, ByVal pfn As Integer)
    If vpn < 0 Or vpn > NUM_VPAGES - 1 Then
        Err.Raise vbObjectError + 540, "VMem", "Pagina virtual fuera de rango (0-" & NUM_VPAGES - 1 & "): " & vpn
    End If
    If pfn < 0 Or pfn > NUM_FRAMES - 1 Then
        Err.Raise vbObjectError + 540, "VMem", "Marco fuera de rango (0-" & NUM_FRAMES - 1 & "): " & pfn
    End If
End Sub

' ---------- 16 x 32 hexadecimal grid on the Memory sheet ----------

Private Function Hex2(ByVal v As Integer) As String
    Hex2 = Right$("0" & Hex$(v And BYTE_MAX), 2)
End Function

Private Function GetVMemSheet() As Worksheet
    On Error Resume Next
    Set GetVMemSheet = ThisWorkbook.Worksheets(SH_MEM)
    On Error GoTo 0
    If GetVMemSheet Is Nothing Then
        Set GetVMemSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetVMemSheet.Name = SH_MEM
    End If
End Function

' Draws the headers (offsets 00..1F, labels P0..P15) and the 512 bytes
Public Sub RefreshVMemGrid()
    Dim ws As Worksheet, p As Integer, o As Integer
    Dim hdr(0 To 0, 0 To PAGE_SIZE - 1) As String
    Dim lbl(0 To NUM_VPAGES - 1, 0 To 0) As String
    Dim body(0 To NUM_VPAGES - 1, 0 To PAGE_SIZE - 1) As String
    Dim lastCol As Long

    Set ws = GetVMemSheet()
    lastCol = GRID_DATA_COL + PAGE_SIZE - 1

    For o = 0 To PAGE_SIZE - 1: hdr(0, o) = Hex2(o): Next o
    For p = 0 To NUM_VPAGES - 1
        lbl(p, 0) = "P" & p
        For o = 0 To PAGE_SIZE - 1
            body(p, o) = Hex2(VMem(p * PAGE_SIZE + o))
        Next o
    Next p

    ' text format so Excel does not turn "10" into a number or "1E" into scientific notation
    ws.Range(ws.Cells(1, GRID_LABEL_COL), ws.Cells(NUM_VPAGES + 1, lastCol)).NumberFormat = "@"
    ws.Range(ws.Cells(1, GRID_DATA_COL), ws.Cells(1, lastCol)).value = hdr
    ws.Range(ws.Cells(2, GRID_LABEL_COL), ws.Cells(NUM_VPAGES + 1, GRID_LABEL_COL)).value = lbl
    ws.Range(ws.Cells(2, GRID_DATA_COL), ws.Cells(NUM_VPAGES + 1, lastCol)).value = body

    With ws.Range(ws.Cells(1, GRID_LABEL_COL), ws.Cells(NUM_VPAGES + 1, lastCol))
        .Font.Name = "Consolas"
        .Font.size = 9
        .HorizontalAlignment = xlCenter
    End With
    ' Header row in navy with bold cream text
    With ws.Range(ws.Cells(1, GRID_LABEL_COL), ws.Cells(1, lastCol))
        .Interior.Color = CLR_BG
        .Font.Color = CLR_TEXT
        .Font.bold = True
    End With
    ' Label column in navy with bold cream text
    With ws.Range(ws.Cells(2, GRID_LABEL_COL), ws.Cells(NUM_VPAGES + 1, GRID_LABEL_COL))
        .Interior.Color = CLR_BG
        .Font.Color = CLR_TEXT
        .Font.bold = True
    End With
    ' Data cells in cream with dark ink
    With ws.Range(ws.Cells(2, GRID_DATA_COL), ws.Cells(NUM_VPAGES + 1, lastCol))
        .Interior.Color = CLR_TEXT
        .Font.Color = CLR_INK
    End With
    With ws.Range(ws.Cells(1, GRID_LABEL_COL), ws.Cells(NUM_VPAGES + 1, lastCol)).Borders
        .LineStyle = xlContinuous
        .weight = xlThin
        .Color = CLR_BORDER
    End With
    ws.Cells(1, GRID_LABEL_COL).value = "VMem"
    ws.Columns(GRID_LABEL_COL).ColumnWidth = 6
    ws.Range(ws.Columns(GRID_DATA_COL), ws.Columns(lastCol)).ColumnWidth = 3.5
End Sub

' Redraws only the row of one page (used after a SwapOut)
Public Sub RefreshVMemRow(ByVal vpn As Integer)
    Dim ws As Worksheet, o As Integer, r As Long
    Dim rowData(0 To 0, 0 To PAGE_SIZE - 1) As String
    Set ws = GetVMemSheet()
    r = vpn + 2
    For o = 0 To PAGE_SIZE - 1: rowData(0, o) = Hex2(VMem(vpn * PAGE_SIZE + o)): Next o
    ws.Range(ws.Cells(r, GRID_DATA_COL), ws.Cells(r, GRID_DATA_COL + PAGE_SIZE - 1)).NumberFormat = "@"
    ws.Range(ws.Cells(r, GRID_DATA_COL), ws.Cells(r, GRID_DATA_COL + PAGE_SIZE - 1)).value = rowData
End Sub

' ---------- TESTS ----------
' Run with F5 (cursor inside the Sub) and watch the Immediate window (Ctrl+G)
Public Sub Test_VMem()
    Dim fails As Integer, i As Integer, ws As Worksheet, ok As Boolean
    Dim blk(0 To 2) As Integer
    Set ws = GetVMemSheet()

    ' Clean starting point: RAM frames 0..6 at zero (frame 7 is I/O, not touched)
    For i = 0 To 7 * PAGE_SIZE - 1: RAM(i) = 0: Next i

    ' 1) ClearVMem sets everything to 0
    VMem(5) = 9: VMem(511) = 1
    ClearVMem
    ok = True
    For i = 0 To VMEM_SIZE - 1
        If VMem(i) <> 0 Then ok = False
    Next i
    If Not ok Then fails = fails + 1: Debug.Print "  FAIL  ClearVMem no dejo VMem en 0"

    ' 2) VMemWriteBlock copies at the requested address
    blk(0) = 11: blk(1) = 22: blk(2) = 300
    VMemWriteBlock 64, blk
    If VMem(64) <> 11 Or VMem(65) <> 22 Or VMem(66) <> 44 Then fails = fails + 1: Debug.Print "  FAIL  VMemWriteBlock (esperado 11, 22, 44)"
    If VMem(67) <> 0 Then fails = fails + 1: Debug.Print "  FAIL  VMemWriteBlock escribio de mas"

    ' 3) SwapIn: swap page 3 -> RAM frame 5
    ClearVMem
    For i = 0 To PAGE_SIZE - 1: VMem(3 * PAGE_SIZE + i) = i + 1: Next i
    SwapIn 3, 5
    ok = True
    For i = 0 To PAGE_SIZE - 1
        If RAM(5 * PAGE_SIZE + i) <> i + 1 Then ok = False
    Next i
    If Not ok Then fails = fails + 1: Debug.Print "  FAIL  SwapIn 3,5 no copio la pagina al marco"
    If RAM(5 * PAGE_SIZE - 1) <> 0 Or RAM(6 * PAGE_SIZE) <> 0 Then fails = fails + 1: Debug.Print "  FAIL  SwapIn toco marcos vecinos"

    ' 4) SwapOut: RAM frame 2 -> swap page 7
    For i = 0 To PAGE_SIZE - 1: RAM(2 * PAGE_SIZE + i) = 200 + i: Next i
    SwapOut 7, 2
    ok = True
    For i = 0 To PAGE_SIZE - 1
        If VMem(7 * PAGE_SIZE + i) <> 200 + i Then ok = False
    Next i
    If Not ok Then fails = fails + 1: Debug.Print "  FAIL  SwapOut 7,2 no copio el marco a la pagina"
    If VMem(7 * PAGE_SIZE - 1) <> 0 Or VMem(8 * PAGE_SIZE) <> 0 Then fails = fails + 1: Debug.Print "  FAIL  SwapOut toco paginas vecinas"

    ' 5) Out-of-range arguments give a clear error
    On Error Resume Next
    SwapIn 16, 0
    If Err.Number <> vbObjectError + 540 Then fails = fails + 1: Debug.Print "  FAIL  SwapIn con pagina 16 no dio error"
    Err.Clear
    SwapOut 0, 8
    If Err.Number <> vbObjectError + 540 Then fails = fails + 1: Debug.Print "  FAIL  SwapOut con marco 8 no dio error"
    Err.Clear
    VMemWriteBlock 510, blk
    If Err.Number <> vbObjectError + 541 Then fails = fails + 1: Debug.Print "  FAIL  VMemWriteBlock que no cabe no dio error"
    Err.Clear
    On Error GoTo 0

    ' 6) The grid on the Memory sheet
    ClearVMem
    VMem(0) = &H1F: VMem(511) = &HAB: VMem(33) = &HA
    RefreshVMemGrid
    If CStr(ws.Cells(1, 9).value) <> "00" Or CStr(ws.Cells(1, 40).value) <> "1F" Then fails = fails + 1: Debug.Print "  FAIL  encabezados de offset (I1 = 00, AN1 = 1F)"
    If CStr(ws.Cells(2, 8).value) <> "P0" Or CStr(ws.Cells(17, 8).value) <> "P15" Then fails = fails + 1: Debug.Print "  FAIL  etiquetas de pagina (H2 = P0, H17 = P15)"
    If CStr(ws.Cells(2, 9).value) <> "1F" Then fails = fails + 1: Debug.Print "  FAIL  VMem(0) deberia verse como 1F"
    If CStr(ws.Cells(17, 40).value) <> "AB" Then fails = fails + 1: Debug.Print "  FAIL  VMem(511) deberia verse como AB"
    If CStr(ws.Cells(3, 10).value) <> "0A" Then fails = fails + 1: Debug.Print "  FAIL  VMem(33) deberia verse como 0A (pagina 1, offset 1)"

    ' Cleanup: leaves VMem and RAM frames 0..6 at 0
    ClearVMem
    For i = 0 To 7 * PAGE_SIZE - 1: RAM(i) = 0: Next i
    Debug.Print IIf(fails = 0, "Test_VMem: ALL OK", "Test_VMem: " & fails & " FAIL(S)")
End Sub

' For the PR screenshot: fills VMem with the pattern 00..FF (twice)
' and draws the grid. Then type ClearVMem in the Immediate window to leave it at zeros.
Public Sub Demo_VMemGrid()
    Dim i As Integer
    For i = 0 To VMEM_SIZE - 1: VMem(i) = i Mod 256: Next i
    RefreshVMemGrid
End Sub
