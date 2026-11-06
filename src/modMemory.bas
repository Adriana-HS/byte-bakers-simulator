Attribute VB_Name = "modMemory"
Option Explicit
' =====================================================================
' modMemory - Unica puerta a memoria para la Unidad de Control (K4)
'   MemRead  : vaddr -> MMU -> (L1I | L1D | RAM si es la pagina 15)
'   MemWrite : vaddr -> MMU -> (L1D write-through | RAM si es pagina 15)
' La pagina 15 (E/S) NUNCA pasa por la cache.
' =====================================================================

Public Function MemRead(ByVal vaddr As Integer, ByVal isInstr As Boolean) As Integer
    Dim paddr As Integer
    paddr = Translate(vaddr, False)
    If vaddr \ PAGE_SIZE = IO_VPAGE Then
        MemRead = RAM(paddr)                 ' E/S: sin cache
    ElseIf isInstr Then
        MemRead = CacheReadL1I(paddr)        ' fetch
    Else
        MemRead = CacheReadL1D(paddr)        ' LOAD
    End If
End Function

Public Sub MemWrite(ByVal vaddr As Integer, ByVal value As Integer)
    Dim paddr As Integer, b As Integer
    b = value And BYTE_MAX                   ' la memoria guarda 8 bits
    paddr = Translate(vaddr, True)
    If vaddr \ PAGE_SIZE = IO_VPAGE Then
        RAM(paddr) = b                       ' E/S: sin cache
    Else
        CacheWriteL1D paddr, b               ' write-through
    End If
End Sub

' ---------- PRUEBA ----------
Public Sub Test_MemAccess()
    Dim fails As Integer
    InitPageTable
    With PageTable(13)
        .Present = True: .pfn = 4
    End With
    FrameOwner(4) = 13

    ' Lectura: pagina 13 offset 5 -> paddr 133
    RAM(133) = 77
    If MemRead(&H1A5, False) <> 77 Then fails = fails + 1: Debug.Print "  FAIL  MemRead de datos no devolvio 77"
    If MemRead(&H1A5, True) <> 77 Then fails = fails + 1: Debug.Print "  FAIL  MemRead de instruccion no devolvio 77"

    ' Escritura: la RAM siempre se actualiza y la pagina queda sucia
    MemWrite &H1A5, 99
    If RAM(133) <> 99 Then fails = fails + 1: Debug.Print "  FAIL  MemWrite no actualizo la RAM (" & RAM(133) & ")"
    If Not PageTable(13).Dirty Then fails = fails + 1: Debug.Print "  FAIL  MemWrite no puso Dirty=1"

    ' Escritura de mas de 8 bits: 300 -> 44
    MemWrite &H1A5, 300
    If RAM(133) <> 44 Then fails = fails + 1: Debug.Print "  FAIL  MemWrite(300) guardo " & RAM(133) & " (esperado 44)"

    ' Pagina 15 (E/S): &H1E3 -> RAM(&HE3), sin cache
    RAM(&HE3) = 5
    If MemRead(&H1E3, False) <> 5 Then fails = fails + 1: Debug.Print "  FAIL  lectura de E/S no devolvio 5"
    MemWrite &H1E3, 7
    If RAM(&HE3) <> 7 Then fails = fails + 1: Debug.Print "  FAIL  escritura de E/S no llego a RAM(&HE3)"
    If Not PageTable(IO_VPAGE).Ref Then fails = fails + 1: Debug.Print "  FAIL  el acceso a E/S no puso Ref=1"

    InitPageTable
    Debug.Print IIf(fails = 0, "Test_MemAccess: ALL OK", "Test_MemAccess: " & fails & " FAIL(S)")
End Sub
