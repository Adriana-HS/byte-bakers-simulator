Attribute VB_Name = "modBus"
' =====================================================================
' modBus - System buses with tri-state buffers (High-Z)
' Owner: Adriana  |  Task: #6 (A6)  |  Rubric: G1
'
' Internal address bus: 9-bit virtual (PC / IR -> MAR)
' External address bus: 8-bit physical (MMU -> cache / RAM / I/O)
' Data bus: 8 bits, bidirectional (registers / MDR <-> memory / KBD / LCD)
' A bus with no driver holds HIGH_Z (-1). Only ONE driver at a time.
' =====================================================================
Option Explicit

' ---------------------------------------------------------------------
' ReleaseBuses - end of micro-step: every line back to High-Z.
' ---------------------------------------------------------------------
Public Sub ReleaseBuses()
    Bus.Data = HIGH_Z
    Bus.IntAddr = HIGH_Z
    Bus.ExtAddr = HIGH_Z
    Bus.DataDriver = ""
    Bus.AddrDriver = ""
    Bus.Ctrl = 0
    Bus.Internal = False
End Sub

' ---------------------------------------------------------------------
' DriveDataBus - one component Enables its buffer onto the data bus.
' isInternal = True  -> CPU internal data path (wire_INT_DATA)
' isInternal = False -> external path CPU <-> memory / I/O (wire_DATA)
' ---------------------------------------------------------------------
Public Sub DriveDataBus(ByVal source As String, ByVal value As Integer, ByVal isInternal As Boolean)
    If Bus.Data <> HIGH_Z Then
        Err.Raise vbObjectError + 513, "modBus", _
            "Bus collision: " & Bus.DataDriver & " is already driving and " & source & " tried to drive."
    End If
    Bus.Data = Wrap8(CLng(value))
    Bus.DataDriver = source
    Bus.Internal = isInternal
    UI_Highlight PFX_BUF & source, CLR_ACTIVE
    UI_Highlight IIf(isInternal, PFX_WIRE & "INT_DATA", PFX_WIRE & "DATA"), CLR_ACTIVE
End Sub

' ---------------------------------------------------------------------
' DriveIntAddrBus - 9-bit virtual address (PC or IR address field -> MAR).
' Never carries data: PC -> MAR stays on the address bus (G1 / rubric 3.6).
' ---------------------------------------------------------------------
Public Sub DriveIntAddrBus(ByVal source As String, ByVal vaddr As Integer)
    If Bus.IntAddr <> HIGH_Z Then
        Err.Raise vbObjectError + 515, "modBus", _
            "Address bus collision: " & Bus.AddrDriver & " and " & source
    End If
    Bus.IntAddr = Wrap9(CLng(vaddr))
    Bus.AddrDriver = source
    UI_Highlight PFX_BUF & source, CLR_ACTIVE
    UI_Highlight PFX_WIRE & "INT_ADDR", CLR_ACTIVE
End Sub

' ---------------------------------------------------------------------
' DriveExtAddrBus - 8-bit physical address after the MMU (PADDR).
' ---------------------------------------------------------------------
Public Sub DriveExtAddrBus(ByVal paddr As Integer)
    Bus.ExtAddr = Wrap8(CLng(paddr))
    UI_Highlight PFX_WIRE & "ADDR", CLR_ACTIVE
End Sub

' ---------------------------------------------------------------------
' LatchFromBus - a register Loads (Ld_*) whatever is on the data bus.
' ---------------------------------------------------------------------
Public Function LatchFromBus(ByVal target As String) As Integer
    If Bus.Data = HIGH_Z Then
        Err.Raise vbObjectError + 514, "modBus", _
            target & " tried to read a bus in High-Z."
    End If
    UI_Highlight PFX_REG & target, CLR_LOAD
    LatchFromBus = Bus.Data
End Function

' =====================================================================
' Test_Bus - self-check of High-Z, collisions, widths and PC->MAR path.
' Returns the number of failed checks (0 = ALL OK).
' =====================================================================
Public Function Test_Bus() As Long
    Dim fails As Long
    Dim raised As Boolean
    Dim v As Integer

    ResetCPU
    ReleaseBuses

    ' --- Release leaves every bus in High-Z ---
    fails = fails + BusCheck("ReleaseBuses: data High-Z", Bus.Data = HIGH_Z And Bus.DataDriver = "")
    fails = fails + BusCheck("ReleaseBuses: int/ext addr High-Z", _
        Bus.IntAddr = HIGH_Z And Bus.ExtAddr = HIGH_Z And Bus.AddrDriver = "")

    ' --- Data bus drive + 8-bit wrap ---
    DriveDataBus "AX", 300, True
    fails = fails + BusCheck("DriveDataBus wraps 8 bits (300 -> 44)", Bus.Data = 44 And Bus.DataDriver = "AX" And Bus.Internal)
    ReleaseBuses

    ' --- Data bus collision ---
    DriveDataBus "AX", 10, True
    raised = False
    On Error Resume Next
    DriveDataBus "BX", 20, True
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    fails = fails + BusCheck("Second data driver raises Bus collision", raised)
    ReleaseBuses

    ' --- Internal address 9 bits; PC -> MAR never uses the data bus ---
    DriveIntAddrBus "PC", 100
    fails = fails + BusCheck("DriveIntAddrBus stores PC on IntAddr", Bus.IntAddr = 100 And Bus.AddrDriver = "PC")
    fails = fails + BusCheck("PC->MAR does not touch the data bus", Bus.Data = HIGH_Z And Bus.DataDriver = "")
    ReleaseBuses
    DriveIntAddrBus "PC", 512
    fails = fails + BusCheck("IntAddr wraps 9 bits (512 -> 0)", Bus.IntAddr = 0)
    ReleaseBuses
    DriveIntAddrBus "PC", 511
    raised = False
    On Error Resume Next
    DriveIntAddrBus "IR", 40
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    fails = fails + BusCheck("Second int-addr driver raises collision", raised)
    ReleaseBuses

    ' --- External address 8 bits (physical) ---
    DriveExtAddrBus 300
    fails = fails + BusCheck("ExtAddr wraps 8 bits (300 -> 44)", Bus.ExtAddr = 44)
    ReleaseBuses

    ' --- Bidirectional external data: MDR -> mem and mem -> MDR ---
    DriveDataBus "MDR", &HAB, False
    fails = fails + BusCheck("External write path MDR drives data bus", _
        Bus.Data = &HAB And Bus.DataDriver = "MDR" And Not Bus.Internal)
    ReleaseBuses
    DriveDataBus "MEM", &HCD, False
    fails = fails + BusCheck("External read path MEM drives data bus", _
        Bus.Data = &HCD And Bus.DataDriver = "MEM" And Not Bus.Internal)
    v = LatchFromBus("MDR")
    fails = fails + BusCheck("LatchFromBus reads external data into MDR path", v = &HCD)
    ReleaseBuses

    ' --- Latch on High-Z must fail ---
    raised = False
    On Error Resume Next
    v = LatchFromBus("AX")
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    fails = fails + BusCheck("LatchFromBus on High-Z raises an error", raised)

    Debug.Print IIf(fails = 0, "Test_Bus: ALL OK", "Test_Bus: " & fails & " FAIL(S)")
    Test_Bus = fails
End Function

Private Function BusCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    BusCheck = IIf(passed, 0, 1)
End Function
