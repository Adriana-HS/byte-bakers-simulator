Attribute VB_Name = "modState"
' =====================================================================
' modState - Machine state contract of the Byte Bakers simulator
' Owner: Adriana  |  Task: #2 (A2)
'
' Every module reads and writes the machine through these types and
' globals: CPU, RAM, virtual memory, page table, caches, buses and
' statistics. Changing a type here changes the contract for everyone.
' =====================================================================
Option Explicit

' A bus with no active driver holds HIGH_Z (high impedance).
Public Const HIGH_Z As Integer = -1

' ---------------------------------------------------------------------
' CPU: 4 general registers + control registers + flags
' ---------------------------------------------------------------------
Public Type TCPU
    R(0 To 3) As Integer   ' AX, BX, CX, DX (8 bits: 0-255)
    PC As Integer          ' program counter, virtual (9 bits: 0-511)
    IR As Long             ' instruction register (16 bits: whole instruction)
    MAR As Integer         ' memory address register, virtual (9 bits: 0-511)
    PADDR As Integer       ' MMU output, physical (8 bits: 0-255)
    MDR As Integer         ' memory data register (8 bits: 0-255)
    TA As Integer          ' ALU temporary A (8 bits)
    TB As Integer          ' ALU temporary B (8 bits)
    ZF As Boolean          ' zero flag
    CF As Boolean          ' carry / borrow flag
    SF As Boolean          ' sign flag (bit 7 of the result)
    Halted As Boolean      ' True after HLT
End Type

' ---------------------------------------------------------------------
' Page table entry (one per virtual page, 16 in total)
' ---------------------------------------------------------------------
Public Type TPageEntry
    Present As Boolean     ' P: page is in RAM
    PFN As Integer         ' physical frame number 0-7 (valid when Present)
    Dirty As Boolean       ' D: page was written since it was loaded
    Ref As Boolean         ' R: page was referenced
    LoadedAt As Long       ' clock when it was loaded (FIFO)
    LastUse As Long        ' clock of the last access (LRU)
    Pinned As Boolean      ' never evicted (page 15, I/O)
End Type

' ---------------------------------------------------------------------
' Cache line (8 bytes). L1I and L1D have 1 line, L2 has 4.
' ---------------------------------------------------------------------
Public Type TCacheLine
    Valid As Boolean
    Tag As Integer         ' paddr \ LINE_SIZE
    Data(0 To 7) As Integer
    LastUse As Long        ' clock of the last access (LRU in L2)
End Type

' ---------------------------------------------------------------------
' System buses. Internal = inside the CPU, External = CPU <-> cache,
' RAM and I/O. Any field equal to HIGH_Z means nobody drives it.
' ---------------------------------------------------------------------
Public Type TBus
    IntAddr As Integer     ' internal address bus: virtual address (9 bits) or HIGH_Z
    ExtAddr As Integer     ' external address bus: physical address (8 bits) or HIGH_Z
    AddrDriver As String   ' who drives the address bus: "PC", "IR" or ""
    Data As Integer        ' data bus value (8 bits) or HIGH_Z
    DataDriver As String   ' who drives the data bus: "AX", "MDR", "KBD"... or ""
    Ctrl As Long           ' control word of the current micro-step (27 bits)
    Clk As Boolean         ' CLK line of the control bus (True during the pulse)
    Internal As Boolean    ' True = transfer on the CPU internal bus
End Type

' ---------------------------------------------------------------------
' Cache and paging statistics
' ---------------------------------------------------------------------
Public Type TStats
    L1IHit As Long
    L1IMiss As Long
    L1DHit As Long
    L1DMiss As Long
    L2Hit As Long
    L2Miss As Long
    PageFaults As Long
End Type

' ---------------------------------------------------------------------
' Global machine state
' ---------------------------------------------------------------------
Public CPU As TCPU
Public RAM(0 To 255) As Integer           ' physical memory: 8 frames x 32 B
Public VMem(0 To 511) As Integer          ' virtual memory / swap: 16 pages x 32 B
Public PageTable(0 To 15) As TPageEntry
Public FrameOwner(0 To 7) As Integer      ' VPN held by each frame, -1 = free
Public L1I As TCacheLine                  ' instruction cache (8 B)
Public L1D As TCacheLine                  ' data cache (8 B)
Public L2(0 To 3) As TCacheLine           ' unified cache (4 lines x 8 B = 32 B)
Public Bus As TBus
Public Stats As TStats
Public Clock As Long                      ' clock cycles since the last reset
Public ReplacePolicy As Integer           ' POLICY_FIFO or POLICY_LRU

' =====================================================================
' ResetMachine - puts the whole machine in its power-on state.
'   * CPU registers, flags and temporaries = 0, PC = 0
'   * RAM = 0, every cache line invalid, statistics and clock = 0
'   * every page P = 0, every frame free, except
'     page 15 (I/O) -> frame 7, present and pinned
'   * every bus in High-Z
' VMem is NOT cleared: it holds the loaded program (the loader clears it).
' ReplacePolicy is NOT changed: it is the user's choice.
' =====================================================================
Public Sub ResetMachine()
    Dim i As Integer
    Dim emptyStats As TStats

    ResetCPU                                  ' modCPU (#4)
    For i = 0 To RAM_SIZE - 1: RAM(i) = 0: Next i
    ResetPageTableState
    ResetCacheState
    ResetBusState
    Stats = emptyStats
    Clock = 0
End Sub

' Every page absent and every frame free, then pin the I/O page.
Private Sub ResetPageTableState()
    Dim i As Integer, emptyPTE As TPageEntry
    For i = 0 To NUM_VPAGES - 1: PageTable(i) = emptyPTE: Next i
    For i = 0 To NUM_FRAMES - 1: FrameOwner(i) = -1: Next i

    With PageTable(IO_VPAGE)
        .Present = True
        .PFN = IO_FRAME
        .Pinned = True
    End With
    FrameOwner(IO_FRAME) = IO_VPAGE
End Sub

' Invalidate L1I, L1D and the 4 lines of L2.
Private Sub ResetCacheState()
    Dim i As Integer, emptyLine As TCacheLine
    L1I = emptyLine
    L1D = emptyLine
    For i = 0 To L2_LINES - 1: L2(i) = emptyLine: Next i
End Sub

' No driver on any bus: every line in High-Z, control word and CLK at 0.
Private Sub ResetBusState()
    With Bus
        .IntAddr = HIGH_Z
        .ExtAddr = HIGH_Z
        .AddrDriver = ""
        .Data = HIGH_Z
        .DataDriver = ""
        .Ctrl = 0
        .Clk = False
        .Internal = False
    End With
End Sub

' =====================================================================
' Test_State - self-check of the state contract.
' Run it from the VBA editor (F5); results go to the Immediate window.
' =====================================================================
Public Sub Test_State()
    Dim i As Integer, fails As Integer

    ' Dirty the machine, then reset it
    CPU.PC = 300: CPU.R(R_AX) = 99: CPU.ZF = True: CPU.Halted = True
    RAM(10) = 77: PageTable(3).Present = True: FrameOwner(2) = 3
    L1I.Valid = True: L2(2).Valid = True
    Bus.Data = 5: Bus.DataDriver = "AX": Bus.Clk = True
    Stats.PageFaults = 4: Clock = 12
    ResetMachine

    fails = fails + Check("CPU cleared", CPU.PC = 0 And CPU.R(R_AX) = 0 And Not CPU.ZF And Not CPU.Halted)
    fails = fails + Check("RAM cleared", RAM(10) = 0)
    fails = fails + Check("Page 15 in frame 7", PageTable(IO_VPAGE).Present And PageTable(IO_VPAGE).PFN = IO_FRAME)
    fails = fails + Check("Page 15 pinned", PageTable(IO_VPAGE).Pinned)
    fails = fails + Check("Frame 7 owned by page 15", FrameOwner(IO_FRAME) = IO_VPAGE)
    For i = 0 To NUM_VPAGES - 2
        If PageTable(i).Present Then fails = fails + Check("Page " & i & " absent", False)
    Next i
    For i = 0 To NUM_FRAMES - 2
        If FrameOwner(i) <> -1 Then fails = fails + Check("Frame " & i & " free", False)
    Next i
    fails = fails + Check("Caches invalid", Not L1I.Valid And Not L1D.Valid And Not L2(2).Valid)
    fails = fails + Check("Data bus High-Z", Bus.Data = HIGH_Z And Bus.DataDriver = "")
    fails = fails + Check("Address buses High-Z", Bus.IntAddr = HIGH_Z And Bus.ExtAddr = HIGH_Z)
    fails = fails + Check("Control word and CLK off", Bus.Ctrl = 0 And Not Bus.Clk)
    fails = fails + Check("Stats and clock cleared", Stats.PageFaults = 0 And Clock = 0)
    fails = fails + Check("Memory map sizes", VMEM_SIZE = NUM_VPAGES * PAGE_SIZE And RAM_SIZE = NUM_FRAMES * PAGE_SIZE)
    fails = fails + Check("I/O cells inside frame 7", KBD_HEAD \ PAGE_SIZE = IO_FRAME And LCD_BASE + IO_BUF_SIZE - 1 = PADDR_MAX)
    fails = fails + Check("Virtual I/O = page 15", V_KBD_BASE \ PAGE_SIZE = IO_VPAGE And V_KBD_BASE Mod PAGE_SIZE = KBD_BASE Mod PAGE_SIZE)

    Debug.Print IIf(fails = 0, "Test_State: ALL OK", "Test_State: " & fails & " FAIL(S)")
End Sub

Private Function Check(ByVal what As String, ByVal passed As Boolean) As Integer
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    Check = IIf(passed, 0, 1)
End Function
