Attribute VB_Name = "modControlUnit"
' =====================================================================
' modControlUnit - Fetch / decode / execute micro-operations (G1, G5)
' Owner: Adriana  |  Task: #8 (A8)
'
' StepMicro = one micro-step (CLK pulse + control word + bus transfer).
' RunProgram loops StepMicro until HLT or Pause (RunFlag=False).
' Fetch F1-F6 via MMU/L1I (MemRead isInstr:=True). Execute covers
' every ISA opcode. Each step is logged on the Log sheet.
' =====================================================================
Option Explicit

Public MicroStep As Integer      ' 1..6 fetch, 7.. execute phase index+6
Public RunFlag As Boolean
Public InstructionIndex As Long  ' 1-based index of current assembly instruction

Private curOp As Integer
Private curReg As Integer
Private curOperand As Integer
Private logRow As Long
Private lastDesc As String
Private lastFaults As Long
Private lastL1IH As Long
Private lastL1IM As Long
Private lastL1DH As Long
Private lastL1DM As Long
Private lastL2H As Long
Private lastL2M As Long
Private Const LOG_COLS As Integer = 28
Private Const LOG_FIRST As Long = 4
Private Const LOG_BLANK_TO As Long = 24

' =====================================================================
' Public entry points (called from modDashboard buttons)
' =====================================================================

Public Sub StepMicro()
    If CPU.Halted Then
        RunFlag = False
        Exit Sub
    End If
    If MicroStep < 1 Then MicroStep = 1

    Clock = Clock + 1
    Bus.Clk = True
    ReleaseBuses

    On Error GoTo EH
    If MicroStep <= 6 Then
        If MicroStep = 1 Then InstructionIndex = InstructionIndex + 1
        DoFetch MicroStep
        MicroStep = MicroStep + 1
        If MicroStep = 7 Then
            Decode CPU.IR, curOp, curReg, curOperand
        End If
    Else
        If DoExecute(MicroStep - 6) Then
            MicroStep = 1
        Else
            MicroStep = MicroStep + 1
        End If
    End If

    On Error Resume Next
    RefreshAll
    On Error GoTo 0
    Bus.Clk = False
    Exit Sub
EH:
    Bus.Clk = False
    RunFlag = False
    Debug.Print "StepMicro ERROR: " & Err.Number & " " & Err.Description
    On Error Resume Next
    RefreshAll
End Sub

Public Sub RunProgram()
    RunFlag = True
    Do While RunFlag And Not CPU.Halted
        StepMicro
        UI_Wait
        DoEvents
    Loop
End Sub

Public Sub ResetControlUnit()
    MicroStep = 0
    RunFlag = False
    curOp = 0
    curReg = 0
    curOperand = 0
    lastDesc = ""
    logRow = 1
    InstructionIndex = 0
    ClearLogSheet
End Sub

' =====================================================================
' Emit - validate control word, publish on Bus, log + soft CW paint
' =====================================================================
Public Function CurrentMicroText() As String
    Dim n As Integer
    If MicroStep >= 1 And MicroStep <= 6 Then
        n = MicroStep
        Select Case n
            Case 1, 4: CurrentMicroText = "F" & n & "  PC -> MAR"
            Case 2, 5: CurrentMicroText = "F" & n & "  M[MAR] -> MDR"
            Case 3: CurrentMicroText = "F3  MDR -> IR, PC+1"
            Case 6: CurrentMicroText = "F6  MDR -> IR, PC+1"
        End Select
    ElseIf MicroStep = 7 Then
        If Len(lastDesc) > 0 Then
            CurrentMicroText = "DEC  " & lastDesc
        Else
            CurrentMicroText = "DECODE"
        End If
    ElseIf MicroStep > 7 Then
        CurrentMicroText = "E" & (MicroStep - 7) & "  " & lastDesc
    Else
        CurrentMicroText = "idle"
    End If
End Function

Private Sub Emit(ByVal cw As Long, ByVal description As String)
    Dim msg As String
    msg = ValidateControlWord(cw)
    If Len(msg) > 0 Then
        Err.Raise vbObjectError + 520, "modControlUnit", msg
    End If
    Bus.Ctrl = cw
    Bus.Clk = True
    lastDesc = description
    LogStep cw, description
    AnimateControlWord cw, description, description
    On Error Resume Next
    UI_SetText "lbl_Transfer", description
    On Error GoTo 0
End Sub

Private Sub DoFetch(ByVal t As Integer)
    Select Case t
        Case 1, 4
            Emit SIG_EN_PC Or SIG_LD_MAR, "PC -> MAR (addr bus)"
            DriveIntAddrBus "PC", CPU.PC
            SetMAR CPU.PC
        Case 2, 5
            Emit SIG_MEM_R Or SIG_LD_MDR, "M[MAR] -> MDR (fetch)"
            SetMDR MemRead(CPU.MAR, True)
            DriveExtAddrBus CPU.PADDR
            DriveDataBus "MEM", CPU.MDR, False
        Case 3
            Emit SIG_EN_MDR Or SIG_LD_IR Or SIG_INC_PC, "MDR -> IR (hi), PC+1"
            DriveDataBus "MDR", CPU.MDR, True
            SetIR CLng(CPU.MDR) * 256
            IncPC
        Case 6
            Emit SIG_EN_MDR Or SIG_LD_IR Or SIG_INC_PC, "MDR -> IR (lo), PC+1"
            DriveDataBus "MDR", CPU.MDR, True
            SetIR (CPU.IR And &HFF00&) Or (CPU.MDR And BYTE_MAX)
            IncPC
    End Select
End Sub

' Returns True when the current instruction has finished execute phase.
Private Function DoExecute(ByVal t As Integer) As Boolean
    Dim src As Integer
    Dim r As Integer
    Dim taken As Boolean

    DoExecute = False

    Select Case curOp
        Case OP_NOP
            Emit 0, "NOP"
            DoExecute = True

        Case OP_HLT
            CPU.Halted = True
            RunFlag = False
            Emit 0, "HLT"
            DoExecute = True

        Case OP_MOV_RI
            Emit SIG_EN_IR_D Or SIG_LD_REG(curReg), "IR(imm) -> " & RegName(curReg)
            DriveDataBus "IR", curOperand And BYTE_MAX, True
            SetReg curReg, curOperand And BYTE_MAX
            DoExecute = True

        Case OP_MOV_RR
            src = curOperand And 3
            Emit SIG_EN_REG(src) Or SIG_LD_REG(curReg), RegName(src) & " -> " & RegName(curReg)
            DriveDataBus RegName(src), GetReg(src), True
            SetReg curReg, GetReg(src)
            DoExecute = True

        Case OP_LOAD
            Select Case t
                Case 1
                    Emit SIG_EN_IR_A Or SIG_LD_MAR, "IR(addr) -> MAR"
                    DriveIntAddrBus "IR", curOperand
                    SetMAR curOperand
                Case 2
                    Emit SIG_MEM_R Or SIG_LD_MDR, "M[MAR] -> MDR (data)"
                    SetMDR MemRead(CPU.MAR, False)
                    DriveExtAddrBus CPU.PADDR
                    DriveDataBus "MEM", CPU.MDR, False
                Case 3
                    Emit SIG_EN_MDR Or SIG_LD_REG(curReg), "MDR -> " & RegName(curReg)
                    DriveDataBus "MDR", CPU.MDR, True
                    SetReg curReg, CPU.MDR
                    DoExecute = True
            End Select

        Case OP_STORE
            Select Case t
                Case 1
                    Emit SIG_EN_IR_A Or SIG_LD_MAR, "IR(addr) -> MAR"
                    DriveIntAddrBus "IR", curOperand
                    SetMAR curOperand
                Case 2
                    Emit SIG_EN_REG(curReg) Or SIG_LD_MDR, RegName(curReg) & " -> MDR"
                    DriveDataBus RegName(curReg), GetReg(curReg), True
                    SetMDR GetReg(curReg)
                Case 3
                    Emit SIG_EN_MDR Or SIG_MEM_W, "MDR -> M[MAR]"
                    DriveExtAddrBus CPU.PADDR
                    DriveDataBus "MDR", CPU.MDR, False
                    MemWrite CPU.MAR, CPU.MDR
                    DoExecute = True
            End Select

        Case OP_ADD_RR, OP_SUB_RR, OP_AND, OP_OR, OP_XOR, OP_CMP_RR
            Select Case t
                Case 1
                    Emit SIG_EN_REG(curReg) Or SIG_LD_TA, RegName(curReg) & " -> TA"
                    DriveDataBus RegName(curReg), GetReg(curReg), True
                    CPU.TA = GetReg(curReg)
                Case 2
                    src = curOperand And 3
                    Emit SIG_EN_REG(src) Or SIG_LD_TB, RegName(src) & " -> TB"
                    DriveDataBus RegName(src), GetReg(src), True
                    CPU.TB = GetReg(src)
                Case 3
                    r = AluExecute(curOp)
                    If curOp = OP_CMP_RR Then
                        Emit SIG_EN_ALU Or SIG_LD_FLAGS, "ALU -> FLAGS (CMP)"
                    Else
                        Emit SIG_EN_ALU Or SIG_LD_REG(curReg) Or SIG_LD_FLAGS, _
                            "ALU -> " & RegName(curReg)
                        DriveDataBus "ALU", r, True
                        SetReg curReg, r
                    End If
                    DoExecute = True
            End Select

        Case OP_ADD_RI, OP_SUB_RI, OP_CMP_RI
            Select Case t
                Case 1
                    Emit SIG_EN_REG(curReg) Or SIG_LD_TA, RegName(curReg) & " -> TA"
                    DriveDataBus RegName(curReg), GetReg(curReg), True
                    CPU.TA = GetReg(curReg)
                Case 2
                    Emit SIG_EN_IR_D Or SIG_LD_TB, "IR(imm) -> TB"
                    DriveDataBus "IR", curOperand And BYTE_MAX, True
                    CPU.TB = curOperand And BYTE_MAX
                Case 3
                    r = AluExecute(curOp)
                    If curOp = OP_CMP_RI Then
                        Emit SIG_EN_ALU Or SIG_LD_FLAGS, "ALU -> FLAGS (CMP)"
                    Else
                        Emit SIG_EN_ALU Or SIG_LD_REG(curReg) Or SIG_LD_FLAGS, _
                            "ALU -> " & RegName(curReg)
                        DriveDataBus "ALU", r, True
                        SetReg curReg, r
                    End If
                    DoExecute = True
            End Select

        Case OP_INC, OP_DEC
            Select Case t
                Case 1
                    Emit SIG_EN_REG(curReg) Or SIG_LD_TA, RegName(curReg) & " -> TA"
                    DriveDataBus RegName(curReg), GetReg(curReg), True
                    CPU.TA = GetReg(curReg)
                    CPU.TB = 0
                Case 2
                    r = AluExecute(curOp)
                    Emit SIG_EN_ALU Or SIG_LD_REG(curReg) Or SIG_LD_FLAGS, _
                        "ALU -> " & RegName(curReg)
                    DriveDataBus "ALU", r, True
                    SetReg curReg, r
                    DoExecute = True
            End Select

        Case OP_JMP
            Emit SIG_EN_IR_A Or SIG_LD_PC, "IR(addr) -> PC"
            DriveIntAddrBus "IR", curOperand
            SetPC curOperand
            DoExecute = True

        Case OP_JZ, OP_JNZ, OP_JC, OP_JS
            taken = False
            Select Case curOp
                Case OP_JZ: taken = CPU.ZF
                Case OP_JNZ: taken = Not CPU.ZF
                Case OP_JC: taken = CPU.CF
                Case OP_JS: taken = CPU.SF
            End Select
            If taken Then
                Emit SIG_EN_IR_A Or SIG_LD_PC, "IR(addr) -> PC (taken)"
                DriveIntAddrBus "IR", curOperand
                SetPC curOperand
            Else
                Emit 0, MnemonicOf(curOp) & " not taken"
            End If
            DoExecute = True

        Case OP_IN
            r = IoRead(curReg)
            SetReg curReg, r
            DoExecute = True

        Case OP_OUT
            Emit SIG_EN_REG(curReg) Or SIG_IO_W Or SIG_LD_LCD, RegName(curReg) & " -> LCD"
            DriveDataBus RegName(curReg), GetReg(curReg), False
            IoWrite GetReg(curReg)
            DoExecute = True

        Case Else
            Emit 0, "??? opcode " & curOp
            DoExecute = True
    End Select
End Function

Private Function MnemonicOf(ByVal op As Integer) As String
    MnemonicOf = Mnemonic(op)
End Function

' =====================================================================
' Log sheet - one boxed row per micro-step.
' Title (row 1) is rewritten on every step: step, phase, micro-op, status.
' =====================================================================
Public Sub ClearLogSheet()
    Dim ws As Worksheet
    Dim titles As Variant
    Dim i As Integer
    Dim r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
    End If
    On Error GoTo 0
    On Error Resume Next
    ws.Range("A1:AB400").UnMerge
    On Error GoTo 0
    ws.Range("A1:AB400").Clear
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    On Error Resume Next
    ws.Tab.Color = CLR_ACTIVE
    On Error GoTo 0

    ws.Range("A1:AB1").Merge
    ws.Range("A1").Value = "STEP 0    IDLE    waiting    --"
    ws.Range("A1").Font.Name = "Segoe UI"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = CLR_TEXT
    ws.Range("A1").Interior.Color = CLR_BG
    ws.Range("A1").HorizontalAlignment = xlLeft
    ws.Range("A1").IndentLevel = 1
    ws.Rows(1).RowHeight = 32

    LogGroup ws, "A2:E2", "STEP"
    LogGroup ws, "F2:K2", "CPU"
    LogGroup ws, "L2:O2", "REGISTERS"
    LogGroup ws, "P2:R2", "FLAGS"
    LogGroup ws, "S2:X2", "BUSES"
    LogGroup ws, "Y2:AB2", "CACHE"
    ws.Rows(2).RowHeight = 18

    titles = Array("CLK", "Phase", "Micro", "Status", "What happened", _
        "PC", "MAR", "MDR", "IR", "TA", "TB", _
        "AX", "BX", "CX", "DX", _
        "ZF", "CF", "SF", _
        "Int addr", "Ext addr", "Data", "Who", "CW", "Signals", _
        "L1I", "L1D", "L2", "Faults")
    For i = 0 To LOG_COLS - 1
        ws.Cells(3, i + 1).Value = titles(i)
    Next i
    With ws.Range("A3:AB3")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = CLR_TEXT
        .Interior.Color = CLR_BG
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 20
    LogBox ws.Range("A1:AB3")

    ws.Columns("A").ColumnWidth = 7
    ws.Columns("B").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 7
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 26
    ws.Columns("F:R").ColumnWidth = 7
    ws.Columns("S:U").ColumnWidth = 11
    ws.Columns("V").ColumnWidth = 10
    ws.Columns("W").ColumnWidth = 12
    ws.Columns("X").ColumnWidth = 34
    ws.Columns("Y:AA").ColumnWidth = 8
    ws.Range("Y4:AA400").NumberFormat = "@"
    ws.Columns("AB").ColumnWidth = 9

    For r = LOG_FIRST To LOG_BLANK_TO
        PaintBlankLogRow ws, r
    Next r
    LogBox ws.Range("A4:AB" & LOG_BLANK_TO)

    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    lastFaults = 0
    lastL1IH = 0
    lastL1IM = 0
    lastL1DH = 0
    lastL1DM = 0
    lastL2H = 0
    lastL2M = 0
    logRow = LOG_FIRST
End Sub

Private Sub PaintBlankLogRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim row As Range
    Set row = ws.Range(ws.Cells(r, 1), ws.Cells(r, LOG_COLS))
    row.Font.Name = "Consolas"
    row.Font.Size = 10
    row.Font.Color = CLR_BG
    row.Font.Bold = False
    If r Mod 2 = 0 Then
        row.Interior.Color = CLR_TEXT
    Else
        row.Interior.Color = CLR_HIGHZ
    End If
    ws.Rows(r).RowHeight = 18
End Sub

Private Sub LogGroup(ByVal ws As Worksheet, ByVal addr As String, ByVal caption As String)
    ws.Range(addr).Merge
    ws.Range(addr).Value = caption
    ws.Range(addr).Font.Name = "Segoe UI"
    ws.Range(addr).Font.Size = 9
    ws.Range(addr).Font.Bold = True
    ws.Range(addr).Font.Color = CLR_BG
    ws.Range(addr).Interior.Color = CLR_TEXT
    ws.Range(addr).HorizontalAlignment = xlLeft
End Sub

Private Sub LogBox(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = CLR_BG
    End With
End Sub

Private Function PhaseName() As String
    If MicroStep >= 1 And MicroStep <= 6 Then
        PhaseName = "FETCH"
    ElseIf MicroStep = 7 Then
        PhaseName = "DECODE"
    ElseIf MicroStep > 7 Then
        PhaseName = "EXECUTE"
    Else
        PhaseName = "IDLE"
    End If
End Function

Private Function FlagText() As String
    FlagText = "Z" & IIf(CPU.ZF, "1", "0") & " C" & IIf(CPU.CF, "1", "0") & " S" & IIf(CPU.SF, "1", "0")
End Function

Private Function BusShown(ByVal v As Integer) As String
    If v = HIGH_Z Then
        BusShown = "HiZ"
    Else
        BusShown = Right$("000" & Hex$(v And &H1FF), 3) & "h"
    End If
End Function

Private Function StepStatus() As String
    If Stats.PageFaults > lastFaults Then
        StepStatus = "FAULT"
    ElseIf CPU.Halted Then
        StepStatus = "HALT"
    Else
        StepStatus = "OK"
    End If
End Function

Private Function WhoDrives() As String
    If Len(Bus.DataDriver) > 0 Then
        WhoDrives = Bus.DataDriver
    ElseIf Len(Bus.AddrDriver) > 0 Then
        WhoDrives = Bus.AddrDriver
    Else
        WhoDrives = "-"
    End If
End Function

Private Function PairCount(ByVal hits As Long, ByVal misses As Long) As String
    PairCount = CStr(hits) & "/" & CStr(misses)
End Function

Private Sub LogStep(ByVal cw As Long, ByVal description As String)
    Dim ws As Worksheet
    Dim row As Range
    Dim phase As String
    Dim status As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then ClearLogSheet: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If logRow < LOG_FIRST Then logRow = LOG_FIRST
    If logRow > LOG_BLANK_TO Then PaintBlankLogRow ws, logRow
    phase = PhaseName()
    status = StepStatus()
    ws.Range("A1").Value = "STEP " & Clock & "    " & phase & "    " & CurrentMicroText() & "    " & status

    Set row = ws.Range(ws.Cells(logRow, 1), ws.Cells(logRow, LOG_COLS))
    row.Font.Name = "Consolas"
    row.Font.Size = 10
    row.Font.Color = CLR_BG
    row.Font.Bold = False
    row.HorizontalAlignment = xlCenter
    If logRow Mod 2 = 0 Then
        row.Interior.Color = CLR_TEXT
    Else
        row.Interior.Color = CLR_HIGHZ
    End If
    ws.Cells(logRow, 5).HorizontalAlignment = xlLeft

    ws.Cells(logRow, 1).Value = Clock
    ws.Cells(logRow, 2).Value = phase
    ws.Cells(logRow, 3).Value = MicroTag()
    ws.Cells(logRow, 4).Value = status
    ws.Cells(logRow, 5).Value = description
    ws.Cells(logRow, 6).Value = CPU.PC
    ws.Cells(logRow, 7).Value = CPU.MAR
    ws.Cells(logRow, 8).Value = CPU.MDR
    ws.Cells(logRow, 9).Value = Right$("0000" & Hex$(CPU.IR And &HFFFF&), 4) & "h"
    ws.Cells(logRow, 10).Value = CPU.TA
    ws.Cells(logRow, 11).Value = CPU.TB
    ws.Cells(logRow, 12).Value = CPU.R(R_AX)
    ws.Cells(logRow, 13).Value = CPU.R(R_BX)
    ws.Cells(logRow, 14).Value = CPU.R(R_CX)
    ws.Cells(logRow, 15).Value = CPU.R(R_DX)
    ws.Cells(logRow, 16).Value = IIf(CPU.ZF, 1, 0)
    ws.Cells(logRow, 17).Value = IIf(CPU.CF, 1, 0)
    ws.Cells(logRow, 18).Value = IIf(CPU.SF, 1, 0)
    ws.Cells(logRow, 19).Value = BusShown(Bus.IntAddr)
    ws.Cells(logRow, 20).Value = BusShown(Bus.ExtAddr)
    ws.Cells(logRow, 21).Value = BusShown(Bus.Data)
    ws.Cells(logRow, 22).Value = WhoDrives()
    ws.Cells(logRow, 23).Value = "0x" & Hex$(cw)
    ws.Cells(logRow, 24).Value = CW_ToString(cw)
    ws.Cells(logRow, 25).NumberFormat = "@"
    ws.Cells(logRow, 26).NumberFormat = "@"
    ws.Cells(logRow, 27).NumberFormat = "@"
    ws.Cells(logRow, 25).Value = PairCount(Stats.L1IHit, Stats.L1IMiss)
    ws.Cells(logRow, 26).Value = PairCount(Stats.L1DHit, Stats.L1DMiss)
    ws.Cells(logRow, 27).Value = PairCount(Stats.L2Hit, Stats.L2Miss)
    ws.Cells(logRow, 28).Value = Stats.PageFaults

    PaintPhaseCell ws.Cells(logRow, 2), phase
    PaintStatusCell ws.Cells(logRow, 4), status
    PaintFlagCell ws.Cells(logRow, 16), CPU.ZF
    PaintFlagCell ws.Cells(logRow, 17), CPU.CF
    PaintFlagCell ws.Cells(logRow, 18), CPU.SF
    PaintCacheCell ws.Cells(logRow, 25), Stats.L1IHit, Stats.L1IMiss, lastL1IH, lastL1IM
    PaintCacheCell ws.Cells(logRow, 26), Stats.L1DHit, Stats.L1DMiss, lastL1DH, lastL1DM
    PaintCacheCell ws.Cells(logRow, 27), Stats.L2Hit, Stats.L2Miss, lastL2H, lastL2M
    If Stats.PageFaults > lastFaults Then
        ws.Cells(logRow, 28).Interior.Color = CLR_FAULT
        ws.Cells(logRow, 28).Font.Color = CLR_TEXT
        ws.Cells(logRow, 28).Font.Bold = True
    End If
    LogBox row
    ws.Rows(logRow).RowHeight = 18

    lastFaults = Stats.PageFaults
    lastL1IH = Stats.L1IHit
    lastL1IM = Stats.L1IMiss
    lastL1DH = Stats.L1DHit
    lastL1DM = Stats.L1DMiss
    lastL2H = Stats.L2Hit
    lastL2M = Stats.L2Miss
    logRow = logRow + 1
End Sub

Private Function MicroTag() As String
    If MicroStep >= 1 And MicroStep <= 6 Then
        MicroTag = "F" & MicroStep
    ElseIf MicroStep = 7 Then
        MicroTag = "DEC"
    ElseIf MicroStep > 7 Then
        MicroTag = "E" & (MicroStep - 7)
    Else
        MicroTag = "-"
    End If
End Function

Private Sub PaintPhaseCell(ByVal cell As Range, ByVal phase As String)
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
    If phase = "FETCH" Then
        cell.Interior.Color = CLR_ACTIVE
        cell.Font.Color = CLR_TEXT
    ElseIf phase = "DECODE" Then
        cell.Interior.Color = CLR_LOAD
        cell.Font.Color = CLR_BG
    ElseIf phase = "EXECUTE" Then
        cell.Interior.Color = CLR_HIT
        cell.Font.Color = CLR_TEXT
    End If
End Sub

Private Sub PaintStatusCell(ByVal cell As Range, ByVal status As String)
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
    If status = "OK" Then
        cell.Interior.Color = CLR_HIT
        cell.Font.Color = CLR_TEXT
    ElseIf status = "HALT" Then
        cell.Interior.Color = CLR_ACTIVE
        cell.Font.Color = CLR_TEXT
    ElseIf status = "FAULT" Then
        cell.Interior.Color = CLR_FAULT
        cell.Font.Color = CLR_TEXT
    End If
End Sub

Private Sub PaintCacheCell(ByVal cell As Range, ByVal hits As Long, ByVal misses As Long, _
                           ByVal prevHits As Long, ByVal prevMisses As Long)
    cell.HorizontalAlignment = xlCenter
    If misses > prevMisses Then
        cell.Interior.Color = CLR_MISS
        cell.Font.Color = CLR_BG
        cell.Font.Bold = True
    ElseIf hits > prevHits Then
        cell.Interior.Color = CLR_HIT
        cell.Font.Color = CLR_TEXT
        cell.Font.Bold = True
    End If
End Sub

Private Sub PaintFlagCell(ByVal cell As Range, ByVal onFlag As Boolean)
    cell.HorizontalAlignment = xlCenter
    If onFlag Then
        cell.Interior.Color = CLR_LOAD
        cell.Font.Bold = True
    End If
End Sub

' =====================================================================
' Test_ControlUnit - fetch + a few executes on a tiny program in VMem/RAM
' =====================================================================
Public Function Test_ControlUnit() As Long
    Dim fails As Long
    Dim i As Integer
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    AnimDelay = 0
    On Error GoTo Clean

    ResetMachine
    ResetControlUnit
    InitPageTable

    ' Map VPN 0 -> frame 0 so fetch at PC=0 works
    With PageTable(0)
        .Present = True
        .PFN = 0
        .Pinned = False
    End With
    FrameOwner(0) = 0

    ' Program at vaddr 0: MOV AX, 7 ; MOV BX, 3 ; ADD AX, BX ; HLT
    ' Each instr = 2 bytes in virtual memory / through MemRead
    PokeInstr 0, Encode(OP_MOV_RI, R_AX, 7)
    PokeInstr 2, Encode(OP_MOV_RR, R_BX, R_AX)   ' BX <- AX (7) - simpler path
    ' Better catalog-like: MOV BX,3 then ADD
    PokeInstr 2, Encode(OP_MOV_RI, R_BX, 3)
    PokeInstr 4, Encode(OP_ADD_RR, R_AX, R_BX)
    PokeInstr 6, Encode(OP_HLT, 0, 0)

    ' Also mirror into physical frame 0 via VMem/RAM for fetch
    For i = 0 To 7
        RAM(i) = VMem(i)
    Next i

    SetPC 0
    MicroStep = 1

    ' Run until HLT or safety cap
    Dim guard As Long
    guard = 0
    Do While Not CPU.Halted And guard < 200
        StepMicro
        guard = guard + 1
    Loop

    fails = fails + CuCheck("halted", CPU.Halted)
    fails = fails + CuCheck("AX = 10 (7+3)", GetReg(R_AX) = 10)
    fails = fails + CuCheck("BX = 3", GetReg(R_BX) = 3)
    fails = fails + CuCheck("log has rows", logRow > 2)
    fails = fails + CuCheck("finished under guard", guard < 200)

    ' Conditional jump: ZF path
    ResetMachine
    ResetControlUnit
    InitPageTable
    With PageTable(0)
        .Present = True
        .PFN = 0
    End With
    FrameOwner(0) = 0
    ' MOV AX, 0 ; CMP AX, 0 ; JZ 8 ; MOV AX, 99 ; HLT ; (at 8) MOV AX, 42 ; HLT
    PokeInstr 0, Encode(OP_MOV_RI, R_AX, 0)
    PokeInstr 2, Encode(OP_CMP_RI, R_AX, 0)
    PokeInstr 4, Encode(OP_JZ, 0, 8)
    PokeInstr 6, Encode(OP_MOV_RI, R_AX, 99)
    PokeInstr 8, Encode(OP_MOV_RI, R_AX, 42)
    PokeInstr 10, Encode(OP_HLT, 0, 0)
    For i = 0 To 11: RAM(i) = VMem(i): Next i
    SetPC 0
    MicroStep = 1
    guard = 0
    Do While Not CPU.Halted And guard < 300
        StepMicro
        guard = guard + 1
    Loop
    fails = fails + CuCheck("JZ taken -> AX=42", GetReg(R_AX) = 42)
    fails = fails + CuCheck("log title names the step", InStr(CStr(ThisWorkbook.Worksheets(SH_LOG).Range("A1").Value), "STEP ") = 1)
    fails = fails + CuCheck("log title shows HALT", InStr(CStr(ThisWorkbook.Worksheets(SH_LOG).Range("A1").Value), "HALT") > 0)
    fails = fails + CuCheck("log header is CLK", CStr(ThisWorkbook.Worksheets(SH_LOG).Range("A3").Value) = "CLK")
    fails = fails + CuCheck("log row is boxed", ThisWorkbook.Worksheets(SH_LOG).Range("A4").Borders(xlEdgeLeft).LineStyle = xlContinuous)

    GoTo Done
Clean:
    fails = fails + 1
    Debug.Print "Test_ControlUnit ERROR: " & Err.Number & " " & Err.Description
Done:
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Debug.Print IIf(fails = 0, "Test_ControlUnit: ALL OK", "Test_ControlUnit: " & fails & " FAIL(S)")
    Test_ControlUnit = fails
End Function

' Write a 16-bit instruction into VMem at virtual byte address.
Private Sub PokeInstr(ByVal vaddr As Integer, ByVal word As Long)
    VMem(vaddr) = CInt((word \ 256) And BYTE_MAX)
    VMem(vaddr + 1) = CInt(word And BYTE_MAX)
End Sub

Private Function CuCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    CuCheck = IIf(passed, 0, 1)
End Function
