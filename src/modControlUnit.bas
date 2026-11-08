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

Private curOp As Integer
Private curReg As Integer
Private curOperand As Integer
Private logRow As Long
Private lastDesc As String

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
    ClearLogSheet
End Sub

' =====================================================================
' Emit - validate control word, publish on Bus, log + soft CW paint
' =====================================================================
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
            Emit 0, "HLT"
            CPU.Halted = True
            RunFlag = False
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
            Emit SIG_IO_R Or SIG_EN_KBD Or SIG_LD_REG(curReg), "KBD -> " & RegName(curReg)
            DriveDataBus "KBD", r, False
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
' Log sheet
' =====================================================================
Private Sub ClearLogSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
    End If
    ws.Cells.Clear
    ws.Range("A1:K1").Value = Array("CLK", "Phase", "PC", "AX", "BX", "CX", "DX", "Flags", "CTRL", "CW hex", "Transfer")
    With ws.Range("A1:K1")
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Color = CLR_TEXT
        .Interior.Color = CLR_BG
    End With
    ws.Range("A2:K200").Font.Name = "Consolas"
    ws.Range("A2:K200").Font.Size = 10
    ws.Columns("A:K").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 28
    ws.Columns("K").ColumnWidth = 22
    ws.Rows(1).RowHeight = 20
    logRow = 2
    On Error GoTo 0
End Sub

Private Function PhaseName() As String
    If MicroStep >= 1 And MicroStep <= 6 Then
        PhaseName = "FETCH"
    ElseIf MicroStep = 7 Then
        PhaseName = "DECODE"
    Else
        PhaseName = "EXECUTE"
    End If
End Function

Private Function FlagText() As String
    FlagText = "Z" & IIf(CPU.ZF, "1", "0") & " C" & IIf(CPU.CF, "1", "0") & " S" & IIf(CPU.SF, "1", "0")
End Function

Private Sub LogStep(ByVal cw As Long, ByVal description As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then ClearLogSheet: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If logRow < 2 Then logRow = 2
    ws.Cells(logRow, 1).Value = Clock
    ws.Cells(logRow, 2).Value = PhaseName()
    ws.Cells(logRow, 3).Value = CPU.PC
    ws.Cells(logRow, 4).Value = CPU.R(R_AX)
    ws.Cells(logRow, 5).Value = CPU.R(R_BX)
    ws.Cells(logRow, 6).Value = CPU.R(R_CX)
    ws.Cells(logRow, 7).Value = CPU.R(R_DX)
    ws.Cells(logRow, 8).Value = FlagText()
    ws.Cells(logRow, 9).Value = CW_ToString(cw)
    ws.Cells(logRow, 10).Value = "0x" & Hex$(cw)
    ws.Cells(logRow, 11).Value = description
    If logRow Mod 2 = 0 Then ws.Range(ws.Cells(logRow, 1), ws.Cells(logRow, 11)).Interior.Color = CLR_TEXT
    logRow = logRow + 1
    On Error GoTo 0
End Sub

' =====================================================================
' Test_ControlUnit - fetch + a few executes on a tiny program in VMem/RAM
' =====================================================================
Public Function Test_ControlUnit() As Long
    Dim fails As Long
    Dim i As Integer

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
