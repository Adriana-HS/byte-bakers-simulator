Attribute VB_Name = "modISA"
' =====================================================================
' modISA - Instruction set of the Byte Bakers CPU
' Owner: Adriana  |  Task: #3 (A3)
'
' Instruction format (fixed 2 bytes):
'   Byte 0: [ opcode: 5 bits | reg: 2 bits | A8: 1 bit ]
'   Byte 1: [ operand: 8 bits ]
'   - Immediate:       operand = value 0-255
'   - Address:         A8 + operand = 9-bit virtual address 0-511
'   - Source register: 2 low bits of the operand (AX=0 BX=1 CX=2 DX=3)
' =====================================================================
Option Explicit

' ---------------------------------------------------------------------
' Opcodes (5 bits: 0-31)
' ---------------------------------------------------------------------
Public Const OP_NOP As Integer = &H0          ' NOP
Public Const OP_MOV_RI As Integer = &H1       ' MOV r, imm
Public Const OP_MOV_RR As Integer = &H2       ' MOV r, r2
Public Const OP_LOAD As Integer = &H3         ' LOAD r, [addr]
Public Const OP_STORE As Integer = &H4        ' STORE [addr], r
Public Const OP_ADD_RR As Integer = &H5       ' ADD r, r2
Public Const OP_ADD_RI As Integer = &H6       ' ADD r, imm
Public Const OP_SUB_RR As Integer = &H7       ' SUB r, r2
Public Const OP_SUB_RI As Integer = &H8       ' SUB r, imm
Public Const OP_INC As Integer = &H9          ' INC r
Public Const OP_DEC As Integer = &HA          ' DEC r
Public Const OP_CMP_RR As Integer = &HB       ' CMP r, r2
Public Const OP_CMP_RI As Integer = &HC       ' CMP r, imm
Public Const OP_AND As Integer = &HD          ' AND r, r2
Public Const OP_OR As Integer = &HE           ' OR r, r2
Public Const OP_XOR As Integer = &HF          ' XOR r, r2
Public Const OP_JMP As Integer = &H10         ' JMP addr
Public Const OP_JZ As Integer = &H11          ' JZ addr
Public Const OP_JNZ As Integer = &H12         ' JNZ addr
Public Const OP_JC As Integer = &H13          ' JC addr
Public Const OP_JS As Integer = &H14          ' JS addr
Public Const OP_IN As Integer = &H15          ' IN r, [KBD]
Public Const OP_OUT As Integer = &H16         ' OUT [LCD], r
Public Const OP_HLT As Integer = &H1F         ' HLT

Public Const OPCODE_MAX As Integer = 31       ' 5-bit opcode field
Public Const NUM_INSTRUCTIONS As Integer = 24 ' opcodes in use

' ISA reference sheet layout (see BuildISASheet)
Public Const ISA_HEADER_ROW As Long = 11      ' table header; instructions start below it
Public Const ISA_TABLE_LAST_COL As Long = 14  ' column N; the CPU panel highlights through here
Private Const ISA_COLS As Long = 13           ' columns B..N

' =====================================================================
' Encode - packs one instruction into 16 bits.
'   op      opcode 0-31
'   reg     destination register 0-3 (or source register for STORE/OUT)
'   operand immediate 0-255, 9-bit address 0-511 or source register 0-3
' Example: Encode(OP_ADD_RR, R_AX, R_BX) = &H2801  (ADD AX, BX)
' =====================================================================
Public Function Encode(ByVal op As Integer, ByVal reg As Integer, ByVal operand As Integer) As Long
    Dim b0 As Long
    op = op And OPCODE_MAX
    reg = reg And 3
    operand = operand And VADDR_MAX               ' 9 bits: A8 + 8-bit operand
    b0 = CLng(op) * 8 + reg * 2 + (operand \ 256) ' opcode | reg | A8
    Encode = b0 * 256 + (operand And BYTE_MAX)
End Function

' =====================================================================
' Decode - splits the 16-bit IR into its fields (inverse of Encode).
'   operand comes back with A8 included (0-511); for immediates and
'   registers A8 is 0, so the value is the same 0-255.
' =====================================================================
Public Sub Decode(ByVal ir As Long, ByRef op As Integer, ByRef reg As Integer, ByRef operand As Integer)
    Dim b0 As Long
    b0 = (ir \ 256) And BYTE_MAX
    op = b0 \ 8
    reg = (b0 \ 2) And 3
    operand = (b0 And 1) * 256 + (ir And BYTE_MAX)
End Sub

' =====================================================================
' ISA table - the single source of truth for every instruction.
' Mnemonic, OpcodeInfo, Disassemble and the ISA sheet all read it,
' so the assembler, the CPU and the documentation can never disagree.
'
' Operand formats:
'   ""  no operands        "R"  register           "RR" reg, reg
'   "RI" reg, immediate    "RM" reg, [address]     "MR" [address], reg
'   "A"  address           "RK" reg, [KBD]         "LR" [LCD], reg
' Flags: Z = ZF, C = CF, S = SF ("-" = flags not changed)
' Returns False if the opcode is not used.
' =====================================================================
Private Function OpInfo(ByVal op As Integer, ByRef mn As String, ByRef fmt As String, _
                        ByRef syntax As String, ByRef flags As String, ByRef desc As String) As Boolean
    OpInfo = True
    Select Case op
        Case OP_NOP:    mn = "NOP":   fmt = "":   syntax = "NOP":             flags = "-":     desc = "Does nothing"
        Case OP_MOV_RI: mn = "MOV":   fmt = "RI": syntax = "MOV r, imm":      flags = "-":     desc = "r <- imm"
        Case OP_MOV_RR: mn = "MOV":   fmt = "RR": syntax = "MOV r, r2":       flags = "-":     desc = "r <- r2"
        Case OP_LOAD:   mn = "LOAD":  fmt = "RM": syntax = "LOAD r, [addr]":  flags = "-":     desc = "r <- M[addr]"
        Case OP_STORE:  mn = "STORE": fmt = "MR": syntax = "STORE [addr], r": flags = "-":     desc = "M[addr] <- r"
        Case OP_ADD_RR: mn = "ADD":   fmt = "RR": syntax = "ADD r, r2":       flags = "Z C S": desc = "r <- r + r2"
        Case OP_ADD_RI: mn = "ADD":   fmt = "RI": syntax = "ADD r, imm":      flags = "Z C S": desc = "r <- r + imm"
        Case OP_SUB_RR: mn = "SUB":   fmt = "RR": syntax = "SUB r, r2":       flags = "Z C S": desc = "r <- r - r2"
        Case OP_SUB_RI: mn = "SUB":   fmt = "RI": syntax = "SUB r, imm":      flags = "Z C S": desc = "r <- r - imm"
        Case OP_INC:    mn = "INC":   fmt = "R":  syntax = "INC r":           flags = "Z S":   desc = "r <- r + 1"
        Case OP_DEC:    mn = "DEC":   fmt = "R":  syntax = "DEC r":           flags = "Z S":   desc = "r <- r - 1"
        Case OP_CMP_RR: mn = "CMP":   fmt = "RR": syntax = "CMP r, r2":       flags = "Z C S": desc = "r - r2 (flags only)"
        Case OP_CMP_RI: mn = "CMP":   fmt = "RI": syntax = "CMP r, imm":      flags = "Z C S": desc = "r - imm (flags only)"
        Case OP_AND:    mn = "AND":   fmt = "RR": syntax = "AND r, r2":       flags = "Z S":   desc = "r <- r AND r2"
        Case OP_OR:     mn = "OR":    fmt = "RR": syntax = "OR r, r2":        flags = "Z S":   desc = "r <- r OR r2"
        Case OP_XOR:    mn = "XOR":   fmt = "RR": syntax = "XOR r, r2":       flags = "Z S":   desc = "r <- r XOR r2"
        Case OP_JMP:    mn = "JMP":   fmt = "A":  syntax = "JMP addr":        flags = "-":     desc = "PC <- addr"
        Case OP_JZ:     mn = "JZ":    fmt = "A":  syntax = "JZ addr":         flags = "-":     desc = "If ZF = 1: PC <- addr"
        Case OP_JNZ:    mn = "JNZ":   fmt = "A":  syntax = "JNZ addr":        flags = "-":     desc = "If ZF = 0: PC <- addr"
        Case OP_JC:     mn = "JC":    fmt = "A":  syntax = "JC addr":         flags = "-":     desc = "If CF = 1: PC <- addr"
        Case OP_JS:     mn = "JS":    fmt = "A":  syntax = "JS addr":         flags = "-":     desc = "If SF = 1: PC <- addr"
        Case OP_IN:     mn = "IN":    fmt = "RK": syntax = "IN r, [KBD]":     flags = "-":     desc = "r <- keyboard queue head (0 if empty)"
        Case OP_OUT:    mn = "OUT":   fmt = "LR": syntax = "OUT [LCD], r":    flags = "-":     desc = "LCD[cursor] <- r; cursor advances"
        Case OP_HLT:    mn = "HLT":   fmt = "":   syntax = "HLT":             flags = "-":     desc = "Halts the CPU"
        Case Else:      mn = "":      fmt = "":   syntax = "":                flags = "":      desc = "": OpInfo = False
    End Select
End Function

' True if the opcode belongs to the ISA.
Public Function IsValidOpcode(ByVal op As Integer) As Boolean
    Dim mn As String, fmt As String, syn As String, fl As String, ds As String
    IsValidOpcode = OpInfo(op, mn, fmt, syn, fl, ds)
End Function

' Text shown on the panel: "ADD", "JNZ"... ("???" if the opcode is not used).
Public Function Mnemonic(ByVal op As Integer) As String
    Dim mn As String, fmt As String, syn As String, fl As String, ds As String
    If OpInfo(op, mn, fmt, syn, fl, ds) Then Mnemonic = mn Else Mnemonic = "???"
End Function

' Operand format code of an opcode ("RR", "RI", "A"...).
Public Function OpFormat(ByVal op As Integer) As String
    Dim mn As String, fmt As String, syn As String, fl As String, ds As String
    OpInfo op, mn, fmt, syn, fl, ds
    OpFormat = fmt
End Function

' Syntax with generic operands: "ADD r, r2".
Public Function OpSyntax(ByVal op As Integer) As String
    Dim mn As String, fmt As String, syn As String, fl As String, ds As String
    OpInfo op, mn, fmt, syn, fl, ds
    OpSyntax = syn
End Function

' Flags the instruction updates: "Z C S", "Z S" or "-".
Public Function OpFlags(ByVal op As Integer) As String
    Dim mn As String, fmt As String, syn As String, fl As String, ds As String
    OpInfo op, mn, fmt, syn, fl, ds
    OpFlags = fl
End Function

' Register transfer description: "r <- r + r2".
Public Function OpDescription(ByVal op As Integer) As String
    Dim mn As String, fmt As String, syn As String, fl As String, ds As String
    OpInfo op, mn, fmt, syn, fl, ds
    OpDescription = ds
End Function

' Micro-steps, signals, cycle count and family. The sheet and the
' defense read this, so the columns cannot drift from each other.
' Fetch is the six steps modControlUnit always runs before execute.
Public Sub ISA_Detail(ByVal op As Integer, ByRef family As String, ByRef steps As String, _
                      ByRef sigs As String, ByRef cycles As Integer)
    Dim ex As String, es As String, n As Integer
    family = ""
    steps = ""
    sigs = ""
    cycles = 0
    If Not IsValidOpcode(op) Then Exit Sub
    ExecDetail op, family, ex, es, n
    ' Fetch F1-F6 is the same for every instruction, so the cell shows
    ' only the execute steps. Cycles still counts the six fetch steps.
    steps = ex
    sigs = es
    cycles = 6 + n
End Sub

Private Sub ExecDetail(ByVal op As Integer, ByRef family As String, ByRef steps As String, _
                       ByRef sigs As String, ByRef n As Integer)
    Select Case op
        Case OP_NOP
            family = "Control": n = 1
            steps = "E1 NOP": sigs = "E1 (none)"
        Case OP_HLT
            family = "Control": n = 1
            steps = "E1 HLT": sigs = "E1 (none)"
        Case OP_MOV_RI
            family = "Transfer": n = 1
            steps = "E1 IR(imm)->r": sigs = "E1 En_IR_D Ld_r"
        Case OP_MOV_RR
            family = "Transfer": n = 1
            steps = "E1 r2->r": sigs = "E1 En_r2 Ld_r"
        Case OP_LOAD
            family = "Transfer": n = 3
            steps = "E1 IR(addr)->MAR | E2 M[MAR]->MDR | E3 MDR->r"
            sigs = "E1 En_IR_A Ld_MAR | E2 Mem_R Ld_MDR | E3 En_MDR Ld_r"
        Case OP_STORE
            family = "Transfer": n = 3
            steps = "E1 IR(addr)->MAR | E2 r->MDR | E3 MDR->M[MAR]"
            sigs = "E1 En_IR_A Ld_MAR | E2 En_r Ld_MDR | E3 En_MDR Mem_W"
        Case OP_ADD_RR, OP_SUB_RR, OP_AND, OP_OR, OP_XOR
            family = IIf(op = OP_AND Or op = OP_OR Or op = OP_XOR, "Logic", "Arithmetic")
            n = 3
            steps = "E1 r->TA | E2 r2->TB | E3 ALU->r + flags"
            sigs = "E1 En_r Ld_TA | E2 En_r2 Ld_TB | E3 En_ALU Ld_r Ld_FLAGS"
        Case OP_ADD_RI, OP_SUB_RI
            family = "Arithmetic": n = 3
            steps = "E1 r->TA | E2 IR(imm)->TB | E3 ALU->r + flags"
            sigs = "E1 En_r Ld_TA | E2 En_IR_D Ld_TB | E3 En_ALU Ld_r Ld_FLAGS"
        Case OP_INC, OP_DEC
            family = "Arithmetic": n = 2
            steps = "E1 r->TA | E2 ALU->r + flags"
            sigs = "E1 En_r Ld_TA | E2 En_ALU Ld_r Ld_FLAGS"
        Case OP_CMP_RR
            family = "Arithmetic": n = 3
            steps = "E1 r->TA | E2 r2->TB | E3 ALU->FLAGS"
            sigs = "E1 En_r Ld_TA | E2 En_r2 Ld_TB | E3 En_ALU Ld_FLAGS"
        Case OP_CMP_RI
            family = "Arithmetic": n = 3
            steps = "E1 r->TA | E2 IR(imm)->TB | E3 ALU->FLAGS"
            sigs = "E1 En_r Ld_TA | E2 En_IR_D Ld_TB | E3 En_ALU Ld_FLAGS"
        Case OP_JMP
            family = "Jump": n = 1
            steps = "E1 IR(addr)->PC": sigs = "E1 En_IR_A Ld_PC"
        Case OP_JZ, OP_JNZ, OP_JC, OP_JS
            family = "Jump": n = 1
            steps = "E1 IR(addr)->PC if taken, else no transfer"
            sigs = "E1 En_IR_A Ld_PC (taken) or none"
        Case OP_IN
            family = "I/O": n = 1
            steps = "E1 KBD->r": sigs = "E1 IO_R En_KBD Ld_r"
        Case OP_OUT
            family = "I/O": n = 1
            steps = "E1 r->LCD": sigs = "E1 En_r IO_W Ld_LCD"
        Case Else
            family = "": n = 0: steps = "": sigs = ""
    End Select
End Sub

Public Function ISA_FamilyColor(ByVal family As String) As Long
    Select Case family
        Case "Transfer": ISA_FamilyColor = CLR_HIGHZ
        Case "Arithmetic": ISA_FamilyColor = CLR_LOAD
        Case "Logic": ISA_FamilyColor = CLR_HIT
        Case "Jump": ISA_FamilyColor = CLR_SOFT
        Case "I/O": ISA_FamilyColor = CLR_LCD
        Case "Control": ISA_FamilyColor = CLR_BG
        Case Else: ISA_FamilyColor = CLR_HIGHZ
    End Select
End Function

' =====================================================================
' OpcodeInfo - used by the assembler (#27): mnemonic + operand format
' -> opcode. Example: OpcodeInfo("ADD", "RR") = &H05.
' Returns -1 if that combination does not exist.
' =====================================================================
Public Function OpcodeInfo(ByVal mn As String, ByVal fmt As String) As Integer
    Dim op As Integer
    mn = UCase$(Trim$(mn))
    fmt = UCase$(Trim$(fmt))
    For op = 0 To OPCODE_MAX
        If IsValidOpcode(op) Then
            If Mnemonic(op) = mn And OpFormat(op) = fmt Then
                OpcodeInfo = op
                Exit Function
            End If
        End If
    Next op
    OpcodeInfo = -1
End Function

' =====================================================================
' Disassemble - IR (16 bits) -> readable text, e.g. "ADD AX, BX",
' "LOAD CX, [0x1FF]", "OUT [LCD], AX". Used by the panel and the Log.
' =====================================================================
Public Function Disassemble(ByVal ir As Long) As String
    Dim op As Integer, reg As Integer, operand As Integer
    Dim mn As String, r As String
    Decode ir, op, reg, operand
    If Not IsValidOpcode(op) Then
        Disassemble = "??? (0x" & Right$("0000" & Hex$(ir), 4) & ")"
        Exit Function
    End If
    mn = Mnemonic(op)
    r = RegName(reg)
    Select Case OpFormat(op)
        Case "":   Disassemble = mn
        Case "R":  Disassemble = mn & " " & r
        Case "RR": Disassemble = mn & " " & r & ", " & RegName(operand And 3)
        Case "RI": Disassemble = mn & " " & r & ", " & operand
        Case "RM": Disassemble = mn & " " & r & ", [" & HexAddr(operand) & "]"
        Case "MR": Disassemble = mn & " [" & HexAddr(operand) & "], " & r
        Case "A":  Disassemble = mn & " " & HexAddr(operand)
        Case "RK": Disassemble = mn & " " & r & ", [KBD]"
        Case "LR": Disassemble = mn & " [LCD], " & r
    End Select
End Function

' 9-bit address as "0x1FF".
Public Function HexAddr(ByVal addr As Integer) As String
    HexAddr = "0x" & Right$("000" & Hex$(addr And VADDR_MAX), 3)
End Function

' =====================================================================
' Test_ISA - self-check of the instruction set.
' Run it from the VBA editor (F5); results go to the Immediate window.
' =====================================================================
Public Sub Test_ISA()
    Dim fails As Integer, used As Integer, roundTripFails As Integer
    Dim op As Integer, reg As Integer, k As Integer
    Dim dOp As Integer, dReg As Integer, dOperand As Integer
    Dim operands As Variant

    ' Edge cases from the task
    fails = fails + IsaCheck("ADD AX, BX = 0x2801", Encode(OP_ADD_RR, R_AX, R_BX) = &H2801&)
    fails = fails + IsaCheck("LOAD DX, [0x1FF] = 0x1FFF", Encode(OP_LOAD, R_DX, &H1FF) = &H1FFF&)
    fails = fails + IsaCheck("MOV DX, 255 = 0x0EFF", Encode(OP_MOV_RI, R_DX, 255) = &HEFF&)
    fails = fails + IsaCheck("HLT = 0xF800", Encode(OP_HLT, 0, 0) = &HF800&)
    Decode &H1FFF&, dOp, dReg, dOperand
    fails = fails + IsaCheck("Decode 0x1FFF -> LOAD, DX, 0x1FF", dOp = OP_LOAD And dReg = R_DX And dOperand = &H1FF)

    ' Encode/Decode round trip for every opcode, register and operand
    operands = Array(0, 1, 2, 3, 127, 128, 255, 256, 300, 511)
    For op = 0 To OPCODE_MAX
        If IsValidOpcode(op) Then used = used + 1
        For reg = 0 To 3
            For k = LBound(operands) To UBound(operands)
                Decode Encode(op, reg, operands(k)), dOp, dReg, dOperand
                If dOp <> op Or dReg <> reg Or dOperand <> operands(k) Then roundTripFails = roundTripFails + 1
            Next k
        Next reg
    Next op
    fails = fails + IsaCheck("Encode/Decode round trip (1280 cases)", roundTripFails = 0)
    fails = fails + IsaCheck(NUM_INSTRUCTIONS & " instructions in the ISA", used = NUM_INSTRUCTIONS)

    ' Assembler lookup
    fails = fails + IsaCheck("OpcodeInfo(""ADD"", ""RR"") = 0x05", OpcodeInfo("ADD", "RR") = &H5)
    fails = fails + IsaCheck("OpcodeInfo is case-insensitive", OpcodeInfo("add", "ri") = OP_ADD_RI)
    fails = fails + IsaCheck("OpcodeInfo of an invalid pair = -1", OpcodeInfo("ADD", "A") = -1 And OpcodeInfo("FOO", "") = -1)
    roundTripFails = 0
    For op = 0 To OPCODE_MAX
        If IsValidOpcode(op) Then
            If OpcodeInfo(Mnemonic(op), OpFormat(op)) <> op Then roundTripFails = roundTripFails + 1
        End If
    Next op
    fails = fails + IsaCheck("Every mnemonic + format is unique", roundTripFails = 0)

    ' Disassembler
    fails = fails + IsaCheck("Disassemble ADD AX, BX", Disassemble(&H2801&) = "ADD AX, BX")
    fails = fails + IsaCheck("Disassemble LOAD CX, [0x1FF]", Disassemble(Encode(OP_LOAD, R_CX, &H1FF)) = "LOAD CX, [0x1FF]")
    fails = fails + IsaCheck("Disassemble STORE [0x040], BX", Disassemble(Encode(OP_STORE, R_BX, &H40)) = "STORE [0x040], BX")
    fails = fails + IsaCheck("Disassemble MOV DX, 255", Disassemble(Encode(OP_MOV_RI, R_DX, 255)) = "MOV DX, 255")
    fails = fails + IsaCheck("Disassemble OUT [LCD], AX", Disassemble(Encode(OP_OUT, R_AX, 0)) = "OUT [LCD], AX")
    fails = fails + IsaCheck("Disassemble JNZ 0x010", Disassemble(Encode(OP_JNZ, 0, &H10)) = "JNZ 0x010")

    Dim fam As String, st As String, sg As String, cy As Integer, nOps As Integer
    ISA_Detail OP_LOAD, fam, st, sg, cy
    fails = fails + IsaCheck("LOAD family and cycles", fam = "Transfer" And cy = 9 And InStr(st, "IR(addr)->MAR") > 0)
    fails = fails + IsaCheck("LOAD signals", InStr(sg, "En_IR_A") > 0 And InStr(sg, "Mem_R") > 0 And InStr(sg, "Ld_MDR") > 0)
    ISA_Detail OP_IN, fam, st, sg, cy
    fails = fails + IsaCheck("IN is I/O", fam = "I/O" And cy = 7 And InStr(sg, "E1 IO_R En_KBD Ld_r") > 0)
    ISA_Detail OP_ADD_RR, fam, st, sg, cy
    fails = fails + IsaCheck("ADD uses TA TB and flags", fam = "Arithmetic" And cy = 9 And InStr(sg, "Ld_FLAGS") > 0)
    nOps = 0
    For op = 0 To OPCODE_MAX
        If IsValidOpcode(op) Then
            nOps = nOps + 1
            ISA_Detail op, fam, st, sg, cy
            If Len(st) = 0 Or Len(sg) = 0 Or cy < 7 Or Len(fam) = 0 Then fails = fails + 1
        End If
    Next op
    fails = fails + IsaCheck("Every instruction has steps, signals, family and cycles", nOps = NUM_INSTRUCTIONS And fails >= 0)

    Debug.Print IIf(fails = 0, "Test_ISA: ALL OK", "Test_ISA: " & fails & " FAIL(S)")
End Sub

Private Function IsaCheck(ByVal what As String, ByVal passed As Boolean) As Integer
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    IsaCheck = IIf(passed, 0, 1)
End Function

' =====================================================================
' ISA reference sheet
' Built from the ISA table above, so it always matches what the CPU
' executes. Cells only: no shapes, so the animation engine never
' looks here.
' =====================================================================

Public Sub BuildISASheet()
    Dim ws As Worksheet, op As Integer, r As Long, ir As Long, lastRow As Long
    Dim fam As String, st As String, sg As String, cy As Integer

    Set ws = IsaSheet()
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 10
    ws.Columns("B:N").NumberFormat = "@"      ' keep "00101" and "0x05" as text

    ' Title
    With ws.Range("B2")
        .Value = "Byte Bakers - Instruction Set Architecture"
        .Font.Size = 14
        .Font.Bold = True
    End With
    ws.Range("B3").Value = NUM_INSTRUCTIONS & " instructions, fixed 2-byte format. Registers: AX=0, BX=1, CX=2, DX=3."
    ws.Range("B4").Value = "Cycles include F1-F6 fetch (PC->MAR, M->MDR, MDR->IR, twice). The columns below are only the execute steps."

    ' Instruction format
    ws.Range("B5").Value = "Instruction format"
    ws.Range("B5").Font.Bold = True
    IsaHeader ws.Range("B6:E6"), Array("Byte", "Bits 7-3", "Bits 2-1", "Bit 0")
    ws.Range("B7:E7").Value = Array("Byte 0", "opcode (5 bits)", "reg (2 bits)", "A8 (1 bit)")
    ws.Range("B8").Value = "Byte 1"
    ws.Range("C8").Value = "operand (8 bits): immediate 0-255 | address low byte (A8 + operand = 9-bit address 0-511) | source register in bits 1-0"
    IsaBorders ws.Range("B6:E8")

    ' Instruction table
    IsaHeader ws.Range("B" & ISA_HEADER_ROW).Resize(1, ISA_COLS), _
        Array("Opcode", "Mnemonic", "Family", "Cycles", "Syntax", "Micro-steps", "Signals", _
              "Format", "Flags", "Description", "Example", "Machine code", "Binary")
    r = ISA_HEADER_ROW
    For op = 0 To OPCODE_MAX
        If IsValidOpcode(op) Then
            r = r + 1
            ir = Encode(op, R_AX, SampleOperand(OpFormat(op)))
            ISA_Detail op, fam, st, sg, cy
            ws.Cells(r, 2).Resize(1, ISA_COLS).Value = Array( _
                "0x" & Right$("0" & Hex$(op), 2), _
                Mnemonic(op), _
                fam, _
                cy, _
                OpSyntax(op), _
                st, _
                sg, _
                IIf(OpFormat(op) = "", "none", OpFormat(op)), _
                OpFlags(op), _
                OpDescription(op), _
                Disassemble(ir), _
                "0x" & Right$("000" & Hex$(ir), 4), _
                Right$("0000" & IsaBin(op), 5))
            ws.Cells(r, 2).Resize(1, ISA_COLS).Interior.Color = CLR_TEXT
            ws.Cells(r, 2).Resize(1, ISA_COLS).Font.Color = CLR_BG
            ws.Cells(r, 2).Resize(1, ISA_COLS).Font.Name = "Consolas"
            ws.Cells(r, 4).Interior.Color = ISA_FamilyColor(fam)
            If fam = "Control" Then ws.Cells(r, 4).Font.Color = CLR_TEXT
            ws.Rows(r).RowHeight = 18
        End If
    Next op
    lastRow = r
    IsaBorders ws.Range("B" & ISA_HEADER_ROW).Resize(lastRow - ISA_HEADER_ROW + 1, ISA_COLS)

    ' Format legend
    r = r + 2
    ws.Cells(r, 2).Value = "Operand formats"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r + 1, 2).Value = "R = register   RR = reg, reg   RI = reg, immediate   RM = reg, [address]   MR = [address], reg"
    ws.Cells(r + 2, 2).Value = "A = address   RK = reg, [KBD]   LR = [LCD], reg   Flags: Z = ZF, C = CF, S = SF, - = unchanged"
    ws.Cells(r + 4, 2).Value = "Signal colors (same as the simulator)"
    ws.Cells(r + 4, 2).Font.Bold = True
    ws.Cells(r + 5, 2).Value = "En active"
    ws.Cells(r + 5, 2).Interior.Color = CLR_ACTIVE
    ws.Cells(r + 5, 2).Font.Color = CLR_TEXT
    ws.Cells(r + 5, 3).Value = "Ld load"
    ws.Cells(r + 5, 3).Interior.Color = CLR_LOAD
    ws.Cells(r + 5, 4).Value = "High-Z"
    ws.Cells(r + 5, 4).Interior.Color = CLR_HIGHZ

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 9
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 36
    ws.Columns("H").ColumnWidth = 36
    ws.Columns("I").ColumnWidth = 8
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").ColumnWidth = 28
    ws.Columns("L").ColumnWidth = 16
    ws.Columns("M").ColumnWidth = 12
    ws.Columns("N").ColumnWidth = 8
    Application.ScreenUpdating = True

    ' A7: append 27-signal control word table + rules (no-op stub until modControlWord lands)
    On Error Resume Next
    PaintControlWordOnISA
    On Error GoTo 0
End Sub

' Row of an opcode in the ISA table (0 if the opcode is not used).
' The CPU panel (#9) uses it to highlight the instruction being executed.
Public Function ISARow(ByVal op As Integer) As Long
    Dim i As Integer, n As Long
    If op < 0 Or op > OPCODE_MAX Then Exit Function
    If Not IsValidOpcode(op) Then Exit Function
    For i = 0 To op
        If IsValidOpcode(i) Then n = n + 1
    Next i
    ISARow = ISA_HEADER_ROW + n
End Function

' Returns the ISA sheet, creating it at the end if the workbook has none.
Private Function IsaSheet() As Worksheet
    On Error Resume Next
    Set IsaSheet = ThisWorkbook.Worksheets(SH_ISA)
    On Error GoTo 0
    If IsaSheet Is Nothing Then
        Set IsaSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        IsaSheet.Name = SH_ISA
    End If
End Function

' Operand used for the example column of each format.
Private Function SampleOperand(ByVal fmt As String) As Integer
    Select Case fmt
        Case "RR": SampleOperand = R_BX
        Case "RI": SampleOperand = 10
        Case "RM", "MR", "A": SampleOperand = &H40
        Case Else: SampleOperand = 0
    End Select
End Function

Private Function IsaBin(ByVal v As Integer) As String
    Do
        IsaBin = CStr(v Mod 2) & IsaBin
        v = v \ 2
    Loop While v > 0
End Function

Private Sub IsaHeader(ByVal rng As Range, ByVal titles As Variant)
    rng.Value = titles
    rng.Font.Bold = True
    rng.Font.Color = CLR_TEXT
    rng.Interior.Color = CLR_BG
    rng.RowHeight = 22
End Sub

Private Sub IsaBorders(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = CLR_BG
    End With
End Sub
