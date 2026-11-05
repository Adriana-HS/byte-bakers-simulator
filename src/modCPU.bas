Attribute VB_Name = "modCPU"
' =====================================================================
' modCPU - Register file of the Byte Bakers CPU
' Owner: Adriana  |  Task: #4 (A4)
'
'   General registers AX, BX, CX, DX ....... 8 bits (0-255)
'   PC, MAR (virtual addresses) ............ 9 bits (0-511)
'   MDR .................................... 8 bits (0-255)
'   IR (whole instruction) ................. 16 bits
' Every write wraps around like real hardware: 256 -> 0, -1 -> 255.
' =====================================================================
Option Explicit

' ---------------------------------------------------------------------
' Reset: every register, temporary and flag to 0, PC = 0, not halted.
' ---------------------------------------------------------------------
Public Sub ResetCPU()
    Dim emptyCPU As TCPU
    CPU = emptyCPU
End Sub

' ---------------------------------------------------------------------
' Wraparound (modular arithmetic, works for negative values too)
' ---------------------------------------------------------------------
Public Function Wrap8(ByVal v As Long) As Integer
    Wrap8 = ((v Mod 256) + 256) Mod 256
End Function

Public Function Wrap9(ByVal v As Long) As Integer
    Wrap9 = ((v Mod 512) + 512) Mod 512
End Function

' ---------------------------------------------------------------------
' General purpose registers: index 0-3 = AX, BX, CX, DX
' ---------------------------------------------------------------------
Public Function GetReg(ByVal i As Integer) As Integer
    CheckRegIndex i
    GetReg = CPU.R(i)
End Function

Public Sub SetReg(ByVal i As Integer, ByVal v As Long)
    CheckRegIndex i
    CPU.R(i) = Wrap8(v)
End Sub

Private Sub CheckRegIndex(ByVal i As Integer)
    If i < 0 Or i > NUM_REGS - 1 Then
        Err.Raise vbObjectError + 600, "modCPU", "Invalid register index " & i & " (valid: 0=AX, 1=BX, 2=CX, 3=DX)"
    End If
End Sub

' ---------------------------------------------------------------------
' Control registers
' ---------------------------------------------------------------------
Public Sub SetPC(ByVal v As Long)
    CPU.PC = Wrap9(v)
End Sub

Public Sub IncPC()                      ' Inc_PC signal: PC <- PC + 1 (511 -> 0)
    CPU.PC = Wrap9(CLng(CPU.PC) + 1)
End Sub

Public Sub SetMAR(ByVal v As Long)
    CPU.MAR = Wrap9(v)
End Sub

Public Sub SetMDR(ByVal v As Long)
    CPU.MDR = Wrap8(v)
End Sub

Public Sub SetIR(ByVal v As Long)
    CPU.IR = v And &HFFFF&
End Sub

' ---------------------------------------------------------------------
' Register names (used by the assembler, the disassembler and the panel)
' ---------------------------------------------------------------------
Public Function RegName(ByVal i As Integer) As String
    Select Case i
        Case R_AX: RegName = "AX"
        Case R_BX: RegName = "BX"
        Case R_CX: RegName = "CX"
        Case R_DX: RegName = "DX"
        Case Else: RegName = "??"
    End Select
End Function

' "CX" -> 2. Case and spaces are ignored. Returns -1 if it is not a register.
Public Function RegIndex(ByVal nm As String) As Integer
    Select Case UCase$(Trim$(nm))
        Case "AX": RegIndex = R_AX
        Case "BX": RegIndex = R_BX
        Case "CX": RegIndex = R_CX
        Case "DX": RegIndex = R_DX
        Case Else: RegIndex = -1
    End Select
End Function

' =====================================================================
' Test_CPU - self-check of the register file.
' Run it from the VBA editor (F5); results go to the Immediate window.
' =====================================================================
Public Sub Test_CPU()
    Dim fails As Integer, i As Integer, ok As Boolean

    ' General registers
    ok = True
    For i = R_AX To R_DX
        SetReg i, 10 + i
        If GetReg(i) <> 10 + i Then ok = False
    Next i
    fails = fails + CpuCheck("GetReg/SetReg for AX, BX, CX, DX", ok)
    SetReg R_AX, 256
    fails = fails + CpuCheck("SetReg 256 stores 0", GetReg(R_AX) = 0)
    SetReg R_BX, -1
    fails = fails + CpuCheck("SetReg -1 stores 255", GetReg(R_BX) = 255)
    SetReg R_CX, 300
    fails = fails + CpuCheck("SetReg 300 stores 44", GetReg(R_CX) = 44)

    ' Control registers
    SetPC 511: IncPC
    fails = fails + CpuCheck("PC wraps from 511 to 0", CPU.PC = 0)
    SetPC 512
    fails = fails + CpuCheck("SetPC 512 stores 0", CPU.PC = 0)
    SetMAR -1
    fails = fails + CpuCheck("MAR is 9 bits (-1 -> 511)", CPU.MAR = 511)
    SetMDR 256
    fails = fails + CpuCheck("MDR is 8 bits (256 -> 0)", CPU.MDR = 0)
    SetIR &H12801
    fails = fails + CpuCheck("IR keeps 16 bits (0x12801 -> 0x2801)", CPU.IR = &H2801&)

    ' Names
    fails = fails + CpuCheck("RegName 0-3 = AX BX CX DX", RegName(0) & RegName(1) & RegName(2) & RegName(3) = "AXBXCXDX")
    fails = fails + CpuCheck("RegIndex(""CX"") = 2", RegIndex("CX") = 2)
    fails = fails + CpuCheck("RegIndex ignores case and spaces", RegIndex(" dx ") = R_DX)
    fails = fails + CpuCheck("RegIndex(""EX"") = -1", RegIndex("EX") = -1)

    ' Invalid index raises an error
    On Error Resume Next
    i = GetReg(4)
    ok = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    fails = fails + CpuCheck("GetReg(4) raises an error", ok)

    ' Reset
    CPU.ZF = True: CPU.CF = True: CPU.SF = True: CPU.Halted = True: SetPC 100
    ResetCPU
    fails = fails + CpuCheck("ResetCPU: PC = 0, flags clear, registers 0", _
        CPU.PC = 0 And Not CPU.ZF And Not CPU.CF And Not CPU.SF And Not CPU.Halted And GetReg(R_BX) = 0)

    Debug.Print IIf(fails = 0, "Test_CPU: ALL OK", "Test_CPU: " & fails & " FAIL(S)")
End Sub

Private Function CpuCheck(ByVal what As String, ByVal passed As Boolean) As Integer
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    CpuCheck = IIf(passed, 0, 1)
End Function
