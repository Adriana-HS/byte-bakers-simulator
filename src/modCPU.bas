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

