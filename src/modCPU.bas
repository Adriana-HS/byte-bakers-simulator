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

