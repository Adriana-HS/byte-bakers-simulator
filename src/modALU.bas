Attribute VB_Name = "modALU"
' =====================================================================
' modALU - Arithmetic Logic Unit of the Byte Bakers CPU
' Owner: Adriana  |  Task: #5 (A5)
'
' Operates on the temporaries CPU.TA and CPU.TB, updates ZF / CF / SF
' according to the ISA flag table, and returns the 8-bit result.
' The control unit decides whether to write that result back to a
' register (CMP must not; ADD/SUB/INC/DEC/AND/OR/XOR must).
' =====================================================================
Option Explicit

' ---------------------------------------------------------------------
' AluExecute - run one ALU operation.
'   ADD/SUB set CF (carry / borrow). INC/DEC and logic leave CF alone
'   (ISA: only Z and S). ZF and SF always follow the wrapped result.
' ---------------------------------------------------------------------
Public Function AluExecute(ByVal op As Integer) As Integer
    Dim a As Long, b As Long, r As Long

    a = CPU.TA
    b = CPU.TB

    Select Case op
        Case OP_ADD_RR, OP_ADD_RI
            r = a + b
            CPU.CF = (r > 255)

        Case OP_SUB_RR, OP_SUB_RI, OP_CMP_RR, OP_CMP_RI
            r = a - b
            CPU.CF = (r < 0)

        Case OP_INC
            r = a + 1          ' CF unchanged (ISA: Z S only)

        Case OP_DEC
            r = a - 1          ' CF unchanged (ISA: Z S only)

        Case OP_AND
            r = a And b        ' CF unchanged (ISA: Z S only)

        Case OP_OR
            r = a Or b

        Case OP_XOR
            r = a Xor b

        Case Else
            Err.Raise vbObjectError + 610, "modALU", _
                "AluExecute: opcode " & op & " is not an ALU operation"
    End Select

    r = Wrap8(r)
    CPU.ZF = (r = 0)
    CPU.SF = (r >= 128)
    AluExecute = CInt(r)
End Function
