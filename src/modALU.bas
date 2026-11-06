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

' =====================================================================
' Test_ALU - self-check of the ALU and flags (G5 / A5 acceptance).
' Run from the VBA editor (F5); results go to the Immediate window.
' Returns the number of failed checks (0 = ALL OK) for automated verify.
' =====================================================================
Public Function Test_ALU() As Long
    Dim fails As Long
    Dim beforeAX As Integer
    Dim cfBefore As Boolean

    ResetCPU

    ' --- Acceptance: 200 + 100 = 44, CF = 1 ---
    CPU.TA = 200: CPU.TB = 100
    fails = fails + AluCheck("200+100 = 44 with CF=1", _
        AluExecute(OP_ADD_RR) = 44 And CPU.CF And Not CPU.ZF And Not CPU.SF)

    ' --- Acceptance: 5 - 5 = 0, ZF = 1 ---
    CPU.TA = 5: CPU.TB = 5
    fails = fails + AluCheck("5-5 = 0 with ZF=1", _
        AluExecute(OP_SUB_RR) = 0 And CPU.ZF And Not CPU.CF And Not CPU.SF)

    ' --- Acceptance: 3 - 5 = 254, CF = 1, SF = 1 ---
    CPU.TA = 3: CPU.TB = 5
    fails = fails + AluCheck("3-5 = 254 with CF=1 and SF=1", _
        AluExecute(OP_SUB_RR) = 254 And CPU.CF And CPU.SF And Not CPU.ZF)

    ' --- ADD immediate form uses the same path ---
    CPU.TA = 200: CPU.TB = 100
    fails = fails + AluCheck("ADD_RI same as ADD_RR", _
        AluExecute(OP_ADD_RI) = 44 And CPU.CF)

    ' --- INC / DEC: result and Z/S; CF unchanged ---
    CPU.CF = True
    CPU.TA = 255: CPU.TB = 0
    fails = fails + AluCheck("INC 255 = 0, ZF=1, CF stays 1", _
        AluExecute(OP_INC) = 0 And CPU.ZF And CPU.CF And Not CPU.SF)

    CPU.CF = False
    CPU.TA = 127: CPU.TB = 0
    fails = fails + AluCheck("INC 127 = 128, SF=1, CF stays 0", _
        AluExecute(OP_INC) = 128 And CPU.SF And Not CPU.CF And Not CPU.ZF)

    CPU.CF = True
    CPU.TA = 1: CPU.TB = 0
    fails = fails + AluCheck("DEC 1 = 0, ZF=1, CF stays 1", _
        AluExecute(OP_DEC) = 0 And CPU.ZF And CPU.CF)

    ' --- Logic: AND / OR / XOR ---
    cfBefore = True
    CPU.CF = cfBefore
    CPU.TA = &HF0: CPU.TB = &H0F
    fails = fails + AluCheck("AND F0,0F = 0, ZF=1, CF unchanged", _
        AluExecute(OP_AND) = 0 And CPU.ZF And CPU.CF = cfBefore)

    CPU.CF = False
    CPU.TA = &H80: CPU.TB = &H7F
    fails = fails + AluCheck("OR 80,7F = FF, SF=1, CF unchanged", _
        AluExecute(OP_OR) = 255 And CPU.SF And Not CPU.CF And Not CPU.ZF)

    CPU.TA = &HFF: CPU.TB = &HFF
    fails = fails + AluCheck("XOR FF,FF = 0, ZF=1", _
        AluExecute(OP_XOR) = 0 And CPU.ZF And Not CPU.SF)

    ' --- Acceptance: CMP does not modify the destination register ---
    ' AluExecute only returns the difference and flags; it never writes R().
    SetReg R_AX, 77
    beforeAX = GetReg(R_AX)
    CPU.TA = 3: CPU.TB = 5
    fails = fails + AluCheck("CMP 3,5 flags CF+SF, result 254", _
        AluExecute(OP_CMP_RR) = 254 And CPU.CF And CPU.SF And Not CPU.ZF)
    fails = fails + AluCheck("CMP does not modify destination register AX", _
        GetReg(R_AX) = beforeAX)

    CPU.TA = 10: CPU.TB = 10
    fails = fails + AluCheck("CMP_RI equal sets ZF", _
        AluExecute(OP_CMP_RI) = 0 And CPU.ZF And Not CPU.CF)

    Debug.Print IIf(fails = 0, "Test_ALU: ALL OK", "Test_ALU: " & fails & " FAIL(S)")
    Test_ALU = fails
End Function

Private Function AluCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    AluCheck = IIf(passed, 0, 1)
End Function
