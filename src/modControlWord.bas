Attribute VB_Name = "modControlWord"
' =====================================================================
' modControlWord - 27-bit control word + exclusion rules (G1)
' Owner: Adriana  |  Task: #7 (A7)
'
' Bits live in a Long. ValidateControlWord enforces single data driver,
' single address driver, and single external access (Mem/IO).
' =====================================================================
Option Explicit

Public Const SIG_EN_PC As Long = &H1&
Public Const SIG_LD_PC As Long = &H2&
Public Const SIG_INC_PC As Long = &H4&
Public Const SIG_LD_MAR As Long = &H8&
Public Const SIG_LD_MDR As Long = &H10&
Public Const SIG_EN_MDR As Long = &H20&
Public Const SIG_LD_IR As Long = &H40&
Public Const SIG_EN_IR_A As Long = &H80&          ' IR address -> address bus
Public Const SIG_EN_AX As Long = &H100&
Public Const SIG_EN_BX As Long = &H200&
Public Const SIG_EN_CX As Long = &H400&
Public Const SIG_EN_DX As Long = &H800&
Public Const SIG_LD_AX As Long = &H1000&
Public Const SIG_LD_BX As Long = &H2000&
Public Const SIG_LD_CX As Long = &H4000&
Public Const SIG_LD_DX As Long = &H8000&
Public Const SIG_LD_TA As Long = &H10000
Public Const SIG_LD_TB As Long = &H20000
Public Const SIG_EN_ALU As Long = &H40000
Public Const SIG_LD_FLAGS As Long = &H80000
Public Const SIG_MEM_R As Long = &H100000
Public Const SIG_MEM_W As Long = &H200000
Public Const SIG_IO_R As Long = &H400000
Public Const SIG_IO_W As Long = &H800000
Public Const SIG_EN_KBD As Long = &H1000000
Public Const SIG_LD_LCD As Long = &H2000000
Public Const SIG_EN_IR_D As Long = &H4000000     ' IR immediate -> data bus
Public Const NUM_SIGNALS As Integer = 27

' At most one active in each mask
Public Const MASK_DATA_DRIVERS As Long = SIG_EN_MDR Or SIG_EN_IR_D Or SIG_EN_AX Or _
    SIG_EN_BX Or SIG_EN_CX Or SIG_EN_DX Or SIG_EN_ALU Or SIG_EN_KBD
Public Const MASK_ADDR_DRIVERS As Long = SIG_EN_PC Or SIG_EN_IR_A
Public Const MASK_EXTERNAL As Long = SIG_MEM_R Or SIG_MEM_W Or SIG_IO_R Or SIG_IO_W

Private Const CW_ISA_START_ROW As Long = 42   ' below the ISA instruction table

Public Function SignalName(ByVal bitIndex As Integer) As String
    Select Case bitIndex
        Case 0: SignalName = "En_PC"
        Case 1: SignalName = "Ld_PC"
        Case 2: SignalName = "Inc_PC"
        Case 3: SignalName = "Ld_MAR"
        Case 4: SignalName = "Ld_MDR"
        Case 5: SignalName = "En_MDR"
        Case 6: SignalName = "Ld_IR"
        Case 7: SignalName = "En_IR_A"
        Case 8: SignalName = "En_AX"
        Case 9: SignalName = "En_BX"
        Case 10: SignalName = "En_CX"
        Case 11: SignalName = "En_DX"
        Case 12: SignalName = "Ld_AX"
        Case 13: SignalName = "Ld_BX"
        Case 14: SignalName = "Ld_CX"
        Case 15: SignalName = "Ld_DX"
        Case 16: SignalName = "Ld_TA"
        Case 17: SignalName = "Ld_TB"
        Case 18: SignalName = "En_ALU"
        Case 19: SignalName = "Ld_FLAGS"
        Case 20: SignalName = "Mem_R"
        Case 21: SignalName = "Mem_W"
        Case 22: SignalName = "IO_R"
        Case 23: SignalName = "IO_W"
        Case 24: SignalName = "En_KBD"
        Case 25: SignalName = "Ld_LCD"
        Case 26: SignalName = "En_IR_D"
        Case Else: SignalName = ""
    End Select
End Function

Public Function SignalMeaning(ByVal bitIndex As Integer) As String
    Select Case bitIndex
        Case 0: SignalMeaning = "PC drives internal address bus"
        Case 1: SignalMeaning = "Load PC from address bus"
        Case 2: SignalMeaning = "PC <- PC + 1"
        Case 3: SignalMeaning = "Load MAR from address bus"
        Case 4: SignalMeaning = "Load MDR from data bus"
        Case 5: SignalMeaning = "MDR drives data bus"
        Case 6: SignalMeaning = "Load IR from data bus"
        Case 7: SignalMeaning = "IR address field drives address bus"
        Case 8: SignalMeaning = "AX drives data bus"
        Case 9: SignalMeaning = "BX drives data bus"
        Case 10: SignalMeaning = "CX drives data bus"
        Case 11: SignalMeaning = "DX drives data bus"
        Case 12: SignalMeaning = "Load AX from data bus"
        Case 13: SignalMeaning = "Load BX from data bus"
        Case 14: SignalMeaning = "Load CX from data bus"
        Case 15: SignalMeaning = "Load DX from data bus"
        Case 16: SignalMeaning = "Load TA (ALU temp A)"
        Case 17: SignalMeaning = "Load TB (ALU temp B)"
        Case 18: SignalMeaning = "ALU result drives data bus"
        Case 19: SignalMeaning = "Update ZF/CF/SF from ALU"
        Case 20: SignalMeaning = "Memory read"
        Case 21: SignalMeaning = "Memory write"
        Case 22: SignalMeaning = "I/O read"
        Case 23: SignalMeaning = "I/O write"
        Case 24: SignalMeaning = "Keyboard drives data bus"
        Case 25: SignalMeaning = "Load LCD from data bus"
        Case 26: SignalMeaning = "IR immediate drives data bus"
        Case Else: SignalMeaning = ""
    End Select
End Function

Private Function BitMask(ByVal bitIndex As Integer) As Long
    BitMask = CLng(2 ^ bitIndex)
End Function

Private Function CountBits(ByVal v As Long) As Integer
    Dim n As Integer
    Dim u As Long
    u = v
    Do While u <> 0
        n = n + (u And 1&)
        u = u \ 2
    Loop
    CountBits = n
End Function

' Returns "" when valid, otherwise the error message.
Public Function ValidateControlWord(ByVal cw As Long) As String
    If CountBits(cw And MASK_DATA_DRIVERS) > 1 Then
        ValidateControlWord = "More than one driver on the data bus: " & CW_ToString(cw And MASK_DATA_DRIVERS)
    ElseIf CountBits(cw And MASK_ADDR_DRIVERS) > 1 Then
        ValidateControlWord = "More than one driver on the address bus: " & CW_ToString(cw And MASK_ADDR_DRIVERS)
    ElseIf CountBits(cw And MASK_EXTERNAL) > 1 Then
        ValidateControlWord = "MEM_R + MEM_W + IO_R + IO_W > 1: " & CW_ToString(cw And MASK_EXTERNAL)
    Else
        ValidateControlWord = ""
    End If
End Function

Public Function CW_ToString(ByVal cw As Long) As String
    Dim i As Integer
    Dim s As String
    For i = 0 To NUM_SIGNALS - 1
        If (cw And BitMask(i)) <> 0 Then s = s & SignalName(i) & " "
    Next i
    CW_ToString = Trim$(s)
End Function

Public Function CW_ToBinary(ByVal cw As Long) As String
    Dim i As Integer
    Dim s As String
    For i = NUM_SIGNALS - 1 To 0 Step -1
        s = s & IIf((cw And BitMask(i)) <> 0, "1", "0")
    Next i
    CW_ToBinary = s
End Function

Public Function SIG_EN_REG(ByVal r As Integer) As Long
    If r < 0 Or r > 3 Then Err.Raise vbObjectError + 701, "modControlWord", "SIG_EN_REG: bad register " & r
    SIG_EN_REG = SIG_EN_AX * CLng(2 ^ r)
End Function

Public Function SIG_LD_REG(ByVal r As Integer) As Long
    If r < 0 Or r > 3 Then Err.Raise vbObjectError + 702, "modControlWord", "SIG_LD_REG: bad register " & r
    SIG_LD_REG = SIG_LD_AX * CLng(2 ^ r)
End Function

' =====================================================================
' ISA sheet: 27-signal table + 3 rules + Force collision button (G1 demo)
' =====================================================================
Public Sub PaintControlWordOnISA()
    Dim ws As Worksheet
    Dim i As Integer
    Dim r As Long
    Dim btn As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_ISA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_ISA
    End If

    Application.ScreenUpdating = False
    r = CW_ISA_START_ROW

    ws.Cells(r, 2).Value = "Control word (27 signals)"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Size = 12
    r = r + 1
    ws.Cells(r, 2).Resize(1, 4).Value = Array("Bit", "Signal", "Mask", "Meaning")
    ws.Cells(r, 2).Resize(1, 4).Font.Bold = True
    r = r + 1
    For i = 0 To NUM_SIGNALS - 1
        ws.Cells(r, 2).Value = i
        ws.Cells(r, 3).Value = SignalName(i)
        ws.Cells(r, 4).Value = "0x" & Hex$(BitMask(i))
        ws.Cells(r, 5).Value = SignalMeaning(i)
        r = r + 1
    Next i

    r = r + 1
    ws.Cells(r, 2).Value = "Validation rules (every micro-step)"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    ws.Cells(r, 2).Value = "1. Single data-bus driver: among En_MDR, En_IR_D, En_AX..En_DX, En_ALU, En_KBD at most one is 1."
    r = r + 1
    ws.Cells(r, 2).Value = "2. Single address-bus driver: among En_PC, En_IR_A at most one is 1."
    r = r + 1
    ws.Cells(r, 2).Value = "3. Single external access: Mem_R + Mem_W + IO_R + IO_W <= 1."
    r = r + 1
    ws.Cells(r, 2).Value = "Note: several Ld_* may be active at once (multiple receivers)."

    On Error Resume Next
    ws.Shapes("btn_ForceCollision").Delete
    On Error GoTo 0
    Set btn = UI_AddBox(ws, "btn_ForceCollision", 20, 10, 160, 28, "FORCE COLLISION", CLR_MISS)
    btn.OnAction = "Demo_ForceCollision"

    Application.ScreenUpdating = True
    Debug.Print "PaintControlWordOnISA: OK"
End Sub

' G1 demo: deliberately invalid control word -> show validation error.
Public Sub Demo_ForceCollision()
    Dim cw As Long
    Dim msg As String
    cw = SIG_EN_AX Or SIG_EN_BX          ' two data drivers
    msg = ValidateControlWord(cw)
    If Len(msg) = 0 Then
        MsgBox "Unexpected: invalid word was accepted.", vbExclamation, "Byte Bakers"
    Else
        MsgBox "Forced collision demo" & vbCrLf & vbCrLf & _
               "CW: " & CW_ToString(cw) & vbCrLf & _
               "Binary: " & CW_ToBinary(cw) & vbCrLf & vbCrLf & _
               msg, vbCritical, "Byte Bakers - Bus conflict"
    End If
End Sub

' =====================================================================
' Test_ControlWord - Immediate self-check for A7 / #7.
' =====================================================================
Public Function Test_ControlWord() As Long
    Dim fails As Long
    Dim msg As String
    Dim i As Integer
    Dim names As String

    ' --- 27 named signals ---
    For i = 0 To NUM_SIGNALS - 1
        If Len(SignalName(i)) = 0 Then fails = fails + CwCheck("SignalName(" & i & ") non-empty", False)
    Next i
    fails = fails + CwCheck("NUM_SIGNALS = 27", NUM_SIGNALS = 27)
    fails = fails + CwCheck("SIG_EN_IR_D is bit 26", SIG_EN_IR_D = CLng(2 ^ 26))

    ' --- valid words ---
    fails = fails + CwCheck("valid En_PC Ld_MAR", ValidateControlWord(SIG_EN_PC Or SIG_LD_MAR) = "")
    fails = fails + CwCheck("valid En_AX Ld_TA", ValidateControlWord(SIG_EN_AX Or SIG_LD_TA) = "")
    fails = fails + CwCheck("valid Mem_R Ld_MDR", ValidateControlWord(SIG_MEM_R Or SIG_LD_MDR) = "")
    fails = fails + CwCheck("several Ld_* OK", ValidateControlWord(SIG_EN_ALU Or SIG_LD_AX Or SIG_LD_FLAGS) = "")

    ' --- data bus exclusion ---
    msg = ValidateControlWord(SIG_EN_AX Or SIG_EN_BX)
    fails = fails + CwCheck("reject two data drivers", InStr(1, msg, "data bus", vbTextCompare) > 0)

    ' --- address bus exclusion ---
    msg = ValidateControlWord(SIG_EN_PC Or SIG_EN_IR_A)
    fails = fails + CwCheck("reject two addr drivers", InStr(1, msg, "address bus", vbTextCompare) > 0)

    ' --- external exclusion ---
    msg = ValidateControlWord(SIG_MEM_R Or SIG_MEM_W)
    fails = fails + CwCheck("reject Mem_R+Mem_W", InStr(1, msg, "MEM_R", vbTextCompare) > 0)
    msg = ValidateControlWord(SIG_IO_R Or SIG_MEM_R)
    fails = fails + CwCheck("reject IO_R+Mem_R", Len(msg) > 0)

    ' --- formatting ---
    names = CW_ToString(SIG_EN_PC Or SIG_LD_MAR)
    fails = fails + CwCheck("CW_ToString En_PC Ld_MAR", names = "En_PC Ld_MAR")
    fails = fails + CwCheck("CW_ToBinary length 27", Len(CW_ToBinary(0)) = 27)
    fails = fails + CwCheck("CW_ToBinary bit0 set", Right$(CW_ToBinary(SIG_EN_PC), 1) = "1")

    ' --- helpers ---
    fails = fails + CwCheck("SIG_EN_REG(0)=En_AX", SIG_EN_REG(0) = SIG_EN_AX)
    fails = fails + CwCheck("SIG_LD_REG(3)=Ld_DX", SIG_LD_REG(3) = SIG_LD_DX)

    ' --- sheet paint (no raise) ---
    On Error Resume Next
    PaintControlWordOnISA
    fails = fails + CwCheck("PaintControlWordOnISA no raise", Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    Debug.Print IIf(fails = 0, "Test_ControlWord: ALL OK", "Test_ControlWord: " & fails & " FAIL(S)")
    Test_ControlWord = fails
End Function

Private Function CwCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    CwCheck = IIf(passed, 0, 1)
End Function
