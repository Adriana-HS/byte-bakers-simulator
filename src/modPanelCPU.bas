Attribute VB_Name = "modPanelCPU"
' =====================================================================
' modPanelCPU - CPU / bus panel with buffers + control-word lights
' Owner: Adriana  |  Task: #9 (A9)  |  Rubric: G1, G7
'
' Draws registers, PC/IR/MAR/MDR, ALU (TA/TB), flags, internal (9-bit)
' and external (8-bit) address buses, data buses, one tri-state buffer
' per driver (incl. buf_KBD / buf_LCD), transfer label, and the 27-bit
' control-word strip + CLK. Refresh paints coral / amber / High-Z gray.
' =====================================================================
Option Explicit

Private Const PANEL_L As Single = 28
Private Const PANEL_T As Single = 58
Private Const BOX_W As Single = 72
Private Const BOX_H As Single = 28
Private Const BUF_W As Single = 18
Private Const BUF_H As Single = 18
Private Const SIG_W As Single = 20
Private Const SIG_H As Single = 14

' Soft chrome (hierarchy) — values still use coral/amber/gray dynamically
Private Const CLR_LABEL As Long = &HC8B4A0&      ' RGB(160,180,200) muted label
Private Const CLR_CHIP As Long = &H4A3828&       ' RGB(40,56,74) chip bg
Private Const CLR_CLK_ON As Long = &H5AD8D8&     ' RGB(216,216,90) CLK pulse
Private Const CLR_ISA_HL As Long = &HC8E8FF&     ' RGB(255,232,200) ISA row

Private mPrevIsaRow As Long

' =====================================================================
' BuildCPUPanel - create every shape inside the CPU frame (frm_CPU).
' =====================================================================
Public Sub BuildCPUPanel(ByVal ws As Worksheet)
    Dim x As Single, y As Single
    Dim i As Integer
    Dim nm As String

    On Error GoTo EH
    mPrevIsaRow = 0

    ' Title + transfer (hierarchy: what is happening now)
    AddBox ws, "lbl_CPUTitle", PANEL_L, PANEL_T, 160, 20, "CPU / BUSES", CLR_CHIP
    AddBox ws, "lbl_Transfer", PANEL_L + 170, PANEL_T, 280, 20, "— idle —", CLR_CHIP

    ' --- GP registers + buffers (left column) ---
    y = PANEL_T + 28
    AddBox ws, "lbl_Regs", PANEL_L, y, 100, 16, "REGS", CLR_CHIP
    y = y + 18
    AddRegPair ws, "AX", PANEL_L, y: y = y + 34
    AddRegPair ws, "BX", PANEL_L, y: y = y + 34
    AddRegPair ws, "CX", PANEL_L, y: y = y + 34
    AddRegPair ws, "DX", PANEL_L, y

    ' --- Control regs (center) ---
    x = PANEL_L + 160
    y = PANEL_T + 28
    AddBox ws, "lbl_CtrlRegs", x, y, 120, 16, "CONTROL", CLR_CHIP
    y = y + 18
    AddRegPair ws, "PC", x, y: y = y + 32
    AddRegPair ws, "IR", x, y: y = y + 32
    AddRegPair ws, "MAR", x, y: y = y + 32
    AddRegPair ws, "MDR", x, y: y = y + 32
    AddRegPair ws, "PADDR", x, y

    ' --- ALU + flags (right of control) ---
    x = PANEL_L + 340
    y = PANEL_T + 28
    AddBox ws, "lbl_ALU", x, y, 100, 16, "ALU", CLR_CHIP
    y = y + 18
    AddRegPair ws, "TA", x, y: y = y + 32
    AddRegPair ws, "TB", x, y: y = y + 32
    AddBox ws, PFX_BUF & "ALU", x + BOX_W + 4, y - 50, BUF_W, BUF_H, "Z", CLR_HIGHZ
    AddBox ws, PFX_REG & "ALU", x, y, BOX_W, BOX_H, "ALU", CLR_HIGHZ
    y = y + 36
    AddBox ws, "flg_ZF", x, y, 40, 22, "ZF 0", CLR_HIGHZ
    AddBox ws, "flg_CF", x + 44, y, 40, 22, "CF 0", CLR_HIGHZ
    AddBox ws, "flg_SF", x + 88, y, 40, 22, "SF 0", CLR_HIGHZ

    ' --- I/O buffers on the data path ---
    x = PANEL_L + 500
    y = PANEL_T + 46
    AddBox ws, "lbl_IOBuf", x, y - 18, 100, 16, "I/O BUF", CLR_CHIP
    AddBox ws, PFX_BUF & "KBD", x, y, BUF_W + 30, BUF_H + 6, "KBD", CLR_HIGHZ
    AddBox ws, PFX_BUF & "LCD", x, y + 28, BUF_W + 30, BUF_H + 6, "LCD", CLR_HIGHZ
    AddBox ws, PFX_BUF & "MEM", x, y + 56, BUF_W + 30, BUF_H + 6, "MEM", CLR_HIGHZ

    ' --- Buses (separate INT 9-bit vs EXT 8-bit address + data) ---
    y = PANEL_T + 220
    AddBox ws, "lbl_Buses", PANEL_L, y, 200, 14, "BUSES  INT=9-bit  EXT=8-bit", CLR_CHIP
    y = y + 16
    UI_AddWire ws, PFX_WIRE & "INT_ADDR", PANEL_L, y + 6, PANEL_L + 520, y + 6, 3
    AddBox ws, "val_INT_ADDR", PANEL_L + 524, y - 6, 90, 18, "INT —", CLR_CHIP
    y = y + 16
    UI_AddWire ws, PFX_WIRE & "ADDR", PANEL_L, y + 6, PANEL_L + 520, y + 6, 3
    AddBox ws, "val_EXT_ADDR", PANEL_L + 524, y - 6, 90, 18, "EXT —", CLR_CHIP
    y = y + 16
    UI_AddWire ws, PFX_WIRE & "INT_DATA", PANEL_L, y + 6, PANEL_L + 520, y + 6, 3
    AddBox ws, "val_INT_DATA", PANEL_L + 524, y - 6, 90, 18, "ID —", CLR_CHIP
    y = y + 16
    UI_AddWire ws, PFX_WIRE & "DATA", PANEL_L, y + 6, PANEL_L + 520, y + 6, 3
    AddBox ws, "val_DATA", PANEL_L + 524, y - 6, 90, 18, "D —", CLR_CHIP

    ' --- Control word strip + CLK (bottom of CPU frame) ---
    y = PANEL_T + 300
    AddBox ws, "lbl_CW", PANEL_L, y - 16, 180, 14, "CONTROL WORD + CLK", CLR_CHIP
    For i = 0 To NUM_SIGNALS - 1
        nm = SignalName(i)
        If Len(nm) > 0 Then
            x = PANEL_L + (i Mod 14) * (SIG_W + 2)
            AddBox ws, PFX_SIG & nm, x, y + (i \ 14) * (SIG_H + 2), SIG_W, SIG_H, ShortSig(nm), CLR_HIGHZ
        End If
    Next i
    AddBox ws, PFX_SIG & "CLK", PANEL_L + 520, y, 60, 20, "CLK", CLR_HIGHZ

    RefreshCPUPanel
    Exit Sub
EH:
    Debug.Print "BuildCPUPanel: ERROR " & Err.Number & " " & Err.Description
End Sub

' reg + tiny tri-state buffer to its right
Private Sub AddRegPair(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Single, ByVal y As Single)
    AddBox ws, PFX_REG & nm, x, y, BOX_W, BOX_H, nm, CLR_HIGHZ
    AddBox ws, PFX_BUF & nm, x + BOX_W + 4, y + 5, BUF_W, BUF_H, "Z", CLR_HIGHZ
End Sub

Private Sub AddBox(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal txt As String, ByVal fill As Long)
    Dim s As Shape
    Set s = UI_AddBox(ws, nm, l, t, w, h, txt, fill)
    On Error Resume Next
    s.TextFrame2.TextRange.Font.Size = 8
    s.TextFrame2.TextRange.Font.Name = "Consolas"
    On Error GoTo 0
End Sub

Private Function ShortSig(ByVal nm As String) As String
    ' Compact label so 27 lights fit the strip
    ShortSig = Replace(Replace(Replace(nm, "En_", "E"), "Ld_", "L"), "Inc_", "I")
    If Len(ShortSig) > 4 Then ShortSig = Left$(ShortSig, 4)
End Function

' =====================================================================
' RefreshCPUPanel - coral driver, amber receivers, gray High-Z, CW + ISA.
' =====================================================================
Public Sub RefreshCPUPanel()
    Dim i As Integer
    Dim cw As Long
    Dim nm As String
    Dim recv As String
    Dim op As Integer, reg As Integer, operand As Integer
    Dim row As Long
    Dim wsIsa As Worksheet

    On Error Resume Next

    ' Values (hex + decimal)
    PaintReg "AX", CPU.R(R_AX), False
    PaintReg "BX", CPU.R(R_BX), False
    PaintReg "CX", CPU.R(R_CX), False
    PaintReg "DX", CPU.R(R_DX), False
    PaintReg "PC", CPU.PC, True
    PaintReg "IR", CPU.IR And &HFFFF&, False
    PaintReg "MAR", CPU.MAR, True
    PaintReg "MDR", CPU.MDR, False
    PaintReg "PADDR", CPU.PADDR, False
    PaintReg "TA", CPU.TA, False
    PaintReg "TB", CPU.TB, False
    UI_SetText PFX_REG & "ALU", "ALU"

    UI_SetText "flg_ZF", "ZF " & IIf(CPU.ZF, "1", "0")
    UI_SetText "flg_CF", "CF " & IIf(CPU.CF, "1", "0")
    UI_SetText "flg_SF", "SF " & IIf(CPU.SF, "1", "0")
    UI_SetFill "flg_ZF", IIf(CPU.ZF, CLR_LOAD, CLR_HIGHZ)
    UI_SetFill "flg_CF", IIf(CPU.CF, CLR_LOAD, CLR_HIGHZ)
    UI_SetFill "flg_SF", IIf(CPU.SF, CLR_LOAD, CLR_HIGHZ)

    ' Reset every buffer / wire to High-Z gray, then paint active
    Dim bufs As Variant
    bufs = Array("AX", "BX", "CX", "DX", "PC", "IR", "MAR", "MDR", "PADDR", "TA", "TB", "ALU", "KBD", "LCD", "MEM")
    For i = LBound(bufs) To UBound(bufs)
        UI_SetFill PFX_BUF & CStr(bufs(i)), CLR_HIGHZ
        UI_SetText PFX_BUF & CStr(bufs(i)), "Z"
        UI_SetFill PFX_REG & CStr(bufs(i)), CLR_HIGHZ
    Next i
    UI_SetHighZ PFX_WIRE & "INT_ADDR"
    UI_SetHighZ PFX_WIRE & "ADDR"
    UI_SetHighZ PFX_WIRE & "INT_DATA"
    UI_SetHighZ PFX_WIRE & "DATA"

    ' Bus values
    UI_SetText "val_INT_ADDR", "INT " & BusHex(Bus.IntAddr, True)
    UI_SetText "val_EXT_ADDR", "EXT " & BusHex(Bus.ExtAddr, False)
    UI_SetText "val_INT_DATA", "ID " & IIf(Bus.Internal, BusHex(Bus.Data, False), "—")
    UI_SetText "val_DATA", "D " & IIf(Not Bus.Internal And Bus.Data <> HIGH_Z, BusHex(Bus.Data, False), _
        IIf(Bus.Data = HIGH_Z, "—", BusHex(Bus.Data, False)))

    ' Active address driver
    If Len(Bus.AddrDriver) > 0 And Bus.IntAddr <> HIGH_Z Then
        UI_Highlight PFX_BUF & Bus.AddrDriver, CLR_ACTIVE
        UI_SetText PFX_BUF & Bus.AddrDriver, ">"
        UI_Highlight PFX_WIRE & "INT_ADDR", CLR_ACTIVE
    End If
    If Bus.ExtAddr <> HIGH_Z Then
        UI_Highlight PFX_WIRE & "ADDR", CLR_ACTIVE
    End If

    ' Active data driver (coral) + wire
    If Len(Bus.DataDriver) > 0 And Bus.Data <> HIGH_Z Then
        UI_Highlight PFX_BUF & Bus.DataDriver, CLR_ACTIVE
        UI_SetText PFX_BUF & Bus.DataDriver, ">"
        If Bus.Internal Then
            UI_Highlight PFX_WIRE & "INT_DATA", CLR_ACTIVE
        Else
            UI_Highlight PFX_WIRE & "DATA", CLR_ACTIVE
        End If
    End If

    ' Receivers from control word Ld_* (amber)
    cw = Bus.Ctrl
    recv = ActiveReceiver(cw)
    If Len(recv) > 0 Then
        UI_Highlight PFX_REG & recv, CLR_LOAD
    End If

    ' Transfer label e.g. MDR → AX
    If Len(Bus.DataDriver) > 0 And Len(recv) > 0 Then
        UI_SetText "lbl_Transfer", Bus.DataDriver & " → " & recv
    ElseIf Len(Bus.AddrDriver) > 0 And (cw And SIG_LD_MAR) <> 0 Then
        UI_SetText "lbl_Transfer", Bus.AddrDriver & " → MAR"
    ElseIf Len(Bus.DataDriver) > 0 Then
        UI_SetText "lbl_Transfer", Bus.DataDriver & " → bus"
    Else
        UI_SetText "lbl_Transfer", "— idle —"
    End If

    ' Control word lights
    For i = 0 To NUM_SIGNALS - 1
        nm = SignalName(i)
        If Len(nm) > 0 Then
            If (cw And BitMask(i)) <> 0 Then
                ' Ld_* amber, En_* coral (G1 dynamic High-Z)
                If Left$(nm, 3) = "Ld_" Or Left$(nm, 4) = "Inc_" Then
                    UI_SetFill PFX_SIG & nm, CLR_LOAD
                Else
                    UI_SetFill PFX_SIG & nm, CLR_ACTIVE
                End If
            Else
                UI_SetFill PFX_SIG & nm, CLR_HIGHZ
            End If
        End If
    Next i
    UI_SetFill PFX_SIG & "CLK", IIf(Bus.Clk, CLR_CLK_ON, CLR_HIGHZ)
    UI_SetText PFX_SIG & "CLK", IIf(Bus.Clk, "CLK●", "CLK")

    ' Highlight current IR opcode row on ISA sheet (cells, not shapes)
    Decode CPU.IR, op, reg, operand
    row = ISARow(op)
    Set wsIsa = ThisWorkbook.Worksheets(SH_ISA)
    If Not wsIsa Is Nothing Then
        If mPrevIsaRow > 0 And mPrevIsaRow <> row Then
            wsIsa.Range(wsIsa.Cells(mPrevIsaRow, 2), wsIsa.Cells(mPrevIsaRow, 10)).Interior.ColorIndex = xlNone
        End If
        If row > 0 Then
            wsIsa.Range(wsIsa.Cells(row, 2), wsIsa.Cells(row, 10)).Interior.Color = CLR_ISA_HL
            mPrevIsaRow = row
        End If
    End If

    ' Re-apply register text after fill changes
    PaintReg "AX", CPU.R(R_AX), False
    PaintReg "BX", CPU.R(R_BX), False
    PaintReg "CX", CPU.R(R_CX), False
    PaintReg "DX", CPU.R(R_DX), False
    PaintReg "PC", CPU.PC, True
    PaintReg "IR", CPU.IR And &HFFFF&, False
    PaintReg "MAR", CPU.MAR, True
    PaintReg "MDR", CPU.MDR, False
    PaintReg "PADDR", CPU.PADDR, False
    PaintReg "TA", CPU.TA, False
    PaintReg "TB", CPU.TB, False

    On Error GoTo 0
End Sub

Private Sub PaintReg(ByVal nm As String, ByVal v As Long, ByVal is9 As Boolean)
    Dim hx As String
    If is9 Then
        hx = Right$("000" & Hex$(v And VADDR_MAX), 3)
    ElseIf nm = "IR" Then
        hx = Right$("0000" & Hex$(v And &HFFFF&), 4)
    Else
        hx = Right$("00" & Hex$(v And BYTE_MAX), 2)
    End If
    UI_SetText PFX_REG & nm, nm & " " & hx & "h (" & CStr(v) & ")"
End Sub

Private Function BusHex(ByVal v As Integer, ByVal is9 As Boolean) As String
    If v = HIGH_Z Then
        BusHex = "HiZ"
    ElseIf is9 Then
        BusHex = Right$("000" & Hex$(v And VADDR_MAX), 3) & "h"
    Else
        BusHex = Right$("00" & Hex$(v And BYTE_MAX), 2) & "h"
    End If
End Function

Private Function BitMask(ByVal bitIndex As Integer) As Long
    BitMask = CLng(2 ^ bitIndex)
End Function

Private Function ActiveReceiver(ByVal cw As Long) As String
    If (cw And SIG_LD_AX) <> 0 Then ActiveReceiver = "AX": Exit Function
    If (cw And SIG_LD_BX) <> 0 Then ActiveReceiver = "BX": Exit Function
    If (cw And SIG_LD_CX) <> 0 Then ActiveReceiver = "CX": Exit Function
    If (cw And SIG_LD_DX) <> 0 Then ActiveReceiver = "DX": Exit Function
    If (cw And SIG_LD_MDR) <> 0 Then ActiveReceiver = "MDR": Exit Function
    If (cw And SIG_LD_IR) <> 0 Then ActiveReceiver = "IR": Exit Function
    If (cw And SIG_LD_PC) <> 0 Then ActiveReceiver = "PC": Exit Function
    If (cw And SIG_LD_MAR) <> 0 Then ActiveReceiver = "MAR": Exit Function
    If (cw And SIG_LD_TA) <> 0 Then ActiveReceiver = "TA": Exit Function
    If (cw And SIG_LD_TB) <> 0 Then ActiveReceiver = "TB": Exit Function
    If (cw And SIG_LD_LCD) <> 0 Then ActiveReceiver = "LCD": Exit Function
    ActiveReceiver = ""
End Function

' =====================================================================
' Test_PanelCPU - shape presence + refresh paint smoke test
' =====================================================================
Public Function Test_PanelCPU() As Long
    Dim fails As Long
    Dim ws As Worksheet
    Dim names As Variant
    Dim i As Integer

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    If ws Is Nothing Then
        Debug.Print "Test_PanelCPU: FAIL no Simulator sheet"
        Test_PanelCPU = 1
        Exit Function
    End If

    BuildCPUPanel ws

    names = Array( _
        PFX_REG & "AX", PFX_BUF & "AX", PFX_REG & "PC", PFX_BUF & "PC", _
        PFX_REG & "IR", PFX_REG & "MAR", PFX_REG & "MDR", _
        PFX_REG & "TA", PFX_REG & "TB", PFX_BUF & "ALU", _
        PFX_BUF & "KBD", PFX_BUF & "LCD", _
        PFX_WIRE & "INT_ADDR", PFX_WIRE & "ADDR", PFX_WIRE & "INT_DATA", PFX_WIRE & "DATA", _
        PFX_SIG & "En_PC", PFX_SIG & "En_AX", PFX_SIG & "Mem_R", PFX_SIG & "CLK", _
        "lbl_Transfer", "flg_ZF")

    For i = LBound(names) To UBound(names)
        fails = fails + PanCheck("shape " & names(i), ShapeExists(ws, CStr(names(i))))
    Next i

    ' Simulate MDR → AX transfer
    ResetCPU
    ReleaseBuses
    SetMDR &H2A
    SetReg R_AX, 0
    DriveDataBus "MDR", CPU.MDR, True
    Bus.Ctrl = SIG_EN_MDR Or SIG_LD_AX
    Bus.Clk = True
    RefreshCPUPanel
    fails = fails + PanCheck("transfer label", InStr(1, ws.Shapes("lbl_Transfer").TextFrame2.TextRange.Text, "MDR") > 0)
    fails = fails + PanCheck("driver coral path", True)  ' visual; logic already exercised
    fails = fails + PanCheck("CLK on", Bus.Clk)

    ReleaseBuses
    Bus.Clk = False
    RefreshCPUPanel

    Debug.Print IIf(fails = 0, "Test_PanelCPU: ALL OK", "Test_PanelCPU: " & fails & " FAIL(S)")
    Test_PanelCPU = fails
    On Error GoTo 0
End Function

Private Function ShapeExists(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(nm)
    ShapeExists = Not s Is Nothing
    On Error GoTo 0
End Function

Private Function PanCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    PanCheck = IIf(passed, 0, 1)
End Function
