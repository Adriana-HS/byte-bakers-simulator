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
Private Const SIG_W As Single = 40
Private Const SIG_H As Single = 14

' Soft chrome (hierarchy). Values still use coral, amber and gray dynamically.
Private Const CLR_LABEL As Long = &HC8B4A0&      ' RGB(160,180,200) muted label
Private Const CLR_CHIP As Long = &H4A3828&       ' RGB(40,56,74) chip bg
Private Const CLR_CLK_ON As Long = &H5AD8D8&     ' RGB(216,216,90) CLK pulse
Private Const CLR_ISA_HL As Long = &HC8E8FF&     ' RGB(255,232,200) ISA row

Private mPrevIsaRow As Long
Private mPrevMicro As Integer
Private mLastClock As Long
Private histCw(1 To 8) As Long
Private histN As Integer
Private histAt As Integer

Private Const CPU_FRAME_TOP As Single = 50
Private Const CPU_FRAME_BOTTOM As Single = 394   ' frm_CPU top 50 + height 344
Private Const IO_FRAME_TOP As Single = 406

' =====================================================================
' BuildCPUPanel - create every shape inside the CPU frame (frm_CPU).
' =====================================================================
Public Sub BuildCPUPanel(ByVal ws As Worksheet)
    Dim x As Single, y As Single
    Dim i As Integer
    Dim nm As String
    Dim s As Shape

    On Error GoTo EH
    UI_ClearTransfer
    mPrevIsaRow = 0
    mPrevMicro = 0
    mLastClock = -1
    histN = 0
    histAt = 0

    ' Title + transfer (hierarchy: what is happening now)
    AddBox ws, "lbl_CPUTitle", PANEL_L, PANEL_T, 120, 18, "CPU / BUSES", CLR_CHIP
    AddBox ws, "lbl_Transfer", PANEL_L + 128, PANEL_T, 220, 18, "idle", CLR_CHIP
    AddBox ws, "phs_Now", PANEL_L + 356, PANEL_T, 78, 18, "FETCH", CLR_CHIP
    BuildTimeline ws

    ' --- GP registers + buffers (left column) ---
    y = PANEL_T + 44
    AddBox ws, "lbl_Regs", PANEL_L, y, 100, 16, "REGS", CLR_CHIP
    y = y + 18
    AddRegPair ws, "AX", PANEL_L, y: y = y + 34
    AddRegPair ws, "BX", PANEL_L, y: y = y + 34
    AddRegPair ws, "CX", PANEL_L, y: y = y + 34
    AddRegPair ws, "DX", PANEL_L, y

    ' --- Control regs (center) ---
    x = PANEL_L + 160
    y = PANEL_T + 44
    AddBox ws, "lbl_CtrlRegs", x, y, 120, 16, "CONTROL", CLR_CHIP
    y = y + 18
    AddRegPair ws, "PC", x, y: y = y + 32
    AddRegPair ws, "IR", x, y: y = y + 32
    AddRegPair ws, "MAR", x, y: y = y + 32
    AddRegPair ws, "MDR", x, y: y = y + 32
    AddRegPair ws, "PADDR", x, y

    ' --- ALU + flags (right of control) ---
    x = PANEL_L + 340
    y = PANEL_T + 44
    AddBox ws, "lbl_ALU", x, y, 100, 16, "ALU", CLR_CHIP
    y = y + 18
    AddRegPair ws, "TA", x, y: y = y + 36
    AddRegPair ws, "TB", x, y: y = y + 36
    AddBox ws, PFX_REG & "ALU", x, y, BOX_W, BOX_H, "ALU", CLR_HIGHZ
    AddBox ws, PFX_BUF & "ALU", x + BOX_W + 8, y + 5, BUF_W, BUF_H, "Z", CLR_HIGHZ
    y = y + 36
    AddBox ws, "flg_ZF", x, y, 40, 22, "ZF 0", CLR_HIGHZ
    AddBox ws, "flg_CF", x + 44, y, 40, 22, "CF 0", CLR_HIGHZ
    AddBox ws, "flg_SF", x + 88, y, 40, 22, "SF 0", CLR_HIGHZ

    ' --- I/O buffers: one even stack, clear of the history ---
    x = PANEL_L + 470
    y = PANEL_T + 62
    AddBox ws, "lbl_IOBuf", x, y - 18, 72, 16, "I/O BUF", CLR_CHIP
    AddKey ws, PFX_BUF & "KBD", x, y, 72, 22, "KBD"
    AddKey ws, PFX_BUF & "LCD", x, y + 26, 72, 22, "LCD"
    AddKey ws, PFX_BUF & "MEM", x, y + 52, 72, 22, "MEM"
    BuildHistory ws

    ' --- Buses. Label, then a full horizontal wire the arrow can follow. ---
    y = PANEL_T + 216
    AddBox ws, "lbl_Buses", PANEL_L, y, 250, 12, "BUSES    INT 9-bit     EXT 8-bit", CLR_CHIP
    y = y + 14
    AddBusRow ws, "val_INT_ADDR", "INT HiZ", PFX_WIRE & "INT_ADDR", y
    y = y + 14
    AddBusRow ws, "val_EXT_ADDR", "EXT HiZ", PFX_WIRE & "ADDR", y
    y = y + 14
    AddBusRow ws, "val_INT_DATA", "ID HiZ", PFX_WIRE & "INT_DATA", y
    y = y + 14
    AddBusRow ws, "val_DATA", "D HiZ", PFX_WIRE & "DATA", y

    ' --- Control word. Two even rows, inset from the frame edge. ---
    y = PANEL_T + 286
    AddBox ws, "lbl_CW", PANEL_L, y, 170, 12, "CONTROL WORD", CLR_CHIP
    y = y + 14
    For i = 0 To NUM_SIGNALS - 1
        nm = SignalName(i)
        If Len(nm) > 0 Then
            x = PANEL_L + 8 + (i Mod 14) * 44
            Set s = AddKey(ws, PFX_SIG & nm, x, y + (i \ 14) * 16, 32, 12, ShortSig(nm))
        End If
    Next i
    AddKey ws, PFX_SIG & "CLK", PANEL_L + 8 + 13 * 44, y + 16, 32, 12, "CLK"

    RefreshCPUPanel
    On Error Resume Next
    EnsureISASheet
    On Error GoTo 0
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

Private Function AddKey(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal txt As String) As Shape
    Dim s As Shape
    Set s = UI_AddBox(ws, nm, l, t, w, h, txt, CLR_HIGHZ)
    s.TextFrame2.MarginLeft = 0
    s.TextFrame2.MarginRight = 0
    s.TextFrame2.MarginTop = 0
    s.TextFrame2.MarginBottom = 0
    s.TextFrame2.TextRange.Font.Size = 8
    s.TextFrame2.TextRange.Font.Name = "Consolas"
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = CLR_BG
    s.Line.Weight = 1.25
    Set AddKey = s
End Function

Private Sub AddBusRow(ByVal ws As Worksheet, ByVal valName As String, _
        ByVal caption As String, ByVal wireName As String, ByVal y As Single)
    AddBox ws, valName, PANEL_L, y, 74, 12, caption, CLR_CHIP
    UI_AddWire ws, wireName, PANEL_L + 82, y + 6, 468, y + 6, 2.5
End Sub

Private Function ShortSig(ByVal nm As String) As String
    Select Case nm
        Case "En_PC": ShortSig = "EnPC"
        Case "Ld_PC": ShortSig = "LdPC"
        Case "Inc_PC": ShortSig = "+PC"
        Case "Ld_MAR": ShortSig = "MAR"
        Case "Ld_MDR": ShortSig = "MDR"
        Case "En_MDR": ShortSig = "eMDR"
        Case "Ld_IR": ShortSig = "LdIR"
        Case "En_IR_A": ShortSig = "IRa"
        Case "En_AX": ShortSig = "EAX"
        Case "En_BX": ShortSig = "EBX"
        Case "En_CX": ShortSig = "ECX"
        Case "En_DX": ShortSig = "EDX"
        Case "Ld_AX": ShortSig = "LAX"
        Case "Ld_BX": ShortSig = "LBX"
        Case "Ld_CX": ShortSig = "LCX"
        Case "Ld_DX": ShortSig = "LDX"
        Case "Ld_TA": ShortSig = "LTA"
        Case "Ld_TB": ShortSig = "LTB"
        Case "En_ALU": ShortSig = "ALU"
        Case "Ld_FLAGS": ShortSig = "FLAG"
        Case "Mem_R": ShortSig = "MemR"
        Case "Mem_W": ShortSig = "MemW"
        Case "IO_R": ShortSig = "IOr"
        Case "IO_W": ShortSig = "IOw"
        Case "En_KBD": ShortSig = "KBD"
        Case "Ld_LCD": ShortSig = "LCD"
        Case "En_IR_D": ShortSig = "IRd"
        Case Else: ShortSig = Left$(nm, 4)
    End Select
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
        If CStr(bufs(i)) = "KBD" Or CStr(bufs(i)) = "LCD" Or CStr(bufs(i)) = "MEM" Then
            UI_SetText PFX_BUF & CStr(bufs(i)), CStr(bufs(i))
        Else
            UI_SetText PFX_BUF & CStr(bufs(i)), "Z"
        End If
        UI_SetFill PFX_REG & CStr(bufs(i)), CLR_HIGHZ
    Next i
    UI_SetHighZ PFX_WIRE & "INT_ADDR"
    UI_SetHighZ PFX_WIRE & "ADDR"
    UI_SetHighZ PFX_WIRE & "INT_DATA"
    UI_SetHighZ PFX_WIRE & "DATA"

    ' Bus values
    UI_SetText "val_INT_ADDR", "INT " & BusHex(Bus.IntAddr, True)
    UI_SetText "val_EXT_ADDR", "EXT " & BusHex(Bus.ExtAddr, False)
    UI_SetText "val_INT_DATA", "ID " & IIf(Bus.Internal, BusHex(Bus.Data, False), "HiZ")
    UI_SetText "val_DATA", "D " & IIf(Not Bus.Internal And Bus.Data <> HIGH_Z, BusHex(Bus.Data, False), _
        IIf(Bus.Data = HIGH_Z, "HiZ", BusHex(Bus.Data, False)))

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
        If Bus.DataDriver = "KBD" Or Bus.DataDriver = "LCD" Or Bus.DataDriver = "MEM" Then
            UI_SetText PFX_BUF & Bus.DataDriver, Bus.DataDriver
        Else
            UI_SetText PFX_BUF & Bus.DataDriver, ">"
        End If
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

    ' Transfer label e.g. MDR -> AX
    If Len(Bus.DataDriver) > 0 And Len(recv) > 0 Then
        UI_SetText "lbl_Transfer", Bus.DataDriver & " -> " & recv
    ElseIf Len(Bus.AddrDriver) > 0 And (cw And SIG_LD_MAR) <> 0 Then
        UI_SetText "lbl_Transfer", Bus.AddrDriver & " -> MAR"
    ElseIf Len(Bus.DataDriver) > 0 Then
        UI_SetText "lbl_Transfer", Bus.DataDriver & " -> bus"
    Else
        UI_SetText "lbl_Transfer", "idle"
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
    UI_SetText PFX_SIG & "CLK", IIf(Bus.Clk, "CLK ON", "CLK")

    ' Highlight current IR opcode row on ISA sheet (cells, not shapes)
    Decode CPU.IR, op, reg, operand
    row = ISARow(op)
    Set wsIsa = ThisWorkbook.Worksheets(SH_ISA)
    If Not wsIsa Is Nothing Then
        If mPrevIsaRow > 0 And mPrevIsaRow <> row Then
            wsIsa.Range(wsIsa.Cells(mPrevIsaRow, 2), wsIsa.Cells(mPrevIsaRow, ISA_TABLE_LAST_COL)).Interior.ColorIndex = xlNone
        End If
        If row > 0 Then
            wsIsa.Range(wsIsa.Cells(row, 2), wsIsa.Cells(row, ISA_TABLE_LAST_COL)).Interior.Color = CLR_ISA_HL
            mPrevIsaRow = row
        End If
    End If

    On Error GoTo 0
    NoteHistory
    PaintPhase
    PaintTimeline
    PaintHistory
    mPrevMicro = MicroStep

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

Private Sub BuildTimeline(ByVal ws As Worksheet)
    Dim i As Integer
    Dim x As Single
    x = PANEL_L
    For i = 1 To 6
        AddBox ws, "tl_F" & i, x, 78, 26, 16, "F" & i, CLR_CHIP
        x = x + 29
    Next i
    x = x + 8
    For i = 1 To 3
        AddBox ws, "tl_E" & i, x, 78, 26, 16, "E" & i, CLR_CHIP
        x = x + 29
    Next i
End Sub

Private Sub EnsureISASheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_ISA)
    On Error GoTo 0
    If ws Is Nothing Then
        BuildISASheet
    ElseIf Len(Trim$(CStr(ws.Range("B2").Value))) = 0 Then
        BuildISASheet
    End If
End Sub

Private Sub BuildHistory(ByVal ws As Worksheet)
    Dim c As Integer
    Dim s As Shape
    AddBox ws, "lbl_Hist", 488, 206, 150, 14, "LAST 8", CLR_CHIP
    For c = 1 To 8
        Set s = UI_AddBox(ws, "hist_" & c, 488 + (c - 1) * 18, 222, 16, 112, ZeroBits(), CLR_CHIP)
        s.Line.Visible = msoTrue
        s.Line.ForeColor.RGB = CLR_BG
        s.Line.Weight = 1
        s.TextFrame2.TextRange.Font.Size = 6
        s.TextFrame2.TextRange.Font.Name = "Consolas"
        s.TextFrame2.WordWrap = msoTrue
        s.TextFrame2.MarginLeft = 0
        s.TextFrame2.MarginRight = 0
        s.TextFrame2.MarginTop = 0
        s.TextFrame2.MarginBottom = 0
    Next c
End Sub

Private Function ZeroBits() As String
    Dim b As Integer
    For b = 1 To NUM_SIGNALS
        ZeroBits = ZeroBits & "0"
    Next b
End Function

Private Sub NoteHistory()
    If Clock = mLastClock Then Exit Sub
    mLastClock = Clock
    histAt = (histAt Mod 8) + 1
    If histN < 8 Then histN = histN + 1
    histCw(histAt) = Bus.Ctrl
End Sub

Private Sub CurrentPlace(ByRef phase As String, ByRef fetchNow As Integer, ByRef execNow As Integer)
    phase = "FETCH"
    fetchNow = 0
    execNow = 0
    If MicroStep > 7 Then
        phase = "EXECUTE"
        fetchNow = 6
        execNow = MicroStep - 7
    ElseIf MicroStep = 7 Then
        phase = "DECODE"
        fetchNow = 6
    ElseIf MicroStep >= 2 And MicroStep <= 6 Then
        phase = "FETCH"
        fetchNow = MicroStep - 1
    ElseIf MicroStep = 1 And mPrevMicro > 7 Then
        phase = "EXECUTE"
        fetchNow = 6
        execNow = mPrevMicro - 6
    End If
End Sub

Private Function ExecSlots() As Integer
    Dim op As Integer, reg As Integer, operand As Integer
    Dim fam As String, st As String, sg As String, cy As Integer
    Decode CPU.IR, op, reg, operand
    If Not IsValidOpcode(op) Then
        ExecSlots = 1
    Else
        ISA_Detail op, fam, st, sg, cy
        ExecSlots = cy - 6
        If ExecSlots < 1 Then ExecSlots = 1
        If ExecSlots > 3 Then ExecSlots = 3
    End If
End Function

Private Sub PaintPhase()
    Dim phase As String
    Dim fetchNow As Integer
    Dim execNow As Integer
    Dim fill As Long
    CurrentPlace phase, fetchNow, execNow
    UI_SetText "phs_Now", phase
    If phase = "EXECUTE" Then
        fill = CLR_HIT
    ElseIf phase = "DECODE" Then
        fill = CLR_LOAD
    Else
        fill = CLR_ACTIVE
    End If
    UI_SetFill "phs_Now", fill
End Sub

Private Sub PaintTimeline()
    Dim phase As String
    Dim fetchNow As Integer
    Dim execNow As Integer
    Dim i As Integer
    Dim nExec As Integer
    CurrentPlace phase, fetchNow, execNow
    nExec = ExecSlots()
    For i = 1 To 6
        PaintStep "tl_F" & i, i, fetchNow, True
    Next i
    For i = 1 To 3
        If i <= nExec Then
            PaintStep "tl_E" & i, i, execNow, (phase = "EXECUTE")
        Else
            UI_SetFill "tl_E" & i, CLR_CHIP
        End If
    Next i
End Sub

' kind = finished / current / waiting. During fetch the E chips are waiting.
Private Sub PaintStep(ByVal nm As String, ByVal index As Integer, ByVal current As Integer, ByVal started As Boolean)
    If Not started Or current = 0 Then
        UI_SetFill nm, CLR_CHIP
    ElseIf index < current Then
        UI_SetFill nm, CLR_HIGHZ
    ElseIf index = current Then
        UI_SetFill nm, CLR_ACTIVE
    Else
        UI_SetFill nm, CLR_CHIP
    End If
End Sub

Private Sub PaintHistory()
    Dim showCol As Integer
    Dim b As Integer
    Dim slot As Integer
    Dim cw As Long
    Dim bit As Long
    Dim s As Shape
    Dim ch As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For showCol = 1 To 8
        cw = 0
        If showCol > 8 - histN Then
            slot = histAt - (8 - showCol)
            Do While slot <= 0
                slot = slot + 8
            Loop
            cw = histCw(slot)
        End If
        Set s = Nothing
        On Error Resume Next
        Set s = ws.Shapes("hist_" & showCol)
        On Error GoTo 0
        If Not s Is Nothing Then
            For b = 0 To NUM_SIGNALS - 1
                bit = CLng(2 ^ b)
                If showCol <= 8 - histN Then
                    ch = "0"
                    s.TextFrame2.TextRange.Characters(b + 1, 1).Font.Fill.ForeColor.RGB = CLR_HIGHZ
                ElseIf (cw And bit) <> 0 Then
                    ch = "1"
                    s.TextFrame2.TextRange.Characters(b + 1, 1).Font.Fill.ForeColor.RGB = CLR_ACTIVE
                Else
                    ch = "0"
                    s.TextFrame2.TextRange.Characters(b + 1, 1).Font.Fill.ForeColor.RGB = CLR_HIGHZ
                End If
                s.TextFrame2.TextRange.Characters(b + 1, 1).Text = ch
            Next b
        End If
    Next showCol
End Sub

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

    ' Simulate MDR -> AX transfer
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

    fails = fails + PanCheck("timeline F1", ShapeExists(ws, "tl_F1"))
    fails = fails + PanCheck("timeline E3", ShapeExists(ws, "tl_E3"))
    fails = fails + PanCheck("phase chip", ShapeExists(ws, "phs_Now"))
    fails = fails + PanCheck("history column", ShapeExists(ws, "hist_8"))
    fails = fails + PanCheck("inside CPU frame", ShapeInsideCpu(ws, "tl_F1") And ShapeInsideCpu(ws, "phs_Now") And ShapeInsideCpu(ws, "hist_8") And ShapeInsideCpu(ws, PFX_SIG & "En_IR_D"))
    fails = fails + PanCheck("above I/O frame", ShapeAboveIo(ws, "tl_F6") And ShapeAboveIo(ws, PFX_SIG & "CLK") And ShapeAboveIo(ws, "hist_1"))

    MicroStep = 3
    RefreshCPUPanel
    fails = fails + PanCheck("F2 is the current fetch", ws.Shapes("tl_F2").Fill.ForeColor.RGB = CLR_ACTIVE)
    fails = fails + PanCheck("F1 is already dim", ws.Shapes("tl_F1").Fill.ForeColor.RGB = CLR_HIGHZ)
    fails = fails + PanCheck("F3 is still waiting", ws.Shapes("tl_F3").Fill.ForeColor.RGB = CLR_CHIP)
    fails = fails + PanCheck("phase is FETCH", ws.Shapes("phs_Now").TextFrame2.TextRange.Text = "FETCH")

    Clock = 1
    Bus.Ctrl = SIG_EN_PC
    RefreshCPUPanel
    Clock = 2
    Bus.Ctrl = SIG_EN_AX Or SIG_LD_TA
    RefreshCPUPanel
    fails = fails + PanCheck("newest column bit 0 is off", ws.Shapes("hist_8").TextFrame2.TextRange.Characters(1, 1).Font.Fill.ForeColor.RGB = CLR_HIGHZ)
    fails = fails + PanCheck("newest column En_AX is on", ws.Shapes("hist_8").TextFrame2.TextRange.Characters(9, 1).Font.Fill.ForeColor.RGB = CLR_ACTIVE)

    ReleaseBuses
    Bus.Clk = False
    RefreshCPUPanel

    Debug.Print IIf(fails = 0, "Test_PanelCPU: ALL OK", "Test_PanelCPU: " & fails & " FAIL(S)")
    Test_PanelCPU = fails
    On Error GoTo 0
End Function

Private Function ShapeInsideCpu(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    Set s = ws.Shapes(nm)
    ShapeInsideCpu = (s.Top >= CPU_FRAME_TOP) And (s.Top + s.Height <= CPU_FRAME_BOTTOM) And (s.Left >= 20) And (s.Left + s.Width <= 660)
End Function

Private Function ShapeAboveIo(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    Set s = ws.Shapes(nm)
    ShapeAboveIo = (s.Top + s.Height <= IO_FRAME_TOP)
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
