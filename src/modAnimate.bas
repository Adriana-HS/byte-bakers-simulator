Attribute VB_Name = "modAnimate"
' =====================================================================
' modAnimate - one arrow per control word (#68 / A12)
' Owner: Adriana
'
' En_* picks the driver shape, Ld_* the receiver. Memory and I/O use
' the external data bus. PC and IR address moves use the internal
' address bus. Everything else uses the internal data bus.
' AnimDelay 0 skips the token inside UI_MoveTokenViaBus, so only the
' arrow remains. LogMicroStep (#77) belongs to modLog and is not here.
' =====================================================================
Option Explicit

Public Sub AnimateControlWord(ByVal cw As Long, ByVal text As String, ByVal value As String)
    Dim drv As String
    Dim recv As String
    Dim recv2 As String
    Dim bus As String
    Dim s1 As String
    Dim s2 As String
    Dim n1 As String
    Dim n2 As String
    Dim shown As String

    shown = value
    If Len(shown) = 0 Then shown = text

    If CollisionOf(cw, True, s1, s2, n1, n2) Then
        UI_ShowCollision s1, s2, PFX_WIRE & "DATA"
        UI_Badge "en", s1, n1, CLR_MISS
        UI_Badge "en2", s2, n2, CLR_MISS
        Exit Sub
    End If
    If CollisionOf(cw, False, s1, s2, n1, n2) Then
        UI_ShowCollision s1, s2, PFX_WIRE & "INT_ADDR"
        UI_Badge "en", s1, n1, CLR_MISS
        UI_Badge "en2", s2, n2, CLR_MISS
        Exit Sub
    End If

    If IoLeavesMemHighZ(cw) Then
        UI_SetHighZ PFX_BUF & "MEM"
        UI_SetText PFX_BUF & "MEM", "Z"
    End If

    drv = DriverShape(cw)
    recv = ReceiverShape(cw)
    recv2 = SecondReceiver(cw)
    bus = BusWireOf(cw)
    If Len(drv) = 0 Or Len(recv) = 0 Then Exit Sub

    UI_AnimateTransfer drv, recv, bus, shown, recv2
    If (cw And SIG_LD_FLAGS) <> 0 Then
        UI_Highlight "flg_CF", CLR_LOAD
        UI_Highlight "flg_SF", CLR_LOAD
    End If
End Sub

Public Function DriverShape(ByVal cw As Long) As String
    Dim nm As String
    nm = NthSignal(cw, False, 1)
    If Len(nm) > 0 Then
        DriverShape = ShapeFor(nm)
    ElseIf (cw And SIG_MEM_R) <> 0 Then
        DriverShape = PFX_BUF & "MEM"
    End If
End Function

Public Function ReceiverShape(ByVal cw As Long) As String
    Dim nm As String
    nm = NthSignal(cw, True, 1)
    If Len(nm) > 0 Then
        ReceiverShape = ShapeFor(nm)
    ElseIf (cw And SIG_MEM_W) <> 0 Then
        ReceiverShape = PFX_BUF & "MEM"
    End If
End Function

Public Function SecondReceiver(ByVal cw As Long) As String
    Dim nm As String
    nm = NthSignal(cw, True, 2)
    If Len(nm) > 0 Then SecondReceiver = ShapeFor(nm)
End Function

Public Function BusWireOf(ByVal cw As Long) As String
    If (cw And (SIG_MEM_R Or SIG_MEM_W Or SIG_IO_R Or SIG_IO_W Or SIG_EN_KBD Or SIG_LD_LCD)) <> 0 Then
        BusWireOf = PFX_WIRE & "DATA"
    ElseIf (cw And (SIG_EN_PC Or SIG_EN_IR_A)) <> 0 And (cw And (SIG_LD_MAR Or SIG_LD_PC)) <> 0 Then
        BusWireOf = PFX_WIRE & "INT_ADDR"
    ElseIf Len(DriverShape(cw)) > 0 Then
        BusWireOf = PFX_WIRE & "INT_DATA"
    End If
End Function

' IN and OUT keep the memory buffer in High-Z. Mem_R and Mem_W stay 0.
Public Function IoLeavesMemHighZ(ByVal cw As Long) As Boolean
    IoLeavesMemHighZ = ((cw And (SIG_IO_R Or SIG_IO_W)) <> 0) And ((cw And (SIG_MEM_R Or SIG_MEM_W)) = 0)
End Function

Public Function Test_Animate() As Long
    Dim fails As Long
    Dim saved As Long
    Dim s1 As String
    Dim s2 As String
    Dim n1 As String
    Dim n2 As String
    Dim hit As Boolean

    fails = fails + AniCheck("PC -> MAR", _
        DriverShape(SIG_EN_PC Or SIG_LD_MAR) = PFX_REG & "PC" And _
        ReceiverShape(SIG_EN_PC Or SIG_LD_MAR) = PFX_REG & "MAR" And _
        BusWireOf(SIG_EN_PC Or SIG_LD_MAR) = PFX_WIRE & "INT_ADDR")
    fails = fails + AniCheck("fetch read", _
        DriverShape(SIG_MEM_R Or SIG_LD_MDR) = PFX_BUF & "MEM" And _
        ReceiverShape(SIG_MEM_R Or SIG_LD_MDR) = PFX_REG & "MDR" And _
        BusWireOf(SIG_MEM_R Or SIG_LD_MDR) = PFX_WIRE & "DATA")
    fails = fails + AniCheck("MDR -> IR", _
        DriverShape(SIG_EN_MDR Or SIG_LD_IR Or SIG_INC_PC) = PFX_REG & "MDR" And _
        ReceiverShape(SIG_EN_MDR Or SIG_LD_IR Or SIG_INC_PC) = PFX_REG & "IR" And _
        BusWireOf(SIG_EN_MDR Or SIG_LD_IR Or SIG_INC_PC) = PFX_WIRE & "INT_DATA")
    fails = fails + AniCheck("LOAD addr", _
        DriverShape(SIG_EN_IR_A Or SIG_LD_MAR) = PFX_REG & "IR" And _
        ReceiverShape(SIG_EN_IR_A Or SIG_LD_MAR) = PFX_REG & "MAR" And _
        BusWireOf(SIG_EN_IR_A Or SIG_LD_MAR) = PFX_WIRE & "INT_ADDR")
    fails = fails + AniCheck("MDR -> reg", _
        DriverShape(SIG_EN_MDR Or SIG_LD_AX) = PFX_REG & "MDR" And _
        ReceiverShape(SIG_EN_MDR Or SIG_LD_AX) = PFX_REG & "AX" And _
        BusWireOf(SIG_EN_MDR Or SIG_LD_AX) = PFX_WIRE & "INT_DATA")
    fails = fails + AniCheck("STORE reg -> MDR", _
        DriverShape(SIG_EN_AX Or SIG_LD_MDR) = PFX_REG & "AX" And _
        ReceiverShape(SIG_EN_AX Or SIG_LD_MDR) = PFX_REG & "MDR" And _
        BusWireOf(SIG_EN_AX Or SIG_LD_MDR) = PFX_WIRE & "INT_DATA")
    fails = fails + AniCheck("STORE write", _
        DriverShape(SIG_EN_MDR Or SIG_MEM_W) = PFX_REG & "MDR" And _
        ReceiverShape(SIG_EN_MDR Or SIG_MEM_W) = PFX_BUF & "MEM" And _
        BusWireOf(SIG_EN_MDR Or SIG_MEM_W) = PFX_WIRE & "DATA")
    fails = fails + AniCheck("reg -> TA", _
        DriverShape(SIG_EN_BX Or SIG_LD_TA) = PFX_REG & "BX" And _
        ReceiverShape(SIG_EN_BX Or SIG_LD_TA) = PFX_REG & "TA")
    fails = fails + AniCheck("reg -> TB", _
        DriverShape(SIG_EN_CX Or SIG_LD_TB) = PFX_REG & "CX" And _
        ReceiverShape(SIG_EN_CX Or SIG_LD_TB) = PFX_REG & "TB")
    fails = fails + AniCheck("ALU -> AX + flags", _
        DriverShape(SIG_EN_ALU Or SIG_LD_AX Or SIG_LD_FLAGS) = PFX_REG & "ALU" And _
        ReceiverShape(SIG_EN_ALU Or SIG_LD_AX Or SIG_LD_FLAGS) = PFX_REG & "AX" And _
        SecondReceiver(SIG_EN_ALU Or SIG_LD_AX Or SIG_LD_FLAGS) = "flg_ZF" And _
        BusWireOf(SIG_EN_ALU Or SIG_LD_AX Or SIG_LD_FLAGS) = PFX_WIRE & "INT_DATA")
    fails = fails + AniCheck("CMP flags only", _
        DriverShape(SIG_EN_ALU Or SIG_LD_FLAGS) = PFX_REG & "ALU" And _
        ReceiverShape(SIG_EN_ALU Or SIG_LD_FLAGS) = "flg_ZF")
    fails = fails + AniCheck("jump", _
        DriverShape(SIG_EN_IR_A Or SIG_LD_PC) = PFX_REG & "IR" And _
        ReceiverShape(SIG_EN_IR_A Or SIG_LD_PC) = PFX_REG & "PC" And _
        BusWireOf(SIG_EN_IR_A Or SIG_LD_PC) = PFX_WIRE & "INT_ADDR")
    fails = fails + AniCheck("imm -> reg", _
        DriverShape(SIG_EN_IR_D Or SIG_LD_DX) = PFX_REG & "IR" And _
        ReceiverShape(SIG_EN_IR_D Or SIG_LD_DX) = PFX_REG & "DX" And _
        BusWireOf(SIG_EN_IR_D Or SIG_LD_DX) = PFX_WIRE & "INT_DATA")
    fails = fails + AniCheck("IN high-Z", _
        DriverShape(SIG_IO_R Or SIG_EN_KBD Or SIG_LD_AX) = PFX_BUF & "KBD" And _
        ReceiverShape(SIG_IO_R Or SIG_EN_KBD Or SIG_LD_AX) = PFX_REG & "AX" And _
        BusWireOf(SIG_IO_R Or SIG_EN_KBD Or SIG_LD_AX) = PFX_WIRE & "DATA" And _
        IoLeavesMemHighZ(SIG_IO_R Or SIG_EN_KBD Or SIG_LD_AX))
    fails = fails + AniCheck("OUT high-Z", _
        DriverShape(SIG_EN_AX Or SIG_IO_W Or SIG_LD_LCD) = PFX_REG & "AX" And _
        ReceiverShape(SIG_EN_AX Or SIG_IO_W Or SIG_LD_LCD) = PFX_BUF & "LCD" And _
        BusWireOf(SIG_EN_AX Or SIG_IO_W Or SIG_LD_LCD) = PFX_WIRE & "DATA" And _
        IoLeavesMemHighZ(SIG_EN_AX Or SIG_IO_W Or SIG_LD_LCD))
    fails = fails + AniCheck("mem read is not high-Z", Not IoLeavesMemHighZ(SIG_MEM_R Or SIG_LD_MDR))

    hit = CollisionOf(SIG_EN_AX Or SIG_EN_BX, True, s1, s2, n1, n2)
    fails = fails + AniCheck("two data drivers", hit And n1 = "En_AX" And n2 = "En_BX" And s1 = PFX_REG & "AX" And s2 = PFX_REG & "BX")

    saved = AnimDelay
    AnimDelay = 0
    On Error Resume Next
    AnimateControlWord SIG_EN_PC Or SIG_LD_MAR, "PC -> MAR", "PC"
    If Err.Number <> 0 Then
        fails = fails + 1
        Debug.Print "  FAIL  draw PC -> MAR " & Err.Description
        Err.Clear
    End If
    AnimateControlWord SIG_EN_AX Or SIG_EN_BX, "collision", "AX"
    If Err.Number <> 0 Then
        fails = fails + 1
        Debug.Print "  FAIL  draw collision " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    AnimDelay = saved

    Debug.Print IIf(fails = 0, "Test_Animate: ALL OK", "Test_Animate: " & fails & " FAIL(S)")
    Test_Animate = fails
End Function

Private Function CollisionOf(ByVal cw As Long, ByVal dataBus As Boolean, _
        ByRef s1 As String, ByRef s2 As String, ByRef n1 As String, ByRef n2 As String) As Boolean
    Dim i As Integer
    Dim nm As String
    Dim bit As Long
    Dim mask As Long
    Dim n As Integer

    s1 = ""
    s2 = ""
    n1 = ""
    n2 = ""
    If dataBus Then mask = MASK_DATA_DRIVERS Else mask = MASK_ADDR_DRIVERS
    For i = 0 To NUM_SIGNALS - 1
        bit = CLng(2 ^ i)
        If (cw And bit) <> 0 And (mask And bit) <> 0 Then
            nm = SignalName(i)
            n = n + 1
            If n = 1 Then
                n1 = nm
                s1 = ShapeFor(nm)
            ElseIf n = 2 Then
                n2 = nm
                s2 = ShapeFor(nm)
            End If
        End If
    Next i
    CollisionOf = (n >= 2)
End Function

Private Function NthSignal(ByVal cw As Long, ByVal wantLoad As Boolean, ByVal nth As Integer) As String
    Dim i As Integer
    Dim n As Integer
    Dim nm As String
    Dim bit As Long

    For i = 0 To NUM_SIGNALS - 1
        bit = CLng(2 ^ i)
        If (cw And bit) <> 0 Then
            nm = SignalName(i)
            If wantLoad Then
                If Left$(nm, 3) = "Ld_" Then
                    n = n + 1
                    If n = nth Then NthSignal = nm: Exit Function
                End If
            ElseIf Left$(nm, 3) = "En_" Then
                n = n + 1
                If n = nth Then NthSignal = nm: Exit Function
            End If
        End If
    Next i
End Function

Private Function ShapeFor(ByVal nm As String) As String
    Select Case nm
        Case "En_PC": ShapeFor = PFX_REG & "PC"
        Case "Ld_PC": ShapeFor = PFX_REG & "PC"
        Case "Ld_MAR": ShapeFor = PFX_REG & "MAR"
        Case "Ld_MDR": ShapeFor = PFX_REG & "MDR"
        Case "En_MDR": ShapeFor = PFX_REG & "MDR"
        Case "Ld_IR": ShapeFor = PFX_REG & "IR"
        Case "En_IR_A": ShapeFor = PFX_REG & "IR"
        Case "En_IR_D": ShapeFor = PFX_REG & "IR"
        Case "En_AX": ShapeFor = PFX_REG & "AX"
        Case "En_BX": ShapeFor = PFX_REG & "BX"
        Case "En_CX": ShapeFor = PFX_REG & "CX"
        Case "En_DX": ShapeFor = PFX_REG & "DX"
        Case "Ld_AX": ShapeFor = PFX_REG & "AX"
        Case "Ld_BX": ShapeFor = PFX_REG & "BX"
        Case "Ld_CX": ShapeFor = PFX_REG & "CX"
        Case "Ld_DX": ShapeFor = PFX_REG & "DX"
        Case "Ld_TA": ShapeFor = PFX_REG & "TA"
        Case "Ld_TB": ShapeFor = PFX_REG & "TB"
        Case "En_ALU": ShapeFor = PFX_REG & "ALU"
        Case "Ld_FLAGS": ShapeFor = "flg_ZF"
        Case "En_KBD": ShapeFor = PFX_BUF & "KBD"
        Case "Ld_LCD": ShapeFor = PFX_BUF & "LCD"
        Case Else: ShapeFor = ""
    End Select
End Function

Private Function AniCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    AniCheck = IIf(passed, 0, 1)
End Function
