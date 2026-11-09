Attribute VB_Name = "modIO"
' =====================================================================
' modIO - IN / OUT through the tri-state buffers (G4 excellent, G1)
' Owner: Misael  |  Task: #30 (M7)
'
' IN  r,[KBD]  enables the keyboard buffer and latches r.
' OUT [LCD],AX enables AX and latches the LCD.
' Memory read and write stay at 0. The value travels on the data bus.
' =====================================================================
Option Explicit

Private mLastCW As Long

' Pops the keyboard into destReg (0..3 = AX..DX). Empty queue latches 0.
Public Function IoRead(ByVal destReg As Integer) As Integer
    Dim cw As Long
    Dim v As Integer
    Dim shown As String
    If destReg < R_AX Or destReg > R_DX Then destReg = R_AX
    cw = SIG_IO_R Or SIG_EN_KBD Or LdSignal(destReg)
    Publish cw
    If (RAM(KBD_COUNT) And &HFF&) <= 0 Then
        KbdNoteWaiting
        v = 0
    Else
        v = KbdPop()
    End If
    DriveDataBus "KBD", v, False
    CPU.R(destReg) = v
    shown = IoGlyph(v)
    UI_AnimateTransfer PFX_BUF & "KBD", PFX_REG & RegName(destReg), PFX_WIRE & "DATA", shown
    RefreshCPUPanel
    RefreshKeyboard
    IoRead = v
End Function

' OUT [LCD],AX. v is the byte AX drives onto the bus.
Public Sub IoWrite(ByVal v As Integer)
    Dim cw As Long
    v = v And &HFF&
    cw = SIG_EN_AX Or SIG_IO_W Or SIG_LD_LCD
    mLastCW = cw
    LcdWrite v
    RefreshCPUPanel
End Sub

' Type F, IN it into AX, OUT it to the LCD.
Public Sub Demo_LetterF()
    KbdInit
    ClearLCD
    KbdPress Asc("F")
    IoRead R_AX
    IoWrite CPU.R(R_AX)
End Sub

Public Function IoLastCW() As Long
    IoLastCW = mLastCW
End Function

Private Sub Publish(ByVal cw As Long)
    Dim msg As String
    msg = ValidateControlWord(cw)
    If Len(msg) > 0 Then
        Debug.Print "IO rejected: " & msg
        Exit Sub
    End If
    ReleaseBuses
    Bus.Ctrl = cw
    mLastCW = cw
End Sub

Private Function LdSignal(ByVal destReg As Integer) As Long
    Select Case destReg
        Case R_BX: LdSignal = SIG_LD_BX
        Case R_CX: LdSignal = SIG_LD_CX
        Case R_DX: LdSignal = SIG_LD_DX
        Case Else: LdSignal = SIG_LD_AX
    End Select
End Function

Private Function RegName(ByVal destReg As Integer) As String
    Select Case destReg
        Case R_BX: RegName = "BX"
        Case R_CX: RegName = "CX"
        Case R_DX: RegName = "DX"
        Case Else: RegName = "AX"
    End Select
End Function

Private Function IoGlyph(ByVal v As Integer) As String
    If v >= 32 And v < 127 Then
        IoGlyph = Chr$(v)
    Else
        IoGlyph = "0x" & Right$("00" & Hex$(v), 2)
    End If
End Function

Public Function Test_IO() As Long
    Dim fails As Long
    Dim v As Integer
    Dim saved As Long
    Dim ws As Worksheet
    Dim s As Shape
    Dim cwIn As Long
    Dim cwOut As Long

    saved = AnimDelay
    AnimDelay = 0

    KbdInit
    ClearLCD
    KbdPress Asc("F")
    v = IoRead(R_AX)
    cwIn = IoLastCW()
    fails = fails + IoCheck("IN returns F", v = Asc("F") And CPU.R(R_AX) = Asc("F"))
    fails = fails + IoCheck("IN signals", _
        (cwIn And (SIG_IO_R Or SIG_EN_KBD Or SIG_LD_AX)) = (SIG_IO_R Or SIG_EN_KBD Or SIG_LD_AX))
    fails = fails + IoCheck("IN memory stays 0", (cwIn And (SIG_MEM_R Or SIG_MEM_W)) = 0)
    fails = fails + IoCheck("IN word is legal", Len(ValidateControlWord(cwIn)) = 0)
    fails = fails + IoCheck("queue consumed", RAM(KBD_COUNT) = 0)

    IoWrite CPU.R(R_AX)
    cwOut = IoLastCW()
    fails = fails + IoCheck("OUT latched F", RAM(LCD_BASE) = Asc("F"))
    fails = fails + IoCheck("OUT signals", _
        (cwOut And (SIG_EN_AX Or SIG_IO_W Or SIG_LD_LCD)) = (SIG_EN_AX Or SIG_IO_W Or SIG_LD_LCD))
    fails = fails + IoCheck("OUT memory stays 0", (cwOut And (SIG_MEM_R Or SIG_MEM_W)) = 0)
    fails = fails + IoCheck("OUT word is legal", Len(ValidateControlWord(cwOut)) = 0)

    KbdInit
    v = IoRead(R_AX)
    fails = fails + IoCheck("empty IN is 0", v = 0)
    fails = fails + IoCheck("empty IN waits", KbdStatusText = "waiting for key")

    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    LcdBuild ws
    Demo_LetterF
    fails = fails + IoCheck("demo F in AX and LCD", CPU.R(R_AX) = Asc("F") And RAM(LCD_BASE) = Asc("F"))
    fails = fails + IoCheck("demo queue empty", RAM(KBD_COUNT) = 0)
    On Error Resume Next
    Set s = ws.Shapes("lcd_c0")
    On Error GoTo 0
    If Not s Is Nothing Then
        fails = fails + IoCheck("demo shows F", s.TextFrame2.TextRange.Text = "F")
    End If

    AnimDelay = saved
    Debug.Print IIf(fails = 0, "Test_IO: ALL OK", "Test_IO: " & fails & " FAIL(S)")
    Test_IO = fails
End Function

Private Function IoCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    IoCheck = IIf(passed, 0, 1)
End Function
