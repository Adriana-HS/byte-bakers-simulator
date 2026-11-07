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
    Publish cw
    CPU.R(R_AX) = v
    DriveDataBus "AX", v, False
    UI_AnimateTransfer PFX_REG & "AX", PFX_BUF & "LCD", PFX_WIRE & "DATA", IoGlyph(v)
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

