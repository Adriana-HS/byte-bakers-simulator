Attribute VB_Name = "modKeyboard"
' =====================================================================
' modKeyboard - Memory-mapped 8-byte circular keyboard queue (G4)
' Owner: Misael  |  Task: #28 (M5)
'
' Physical map (frame 7 / I/O page):
'   HEAD  @ &HE0   next byte to pop
'   TAIL  @ &HE1   next free slot to push
'   COUNT @ &HE3   bytes waiting (0..8)
'   QUEUE @ &HF0..&HF7
'
' Soft pastel panel on Simulator: hierarchy, HEAD/TAIL cues, Type action.
' =====================================================================
Option Explicit

' Pastel palette (local to this panel — soft UX, not neon)
Private Const CLR_LILAC As Long = &HF5E8E8&       ' RGB(232, 228, 245)
Private Const CLR_MINT As Long = &HDEEBD2&        ' RGB(210, 235, 222)
Private Const CLR_ROSE As Long = &HE1E4FF&        ' RGB(255, 228, 225)
Private Const CLR_CREAM As Long = &HF5F8FC&       ' RGB(252, 248, 245)
Private Const CLR_SKY As Long = &HF7E8D8&         ' RGB(216, 232, 247)
Private Const CLR_INK As Long = &H503C2A&         ' RGB(42, 60, 80)
Private Const CLR_MUTED As Long = &H8C7A6A&       ' RGB(106, 122, 140)
Private Const CLR_HEAD As Long = &HB8D89A&        ' RGB(154, 216, 184) head accent
Private Const CLR_TAIL As Long = &HE8C8D8&        ' RGB(216, 200, 232) tail accent
Private Const CLR_SLOT As Long = &HF0E6DE&        ' RGB(222, 230, 240) empty slot
Private Const CLR_FILLED As Long = &HE8F0D2&      ' RGB(210, 240, 232) occupied

Private Const CELL_W As Single = 48
Private Const CELL_H As Single = 40
Private Const CELL_GAP As Single = 8

Public KbdStatusText As String

' =====================================================================
' Queue primitives (memory-mapped)
' =====================================================================

Public Sub KbdInit()
    RAM(KBD_HEAD) = 0
    RAM(KBD_TAIL) = 0
    RAM(KBD_COUNT) = 0
    Dim i As Integer
    For i = 0 To IO_BUF_SIZE - 1
        RAM(KBD_BASE + i) = 0
    Next i
    KbdStatusText = "Ready — type to enqueue"
End Sub

' Push one ASCII byte. Returns True on success; False if full (no overwrite).
Public Function KbdPush(ByVal ascii As Integer) As Boolean
    Dim cnt As Integer
    Dim tl As Integer
    ascii = ascii And &HFF&
    cnt = RAM(KBD_COUNT) And &HFF&
    If cnt >= IO_BUF_SIZE Then
        KbdStatusText = "Buffer full"
        Debug.Print "KBD: Buffer full"
        KbdPush = False
        Exit Function
    End If
    tl = RAM(KBD_TAIL) And &HFF&
    If tl < 0 Or tl > 7 Then tl = 0
    RAM(KBD_BASE + tl) = ascii
    RAM(KBD_TAIL) = (tl + 1) Mod IO_BUF_SIZE
    RAM(KBD_COUNT) = cnt + 1
    KbdStatusText = "Queued '" & PrintableChar(ascii) & "'  ·  " & CStr(cnt + 1) & "/8"
    KbdPush = True
End Function

' Pop FIFO head. Returns 0 when empty (does not underflow metadata).
Public Function KbdPop() As Integer
    Dim cnt As Integer
    Dim hd As Integer
    Dim v As Integer
    cnt = RAM(KBD_COUNT) And &HFF&
    If cnt <= 0 Then
        KbdStatusText = "Empty — nothing to pop"
        KbdPop = 0
        Exit Function
    End If
    hd = RAM(KBD_HEAD) And &HFF&
    If hd < 0 Or hd > 7 Then hd = 0
    v = RAM(KBD_BASE + hd) And &HFF&
    RAM(KBD_BASE + hd) = 0
    RAM(KBD_HEAD) = (hd + 1) Mod IO_BUF_SIZE
    RAM(KBD_COUNT) = cnt - 1
    KbdStatusText = "Pop '" & PrintableChar(v) & "'  ·  " & CStr(cnt - 1) & "/8 left"
    KbdPop = v
End Function

' Enqueue every character of s, one by one. Stops on Buffer full.
Public Sub KbdTypeString(ByVal s As String)
    Dim i As Long
    Dim ok As Boolean
    If Len(s) = 0 Then
        KbdStatusText = "Nothing typed"
        Exit Sub
    End If
    For i = 1 To Len(s)
        ok = KbdPush(Asc(Mid$(s, i, 1)))
        If Not ok Then Exit Sub
    Next i
End Sub

Private Function PrintableChar(ByVal ascii As Integer) As String
    If ascii >= 32 And ascii < 127 Then
        PrintableChar = Chr$(ascii)
    Else
        PrintableChar = "0x" & Right$("00" & Hex$(ascii), 2)
    End If
End Function

' =====================================================================
' Panel build / refresh (pastel UX)
' =====================================================================

' Called from BuildDashboard. Keyboard half now; soft LCD placeholder for M6.
Public Sub BuildIOPanel(ByVal ws As Worksheet)
    Dim left0 As Single, top0 As Single
    left0 = 28
    top0 = 398

    KbdInit

    ' --- Hierarchy: title → status → queue → actions ---
    SoftBox ws, "kbd_title", left0, top0, 120, 22, "KEYBOARD", CLR_LILAC, CLR_INK, 11, True
    SoftBox ws, "kbd_map", left0 + 128, top0, 200, 22, "0xF0..F7  ·  HEAD E0  TAIL E1", CLR_CREAM, CLR_MUTED, 8, False

    SoftBox ws, "kbd_status", left0, top0 + 26, 400, 22, KbdStatusText, CLR_MINT, RGB(40, 90, 70), 9, True

    BuildQueueCells ws, left0, top0 + 56
    SoftBox ws, "kbd_legend", left0, top0 + 104, 280, 18, _
        "mint = HEAD (pop)   lilac = TAIL (push)", CLR_CREAM, CLR_MUTED, 8, False

    ' Actions — clear primary jobs
    AddKbdButton ws, "btn_KbdType", left0 + 420, top0 + 26, 90, 28, "TYPE", "BtnKbdType", CLR_SKY
    AddKbdButton ws, "btn_KbdPop", left0 + 520, top0 + 26, 90, 28, "POP", "BtnKbdPop", CLR_MINT
    AddKbdButton ws, "btn_KbdClear", left0 + 420, top0 + 60, 90, 28, "CLEAR", "BtnKbdClear", CLR_ROSE

    SoftBox ws, "kbd_hint", left0 + 420, top0 + 96, 190, 36, _
        "TYPE → enter text" & vbLf & "each char queues FIFO", CLR_CREAM, CLR_MUTED, 8, False

    ' LCD slot reserved (M6) — quiet placeholder, no fake controls
    SoftBox ws, "lcd_placeholder", left0 + 290, top0 + 104, 120, 40, "LCD · M6", CLR_LILAC, CLR_MUTED, 9, False

    RefreshKeyboard
End Sub

Private Sub BuildQueueCells(ByVal ws As Worksheet, ByVal left0 As Single, ByVal top0 As Single)
    Dim i As Integer
    Dim x As Single
    For i = 0 To IO_BUF_SIZE - 1
        x = left0 + i * (CELL_W + CELL_GAP)
        SoftBox ws, PFX_KBD & "q" & i, x, top0, CELL_W, CELL_H, "·", CLR_SLOT, CLR_INK, 12, True
        SoftBox ws, PFX_KBD & "tag" & i, x, top0 + CELL_H + 2, CELL_W, 14, CStr(i), CLR_CREAM, CLR_MUTED, 7, False
    Next i
End Sub

Private Sub SoftBox(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, ByVal txt As String, ByVal fill As Long, _
        ByVal ink As Long, ByVal fontSize As Single, ByVal bold As Boolean)
    Dim s As Shape
    On Error Resume Next: ws.Shapes(nm).Delete: On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    s.Name = nm
    s.Line.Visible = msoFalse
    s.Fill.ForeColor.RGB = fill
    With s.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Size = fontSize
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Bold = bold
        .TextRange.Font.Fill.ForeColor.RGB = ink
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
End Sub

Private Sub AddKbdButton(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal txt As String, _
        ByVal macro As String, ByVal fill As Long)
    Dim s As Shape
    SoftBox ws, nm, l, t, w, h, txt, fill, CLR_INK, 10, True
    Set s = ws.Shapes(nm)
    s.OnAction = macro
End Sub

Public Sub RefreshKeyboard()
    Dim i As Integer
    Dim hd As Integer, tl As Integer, cnt As Integer
    Dim v As Integer
    Dim fill As Long
    Dim label As String
    Dim tag As String
    Dim isHead As Boolean, isTail As Boolean

    On Error Resume Next
    hd = RAM(KBD_HEAD) And &HFF&
    tl = RAM(KBD_TAIL) And &HFF&
    cnt = RAM(KBD_COUNT) And &HFF&
    If hd > 7 Then hd = 0
    If tl > 7 Then tl = 0

    For i = 0 To IO_BUF_SIZE - 1
        v = RAM(KBD_BASE + i) And &HFF&
        isHead = (i = hd)
        isTail = (i = tl)

        If SlotOccupied(i, hd, cnt) Then
            fill = CLR_FILLED
            label = PrintableChar(v)
        Else
            fill = CLR_SLOT
            label = "·"
        End If

        ' Arrows under cells: H = next pop, T = next push
        If cnt = 0 And isHead And isTail Then
            tag = "H/T"
            fill = CLR_SKY
        ElseIf cnt >= IO_BUF_SIZE And isHead Then
            tag = "H=T"          ' full: head and tail coincide
            fill = CLR_ROSE
        ElseIf isHead And isTail Then
            tag = "H/T"
            fill = CLR_HEAD
        ElseIf isHead Then
            tag = "▲ H"
            fill = CLR_HEAD
        ElseIf isTail Then
            tag = "▼ T"
            fill = CLR_TAIL
        Else
            tag = CStr(i)
        End If

        UI_SetFill PFX_KBD & "q" & i, fill
        UI_SetText PFX_KBD & "q" & i, label
        UI_SetText PFX_KBD & "tag" & i, tag
    Next i

    If Len(KbdStatusText) = 0 Then KbdStatusText = CStr(cnt) & "/8 waiting"
    UI_SetText "kbd_status", KbdStatusText
    If cnt >= IO_BUF_SIZE Then
        UI_SetFill "kbd_status", CLR_ROSE
    ElseIf cnt = 0 Then
        UI_SetFill "kbd_status", CLR_CREAM
    Else
        UI_SetFill "kbd_status", CLR_MINT
    End If
    On Error GoTo 0
End Sub

' True if index i holds a live queued byte given head + count.
Private Function SlotOccupied(ByVal i As Integer, ByVal hd As Integer, ByVal cnt As Integer) As Boolean
    Dim k As Integer
    If cnt <= 0 Then Exit Function
    For k = 0 To cnt - 1
        If ((hd + k) Mod IO_BUF_SIZE) = i Then
            SlotOccupied = True
            Exit Function
        End If
    Next k
End Function

' =====================================================================
' Button handlers
' =====================================================================

Public Sub BtnKbdType()
    Dim s As String
    On Error Resume Next
    s = CStr(Application.InputBox( _
        Prompt:="Type characters to enqueue (each one is pushed FIFO)." & vbLf & _
                "Example: HELLO", _
        Title:="Keyboard — TYPE", _
        Type:=2))
    On Error GoTo 0
    If s = "False" Or Len(s) = 0 Then
        KbdStatusText = "Type cancelled"
        RefreshKeyboard
        Exit Sub
    End If
    KbdTypeString s
    FlashLastPushed
    RefreshKeyboard
End Sub

Public Sub BtnKbdPop()
    Dim v As Integer
    v = KbdPop()
    If v <> 0 Then
        UI_Highlight PFX_KBD & "q" & ((RAM(KBD_HEAD) - 1 + IO_BUF_SIZE) Mod IO_BUF_SIZE), CLR_MINT
        UI_Wait AnimDelay
    End If
    RefreshKeyboard
End Sub

Public Sub BtnKbdClear()
    KbdInit
    KbdStatusText = "Cleared"
    RefreshKeyboard
End Sub

Private Sub FlashLastPushed()
    Dim tl As Integer
    On Error Resume Next
    tl = (RAM(KBD_TAIL) - 1 + IO_BUF_SIZE) Mod IO_BUF_SIZE
    UI_Highlight PFX_KBD & "q" & tl, CLR_ACTIVE
    UI_Wait AnimDelay
    On Error GoTo 0
End Sub

' =====================================================================
' Test_Keyboard - Immediate self-check for M5 / #28
' =====================================================================
Public Function Test_Keyboard() As Long
    Dim fails As Long
    Dim v As Integer

    KbdInit
    fails = fails + KbdCheck("init empty count", RAM(KBD_COUNT) = 0)
    fails = fails + KbdCheck("pop empty -> 0", KbdPop() = 0)

    fails = fails + KbdCheck("push A", KbdPush(Asc("A")))
    fails = fails + KbdCheck("push B", KbdPush(Asc("B")))
    fails = fails + KbdCheck("push C", KbdPush(Asc("C")))
    fails = fails + KbdCheck("count 3", RAM(KBD_COUNT) = 3)
    fails = fails + KbdCheck("pop A", KbdPop() = Asc("A"))
    fails = fails + KbdCheck("pop B", KbdPop() = Asc("B"))
    fails = fails + KbdCheck("count 1", RAM(KBD_COUNT) = 1)
    fails = fails + KbdCheck("pop C", KbdPop() = Asc("C"))
    fails = fails + KbdCheck("empty again", RAM(KBD_COUNT) = 0)

    ' Fill all 8 — no overwrite
    KbdInit
    Dim i As Integer
    Dim ok As Boolean
    For i = 1 To 8
        ok = KbdPush(Asc("0") + i - 1)
        If Not ok Then fails = fails + 1: Debug.Print "  FAIL  push " & i
    Next i
    fails = fails + KbdCheck("full count 8", RAM(KBD_COUNT) = 8)
    fails = fails + KbdCheck("9th push refused", KbdPush(Asc("X")) = False)
    fails = fails + KbdCheck("still 8 after refuse", RAM(KBD_COUNT) = 8)
    fails = fails + KbdCheck("head byte still '0'", RAM(KBD_BASE + RAM(KBD_HEAD)) = Asc("0"))

    ' Wrap-around FIFO
    KbdInit
    For i = 1 To 5: KbdPush Asc("a") + i - 1: Next i
    For i = 1 To 3: v = KbdPop(): Next i
    For i = 1 To 5: KbdPush Asc("A") + i - 1: Next i
    fails = fails + KbdCheck("wrap count 7", RAM(KBD_COUNT) = 7)
    fails = fails + KbdCheck("wrap pop d", KbdPop() = Asc("d"))
    fails = fails + KbdCheck("wrap pop e", KbdPop() = Asc("e"))
    fails = fails + KbdCheck("wrap pop A", KbdPop() = Asc("A"))

    ' Multi-char type
    KbdInit
    KbdTypeString "Hi!"
    fails = fails + KbdCheck("typed 3", RAM(KBD_COUNT) = 3)
    fails = fails + KbdCheck("typed H", KbdPop() = Asc("H"))
    fails = fails + KbdCheck("typed i", KbdPop() = Asc("i"))
    fails = fails + KbdCheck("typed !", KbdPop() = Asc("!"))

    ' Multi-char stops at full
    KbdInit
    For i = 1 To 6: KbdPush Asc("x"): Next i
    KbdTypeString "ABCDEF"   ' only 2 fit
    fails = fails + KbdCheck("partial type count 8", RAM(KBD_COUNT) = 8)

    Debug.Print IIf(fails = 0, "Test_Keyboard: ALL OK", "Test_Keyboard: " & fails & " FAIL(S)")
    Test_Keyboard = fails
End Function

Private Function KbdCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    KbdCheck = IIf(passed, 0, 1)
End Function
