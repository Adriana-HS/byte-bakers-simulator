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
' Panel on Simulator (inside frm_IO): brand palette only (modConfig).
'   empty slot  gray   CLR_HIGHZ     occupied  cream  CLR_TEXT
'   HEAD border coral  CLR_ACTIVE    TAIL border amber CLR_LOAD
'   full        dark red CLR_FAULT   text on light fills = CLR_BG (navy)
' =====================================================================
Option Explicit

Private Const KBD_INK As Long = CLR_BG
Private Const NO_FILL As Long = -1

' 12 pt grid inside frm_IO (20, 390, 640, 160)
Private Const PAD_L As Single = 32
Private Const PAD_T As Single = 402
Private Const CELL_W As Single = 36
Private Const CELL_H As Single = 36
Private Const CELL_GAP As Single = 6
Private Const QUEUE_W As Single = 330          ' 8 * CELL_W + 7 * CELL_GAP
Private Const BTN_X As Single = 374            ' PAD_L + QUEUE_W + 12
Private Const BTN_W As Single = 80
Private Const BTN_H As Single = 28
Private Const LCD_X As Single = 466            ' BTN_X + BTN_W + 12
Private Const LCD_W As Single = 182            ' up to the frame padding (648)

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
    KbdStatusText = "Ready - type to enqueue"
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
    KbdStatusText = "Queued '" & PrintableChar(ascii) & "'  |  " & CStr(cnt + 1) & "/8"
    KbdPush = True
End Function

' Pop FIFO head. Returns 0 when empty (does not underflow metadata).
Public Function KbdPop() As Integer
    Dim cnt As Integer
    Dim hd As Integer
    Dim v As Integer
    cnt = RAM(KBD_COUNT) And &HFF&
    If cnt <= 0 Then
        KbdStatusText = "Empty - nothing to pop"
        KbdPop = 0
        Exit Function
    End If
    hd = RAM(KBD_HEAD) And &HFF&
    If hd < 0 Or hd > 7 Then hd = 0
    v = RAM(KBD_BASE + hd) And &HFF&
    RAM(KBD_BASE + hd) = 0
    RAM(KBD_HEAD) = (hd + 1) Mod IO_BUF_SIZE
    RAM(KBD_COUNT) = cnt - 1
    KbdStatusText = "Pop '" & PrintableChar(v) & "'  |  " & CStr(cnt - 1) & "/8 left"
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
' Panel build / refresh
' =====================================================================

' Called from BuildDashboard. Keyboard on the left of frm_IO; the right
' column (LCD_X) stays reserved for the LCD (M6).
Public Sub BuildIOPanel(ByVal ws As Worksheet)
    Dim s As Shape

    KbdInit

    ' Row 1: title + memory map
    SoftBox ws, "kbd_title", PAD_L, PAD_T, 110, 20, "KEYBOARD", NO_FILL, CLR_TEXT, 11, True
    SoftBox ws, "kbd_map", PAD_L + 116, PAD_T, QUEUE_W - 116, 20, _
        "0xF0-F7  HEAD E0  TAIL E1  COUNT E3", NO_FILL, CLR_HIGHZ, 8, False, "Consolas"

    ' Rows 2-3: queue cells + HEAD/TAIL tags
    BuildQueueCells ws, PAD_L, PAD_T + 26

    ' Row 4: status, Row 5: legend
    SoftBox ws, "kbd_status", PAD_L, PAD_T + 84, QUEUE_W, 20, KbdStatusText, NO_FILL, CLR_TEXT, 9, True
    SoftBox ws, "kbd_legend", PAD_L, PAD_T + 110, QUEUE_W, 16, _
        "H = next pop (coral)    T = next push (amber)", NO_FILL, CLR_HIGHZ, 8, False

    ' Actions: TYPE writes at TAIL (amber), POP reads HEAD (coral edge)
    AddKbdButton ws, "btn_KbdType", BTN_X, PAD_T + 26, "TYPE", "BtnKbdType", CLR_LOAD
    AddKbdButton ws, "btn_KbdPop", BTN_X, PAD_T + 60, "POP", "BtnKbdPop", CLR_TEXT
    Set s = ws.Shapes("btn_KbdPop")
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = CLR_ACTIVE
    s.Line.Weight = 2
    AddKbdButton ws, "btn_KbdClear", BTN_X, PAD_T + 94, "CLEAR", "BtnKbdClear", CLR_HIGHZ

    ' LCD column reserved for M6
    SoftBox ws, "lcd_placeholder", LCD_X, PAD_T + 26, LCD_W, 62, "LCD 8x1  (M6)", NO_FILL, CLR_HIGHZ, 9, False
    Set s = ws.Shapes("lcd_placeholder")
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = CLR_HIGHZ
    s.Line.DashStyle = msoLineDash
    s.Line.Weight = 1

    RefreshKeyboard
End Sub

Private Sub BuildQueueCells(ByVal ws As Worksheet, ByVal left0 As Single, ByVal top0 As Single)
    Dim i As Integer
    Dim x As Single
    For i = 0 To IO_BUF_SIZE - 1
        x = left0 + i * (CELL_W + CELL_GAP)
        SoftBox ws, PFX_KBD & "q" & i, x, top0, CELL_W, CELL_H, "", CLR_HIGHZ, KBD_INK, 14, True, "Consolas"
        SoftBox ws, PFX_KBD & "tag" & i, x, top0 + CELL_H + 2, CELL_W, 14, CStr(i), NO_FILL, CLR_HIGHZ, 8, True
    Next i
End Sub

' fill = NO_FILL draws text straight on the navy background.
Private Sub SoftBox(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, ByVal txt As String, ByVal fill As Long, _
        ByVal ink As Long, ByVal fontSize As Single, ByVal bold As Boolean, _
        Optional ByVal fontName As String = "Segoe UI")
    Dim s As Shape
    On Error Resume Next: ws.Shapes(nm).Delete: On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    s.Name = nm
    s.Line.Visible = msoFalse
    If fill = NO_FILL Then
        s.Fill.Visible = msoFalse
    Else
        s.Fill.ForeColor.RGB = fill
    End If
    With s.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .TextRange.Text = txt
        .TextRange.Font.Size = fontSize
        .TextRange.Font.Name = fontName
        .TextRange.Font.Bold = bold
        .TextRange.Font.Fill.ForeColor.RGB = ink
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
End Sub

Private Sub AddKbdButton(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal txt As String, ByVal macro As String, ByVal fill As Long)
    SoftBox ws, nm, l, t, BTN_W, BTN_H, txt, fill, KBD_INK, 10, True
    ws.Shapes(nm).OnAction = macro
End Sub

Public Sub RefreshKeyboard()
    Dim i As Integer
    Dim hd As Integer, tl As Integer, cnt As Integer
    Dim v As Integer
    Dim fill As Long, edge As Long, tagInk As Long
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
            fill = CLR_TEXT
            label = PrintableChar(v)
        Else
            fill = CLR_HIGHZ
            label = ""
        End If

        ' Fill = occupancy; edge + tag = role (H next pop, T next push)
        edge = NO_FILL
        tagInk = CLR_HIGHZ
        tag = CStr(i)
        If isHead And isTail Then
            tag = "H T"
            If cnt >= IO_BUF_SIZE Then
                edge = CLR_ACTIVE        ' full: next action is a pop
                tagInk = CLR_ACTIVE
            Else
                edge = CLR_LOAD          ' empty: next action is a push
                tagInk = CLR_LOAD
            End If
        ElseIf isHead Then
            tag = "H"
            edge = CLR_ACTIVE
            tagInk = CLR_ACTIVE
        ElseIf isTail Then
            tag = "T"
            edge = CLR_LOAD
            tagInk = CLR_LOAD
        End If

        PaintCell i, fill, label, edge, tag, tagInk
    Next i

    If Len(KbdStatusText) = 0 Then KbdStatusText = CStr(cnt) & "/8 waiting"
    UI_SetText "kbd_status", KbdStatusText
    PaintStatus cnt >= IO_BUF_SIZE
    On Error GoTo 0
End Sub

Private Sub PaintCell(ByVal i As Integer, ByVal fill As Long, ByVal label As String, _
        ByVal edge As Long, ByVal tag As String, ByVal tagInk As Long)
    Dim ws As Worksheet
    Dim s As Shape
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    Set s = ws.Shapes(PFX_KBD & "q" & i)
    If Not s Is Nothing Then
        s.Fill.ForeColor.RGB = fill
        s.TextFrame2.TextRange.Text = label
        If edge = NO_FILL Then
            s.Line.Visible = msoFalse
        Else
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = edge
            s.Line.Weight = 2.5
        End If
    End If
    Set s = Nothing
    Set s = ws.Shapes(PFX_KBD & "tag" & i)
    If Not s Is Nothing Then
        s.TextFrame2.TextRange.Text = tag
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tagInk
    End If
    On Error GoTo 0
End Sub

' Full queue = dark red chip with cream text; otherwise plain text on navy.
Private Sub PaintStatus(ByVal isFull As Boolean)
    Dim s As Shape
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(SH_SIM).Shapes("kbd_status")
    If s Is Nothing Then Exit Sub
    If isFull Then
        s.Fill.Visible = msoTrue
        s.Fill.ForeColor.RGB = CLR_FAULT
    Else
        s.Fill.Visible = msoFalse
    End If
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_TEXT
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
        Title:="Keyboard - TYPE", _
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
        UI_Highlight PFX_KBD & "q" & ((RAM(KBD_HEAD) - 1 + IO_BUF_SIZE) Mod IO_BUF_SIZE), CLR_ACTIVE
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
    UI_Highlight PFX_KBD & "q" & tl, CLR_LOAD
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

    ' Fill all 8 - no overwrite
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
