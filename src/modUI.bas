Attribute VB_Name = "modUI"
' =====================================================================
' modUI - Animation engine for the Byte Bakers simulator
' Owner: Misael  |  Task: #25 (M2)
'
' Library every panel uses: create boxes and wires, highlight, High-Z,
' pulse a wire, move a value token and wait according to AnimDelay.
' Basis of the "wire animation" asked by rubric G7 / G1.
'
' Transfers (#65 / M12): UI_AnimateTransfer shows one micro-operation in
' four beats - driver EN (coral), bus (coral), value token travelling in
' right angles along the bus, receiver LD (amber) - and leaves an arrow.
' Every temporary shape is named anim_* and removed by UI_ClearTransfer.
' =====================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

' ms per animation step (speed slider 0..1000). 0 = instant.
Public AnimDelay As Long

' Design tokens derived from the brand palette (modConfig)
Public Const CLR_SURFACE As Long = &H453427&     ' RGB(39, 52, 69)    panel card, navy +1 step
Public Const CLR_BORDER As Long = &H59473A&      ' RGB(58, 71, 89)    1 pt card edge
Public Const CLR_SOFT As Long = &HC8D5F6&        ' RGB(246, 213, 200) coral tint for primary buttons
Public Const CLR_INK As Long = CLR_BG            ' navy text on light fills

Private Const ANIM_PFX As String = "anim_"

' ---------------------------------------------------------------------
' Shp - resolve a shape on Simulator or Memory without crashing.
' Prints a warning to the Immediate window when the name is missing.
' ---------------------------------------------------------------------
Private Function Shp(ByVal nm As String) As Shape
    On Error Resume Next
    Set Shp = ThisWorkbook.Worksheets(SH_SIM).Shapes(nm)
    If Shp Is Nothing Then Set Shp = ThisWorkbook.Worksheets(SH_MEM).Shapes(nm)
    On Error GoTo 0
    If Shp Is Nothing Then Debug.Print "UI: shape not found " & nm
End Function

' Clamp AnimDelay into the accepted 0..1000 ms range.
Private Function ClampedDelay(ByVal ms As Long) As Long
    If ms < 0 Then
        ClampedDelay = 0
    ElseIf ms > 1000 Then
        ClampedDelay = 1000
    Else
        ClampedDelay = ms
    End If
End Function

' Keep AnimDelay inside the accepted 0..1000 ms range (0 = instant).
Private Sub EnsureAnimDelay()
    If AnimDelay < 0 Then AnimDelay = 0
    If AnimDelay > 1000 Then AnimDelay = 1000
End Sub

' =====================================================================
' Shape creation
' =====================================================================

Public Function UI_AddBox(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal txt As String, _
        Optional ByVal fill As Long = -1) As Shape
    Dim s As Shape
    On Error Resume Next: ws.Shapes(nm).Delete: On Error GoTo 0
    If fill = -1 Then fill = CLR_HIGHZ
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    s.Name = nm
    s.Line.Visible = msoFalse
    s.Fill.ForeColor.RGB = fill
    With s.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Size = 10
        .TextRange.Font.Name = "Consolas"
        .TextRange.Font.Fill.ForeColor.RGB = UI_InkFor(fill)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    Set UI_AddBox = s
End Function

Public Function UI_AddWire(ByVal ws As Worksheet, ByVal nm As String, ByVal x1 As Single, _
        ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, Optional ByVal weight As Single = 3) As Shape
    Dim s As Shape
    On Error Resume Next: ws.Shapes(nm).Delete: On Error GoTo 0
    Set s = ws.Shapes.AddConnector(msoConnectorStraight, x1, y1, x2, y2)
    s.Name = nm
    s.Line.Weight = weight
    s.Line.ForeColor.RGB = CLR_HIGHZ
    Set UI_AddWire = s
End Function

' =====================================================================
' Paint helpers
' =====================================================================

Public Sub UI_SetFill(ByVal nm As String, ByVal clr As Long)
    Dim s As Shape
    Set s = Shp(nm)
    If s Is Nothing Then Exit Sub
    If s.Connector Then
        s.Line.ForeColor.RGB = clr
    Else
        s.Fill.ForeColor.RGB = clr
        ' Keep text readable when the fill changes (gray, coral, amber...)
        On Error Resume Next
        If s.Fill.Visible And s.TextFrame2.HasText Then
            s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = UI_InkFor(clr)
        End If
        On Error GoTo 0
    End If
End Sub

' =====================================================================
' Readable ink (WCAG relative luminance)
' =====================================================================

' Navy or cream, whichever contrasts more with the fill.
Public Function UI_InkFor(ByVal fill As Long) As Long
    If fill < 0 Then
        UI_InkFor = CLR_TEXT
    ElseIf UI_Contrast(fill, CLR_INK) >= UI_Contrast(fill, CLR_TEXT) Then
        UI_InkFor = CLR_INK
    Else
        UI_InkFor = CLR_TEXT
    End If
End Function

' Contrast ratio 1..21 between two colors (BGR Longs).
Public Function UI_Contrast(ByVal clrA As Long, ByVal clrB As Long) As Double
    Dim la As Double, lb As Double
    la = Luminance(clrA)
    lb = Luminance(clrB)
    If la < lb Then
        UI_Contrast = (lb + 0.05) / (la + 0.05)
    Else
        UI_Contrast = (la + 0.05) / (lb + 0.05)
    End If
End Function

Private Function Luminance(ByVal clr As Long) As Double
    Luminance = 0.2126 * Linear(clr And &HFF&) _
              + 0.7152 * Linear((clr \ &H100&) And &HFF&) _
              + 0.0722 * Linear((clr \ &H10000) And &HFF&)
End Function

Private Function Linear(ByVal channel As Long) As Double
    Dim v As Double
    v = channel / 255#
    If v <= 0.03928 Then
        Linear = v / 12.92
    Else
        Linear = ((v + 0.055) / 1.055) ^ 2.4
    End If
End Function

Public Sub UI_SetText(ByVal nm As String, ByVal txt As String)
    Dim s As Shape
    Set s = Shp(nm)
    If Not s Is Nothing Then s.TextFrame2.TextRange.Text = txt
End Sub

Public Sub UI_SetWidth(ByVal nm As String, ByVal w As Single)
    Dim s As Shape
    Set s = Shp(nm)
    If Not s Is Nothing Then s.Width = IIf(w < 1, 1, w)
End Sub

Public Sub UI_Highlight(ByVal nm As String, ByVal clr As Long)
    UI_SetFill nm, clr
End Sub

Public Sub UI_SetHighZ(ByVal nm As String)
    UI_SetFill nm, CLR_HIGHZ
End Sub

' =====================================================================
' Motion
' =====================================================================

' Pulses a wire to show the data passing through it.
Public Sub UI_PulseWire(ByVal nm As String, Optional ByVal times As Integer = 2)
    Dim k As Integer
    For k = 1 To times
        UI_SetFill nm, CLR_ACTIVE
        UI_Wait
        UI_SetFill nm, CLR_LOAD
        UI_Wait
    Next k
    UI_SetFill nm, CLR_ACTIVE
End Sub

' Moves a "token" carrying the value from one shape to another.
Public Sub UI_MoveToken(ByVal tokenName As String, ByVal fromShape As String, _
        ByVal toShape As String, ByVal label As String, Optional ByVal steps As Integer = 12)
    Dim a As Shape, b As Shape, tk As Shape
    Dim i As Integer
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
    Dim stepWait As Long

    Set a = Shp(fromShape)
    Set b = Shp(toShape)
    If a Is Nothing Or b Is Nothing Then Exit Sub

    x1 = a.Left + a.Width / 2
    y1 = a.Top + a.Height / 2
    x2 = b.Left + b.Width / 2
    y2 = b.Top + b.Height / 2

    Set tk = UI_AddBox(a.Parent, tokenName, x1 - 18, y1 - 11, 36, 22, label, CLR_ACTIVE)

    EnsureAnimDelay
    If steps < 1 Then steps = 1
    stepWait = ClampedDelay(AnimDelay) \ steps

    For i = 1 To steps
        tk.Left = x1 + (x2 - x1) * i / steps - 18
        tk.Top = y1 + (y2 - y1) * i / steps - 11
        UI_Wait stepWait
    Next i
    tk.Delete
End Sub

Public Sub UI_CacheFlash(ByVal nm As String, ByVal isHit As Boolean)
    UI_SetFill nm, IIf(isHit, CLR_HIT, CLR_MISS)
    UI_SetText "lbl_" & nm, IIf(isHit, "HIT", "MISS")
    UI_Wait
End Sub

' =====================================================================
' Transfers - one micro-operation made visible (#65 / M12)
' =====================================================================

' One micro-operation in four beats. busWire = "" goes straight.
' toShape2 lets one driver load two receivers (e.g. ALU -> AX + FLAGS).
Public Sub UI_AnimateTransfer(ByVal fromShape As String, ByVal toShape As String, _
        ByVal busWire As String, ByVal value As String, Optional ByVal toShape2 As String = "")
    UI_ClearTransfer

    ' Beat 1: driver enables its tri-state buffer
    UI_Highlight fromShape, CLR_ACTIVE
    UI_Badge "en", fromShape, "EN", CLR_ACTIVE
    UI_Wait

    ' Beat 2: the bus carries the value
    If Len(busWire) > 0 Then UI_SetFill busWire, CLR_ACTIVE

    ' Beat 3: the value token travels along the bus
    UI_MoveTokenViaBus "tok", fromShape, toShape, busWire, value
    If Len(toShape2) > 0 Then UI_MoveTokenViaBus "tok2", fromShape, toShape2, busWire, value

    ' Beat 4: receiver(s) latch; the arrow stays until the next transfer
    UI_Highlight toShape, CLR_LOAD
    UI_Badge "ld", toShape, "LD", CLR_LOAD
    UI_ShowArrow "arrow", fromShape, toShape, busWire
    If Len(toShape2) > 0 Then
        UI_Highlight toShape2, CLR_LOAD
        UI_Badge "ld2", toShape2, "LD", CLR_LOAD
        UI_ShowArrow "arrow2", fromShape, toShape2, busWire
    End If
    UI_Wait
End Sub

' Two drivers on one bus: both drivers and the bus turn red.
Public Sub UI_ShowCollision(ByVal driver1 As String, ByVal driver2 As String, ByVal busWire As String)
    UI_ClearTransfer
    UI_Highlight driver1, CLR_MISS
    UI_Highlight driver2, CLR_MISS
    UI_Badge "en", driver1, "EN!", CLR_MISS
    UI_Badge "en2", driver2, "EN!", CLR_MISS
    If Len(busWire) > 0 Then UI_SetFill busWire, CLR_MISS
    UI_Wait
End Sub

' Moves the value token source -> bus -> destination in right angles.
' Instant when AnimDelay = 0 (the arrow alone shows the transfer).
Public Sub UI_MoveTokenViaBus(ByVal tokenName As String, ByVal fromShape As String, _
        ByVal toShape As String, ByVal busWire As String, ByVal label As String)
    Dim a As Shape, b As Shape, tk As Shape
    Dim xs() As Single, ys() As Single
    Dim seg As Integer, i As Integer, steps As Integer
    Dim segLen As Single, totalLen As Single
    Dim travelMs As Long

    Set a = Shp(fromShape)
    Set b = Shp(toShape)
    If a Is Nothing Or b Is Nothing Then Exit Sub
    EnsureAnimDelay
    If AnimDelay = 0 Then Exit Sub

    RoutePoints a, b, busWire, xs, ys
    For seg = 1 To UBound(xs)
        totalLen = totalLen + Abs(xs(seg) - xs(seg - 1)) + Abs(ys(seg) - ys(seg - 1))
    Next seg
    If totalLen <= 0 Then Exit Sub
    travelMs = 2 * ClampedDelay(AnimDelay)

    Set tk = UI_AddBox(a.Parent, ANIM_PFX & tokenName, xs(0) - 20, ys(0) - 10, 40, 20, label, CLR_ACTIVE)
    tk.TextFrame2.TextRange.Font.Bold = msoTrue
    tk.TextFrame2.TextRange.Font.Size = 9

    For seg = 1 To UBound(xs)
        segLen = Abs(xs(seg) - xs(seg - 1)) + Abs(ys(seg) - ys(seg - 1))
        If segLen > 0 Then
            steps = CInt(segLen / 10)
            If steps < 2 Then steps = 2
            For i = 1 To steps
                tk.Left = xs(seg - 1) + (xs(seg) - xs(seg - 1)) * i / steps - 20
                tk.Top = ys(seg - 1) + (ys(seg) - ys(seg - 1)) * i / steps - 10
                UI_Wait CLng(travelMs * (segLen / totalLen) / steps)
            Next i
        End If
    Next seg
    tk.Delete
End Sub

' Arrow along the same route as the token; stays until UI_ClearTransfer.
Public Function UI_ShowArrow(ByVal arrowName As String, ByVal fromShape As String, _
        ByVal toShape As String, Optional ByVal busWire As String = "", _
        Optional ByVal clr As Long = CLR_TEXT) As Shape
    Dim a As Shape, b As Shape, s As Shape
    Dim xs() As Single, ys() As Single
    Dim fb As FreeformBuilder
    Dim i As Integer
    Dim nm As String

    Set a = Shp(fromShape)
    Set b = Shp(toShape)
    If a Is Nothing Or b Is Nothing Then Exit Function
    nm = ANIM_PFX & arrowName

    RoutePoints a, b, busWire, xs, ys
    On Error Resume Next: a.Parent.Shapes(nm).Delete: On Error GoTo 0
    Set fb = a.Parent.Shapes.BuildFreeform(msoEditingAuto, xs(0), ys(0))
    For i = 1 To UBound(xs)
        fb.AddNodes msoSegmentLine, msoEditingAuto, xs(i), ys(i)
    Next i
    Set s = fb.ConvertToShape
    s.Name = nm
    s.Fill.Visible = msoFalse
    s.Line.ForeColor.RGB = clr
    s.Line.Weight = 1.75
    s.Line.DashStyle = msoLineRoundDot
    s.Line.EndArrowheadStyle = msoArrowheadTriangle
    s.Line.EndArrowheadLength = msoArrowheadLong
    s.Line.EndArrowheadWidth = msoArrowheadWide
    Set UI_ShowArrow = s
End Function

' Small chip (EN / LD / Z / EN!) on the top-right corner of a shape.
Public Sub UI_Badge(ByVal badgeName As String, ByVal targetShape As String, _
        ByVal txt As String, ByVal clr As Long)
    Dim t As Shape, s As Shape
    Set t = Shp(targetShape)
    If t Is Nothing Then Exit Sub
    Set s = UI_AddBox(t.Parent, ANIM_PFX & badgeName, t.Left + t.Width - 18, t.Top - 9, 26, 14, txt, clr)
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 7
        .TextRange.Font.Bold = msoTrue
    End With
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = CLR_BG
    s.Line.Weight = 1
End Sub

' Removes every temporary animation shape (tokens, arrows, chips).
Public Sub UI_ClearTransfer()
    ClearAnimOn SH_SIM
    ClearAnimOn SH_MEM
End Sub

Private Sub ClearAnimOn(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(ANIM_PFX)) = ANIM_PFX Then ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

' Route between two shape centers through the bus line (right angles).
' Horizontal bus: down/up to the bus, along it, up/down to the target.
' Vertical bus: the same turned 90 degrees. No bus: straight segment.
Private Sub RoutePoints(ByVal a As Shape, ByVal b As Shape, ByVal busWire As String, _
        ByRef xs() As Single, ByRef ys() As Single)
    Dim bus As Shape
    Dim ax As Single, ay As Single, bx As Single, by As Single
    Dim busX As Single, busY As Single

    ax = a.Left + a.Width / 2: ay = a.Top + a.Height / 2
    bx = b.Left + b.Width / 2: by = b.Top + b.Height / 2
    If Len(busWire) > 0 Then Set bus = Shp(busWire)

    If bus Is Nothing Then
        ReDim xs(0 To 1): ReDim ys(0 To 1)
        xs(0) = ax: ys(0) = ay: xs(1) = bx: ys(1) = by
    ElseIf bus.Width >= bus.Height Then
        busY = bus.Top + bus.Height / 2
        ReDim xs(0 To 3): ReDim ys(0 To 3)
        xs(0) = ax: ys(0) = ay
        xs(1) = ax: ys(1) = busY
        xs(2) = bx: ys(2) = busY
        xs(3) = bx: ys(3) = by
    Else
        busX = bus.Left + bus.Width / 2
        ReDim xs(0 To 3): ReDim ys(0 To 3)
        xs(0) = ax: ys(0) = ay
        xs(1) = busX: ys(1) = ay
        xs(2) = busX: ys(2) = by
        xs(3) = bx: ys(3) = by
    End If
End Sub

' Speed-controlled wait. ms = -1 uses AnimDelay (clamped 0..1000).
Public Sub UI_Wait(Optional ByVal ms As Long = -1)
    Dim t As Single
    Dim total As Long

    EnsureAnimDelay
    If ms = -1 Then
        total = ClampedDelay(AnimDelay)
    Else
        total = ClampedDelay(ms)
    End If

    If total <= 0 Then
        DoEvents
        Exit Sub
    End If

    t = Timer
    Do While (Timer - t) * 1000 < total
        DoEvents
        Sleep 10
    Loop
    DoEvents
End Sub

' =====================================================================
' Demo_UI - visual smoke test of every public helper (G7 / M2).
' Creates temporary shapes on Simulator, runs the API, then cleans up.
' =====================================================================
Public Sub Demo_UI()
    Dim ws As Worksheet
    Dim saved As Long

    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_SIM)

    saved = AnimDelay
    AnimDelay = 80

    Application.ScreenUpdating = True
    Debug.Print "Demo_UI: start"

    UI_AddBox ws, "demo_boxA", 40, 40, 70, 36, "A", CLR_HIGHZ
    UI_AddBox ws, "demo_boxB", 220, 40, 70, 36, "B", CLR_HIGHZ
    UI_AddBox ws, "demo_cache", 40, 100, 80, 36, "L1", CLR_HIGHZ
    UI_AddBox ws, "lbl_demo_cache", 130, 100, 50, 36, "-", CLR_BG
    UI_AddWire ws, "demo_wire", 110, 58, 220, 58, 3

    UI_SetText "demo_boxA", "AX"
    UI_SetWidth "demo_boxA", 80
    UI_Highlight "demo_boxA", CLR_ACTIVE
    UI_Wait 60
    UI_SetHighZ "demo_boxA"
    UI_PulseWire "demo_wire", 2
    UI_MoveToken "demo_tok", "demo_boxA", "demo_boxB", "42", 8
    UI_CacheFlash "demo_cache", True
    UI_Wait 60
    UI_CacheFlash "demo_cache", False

    ' Missing shape must not crash — only warn.
    UI_Highlight "demo_missing_zzz", CLR_ACTIVE

    DemoCleanup ws
    AnimDelay = saved
    Debug.Print "Demo_UI: done OK"
    Exit Sub
EH:
    Debug.Print "Demo_UI: ERROR " & Err.Number & " " & Err.Description
    On Error Resume Next
    DemoCleanup ThisWorkbook.Worksheets(SH_SIM)
    AnimDelay = saved
End Sub

Private Sub DemoCleanup(ByVal ws As Worksheet)
    Dim names As Variant
    Dim i As Integer
    names = Array("demo_boxA", "demo_boxB", "demo_cache", "lbl_demo_cache", _
                  "demo_wire", "demo_tok")
    On Error Resume Next
    For i = LBound(names) To UBound(names)
        ws.Shapes(CStr(names(i))).Delete
    Next i
    On Error GoTo 0
End Sub

' =====================================================================
' Test_UI - Immediate-window self-check (no crash + API present).
' Returns number of failed checks (0 = ALL OK) for automated verify.
' =====================================================================
Public Function Test_UI() As Long
    Dim fails As Long
    Dim saved As Long

    saved = AnimDelay
    AnimDelay = 0

    fails = fails + UiCheck("AnimDelay clamps to 0..1000", _
        ClampedDelay(-5) = 0 And ClampedDelay(1500) = 1000 And ClampedDelay(250) = 250)

    ' Missing shape must not raise.
    On Error Resume Next
    UI_Highlight "no_such_shape_m2", CLR_ACTIVE
    UI_SetHighZ "no_such_shape_m2"
    UI_SetFill "no_such_shape_m2", CLR_LOAD
    UI_SetText "no_such_shape_m2", "x"
    UI_SetWidth "no_such_shape_m2", 10
    UI_PulseWire "no_such_shape_m2", 1
    UI_MoveToken "tok_x", "no_such_a", "no_such_b", "0", 2
    UI_CacheFlash "no_such_shape_m2", True
    fails = fails + UiCheck("missing shapes do not raise", Err.Number = 0)
    On Error GoTo 0

    UI_Wait 0
    fails = fails + UiCheck("UI_Wait(0) returns", True)

    AnimDelay = saved
    Debug.Print IIf(fails = 0, "Test_UI: ALL OK", "Test_UI: " & fails & " FAIL(S)")
    Test_UI = fails
End Function

Private Function UiCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    UiCheck = IIf(passed, 0, 1)
End Function
