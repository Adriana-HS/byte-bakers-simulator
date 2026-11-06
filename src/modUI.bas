Attribute VB_Name = "modUI"
' =====================================================================
' modUI - Animation engine for the Byte Bakers simulator
' Owner: Misael  |  Task: #25 (M2)
'
' Library every panel uses: create boxes and wires, highlight, High-Z,
' pulse a wire, move a value token and wait according to AnimDelay.
' Basis of the "wire animation" asked by rubric G7 / G1.
' =====================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

' ms per animation step (speed slider 0..1000). 0 = instant.
Public AnimDelay As Long

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
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    s.Name = nm
    s.Line.Visible = msoFalse
    s.Fill.ForeColor.RGB = IIf(fill = -1, CLR_HIGHZ, fill)
    With s.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Size = 10
        .TextRange.Font.Name = "Consolas"
        .TextRange.Font.Fill.ForeColor.RGB = CLR_TEXT
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
    End If
End Sub

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
