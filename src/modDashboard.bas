Attribute VB_Name = "modDashboard"
' =====================================================================
' modDashboard - Main Simulator sheet layout and controls
' Owner: Misael  |  Tasks: #26 (M3), #66 (M14)
'
' Builds the dark dashboard (G7): logo, title, panel frames, Step/Run/
' Pause/Reset, Demo, presentation mode, speed presets + slider (AnimDelay
' 0..1000), color legend and status bar.
' RefreshAll is the single entry point used by Step, Run, Reset and Load.
' The zoom is computed so the whole layout fits the window (1366 x 768).
' =====================================================================
Option Explicit

Private Const LOGO_FILE As String = "logo-transparent.png"

' Layout grid (points). Frames: two columns, 640 wide, 20 pt gutter.
Private Const COL_L As Single = 20
Private Const COL_R As Single = 680
Private Const FRAME_W As Single = 640
Private Const TOP_T As Single = 50
Private Const TOP_H As Single = 344          ' CPU panel signal rows end at 388
Private Const BOT_T As Single = 406
Private Const BOT_H As Single = 320          ' keyboard rows end near 710 (#67)
Private Const HDR_H As Single = 16           ' header tab straddles the frame edge
Private Const MIN_PT As Single = 9

' Controls frame rows
Private Const CTL_X As Single = 700
Private Const ROW_BTN As Single = 426
Private Const ROW_SPEED As Single = 468
Private Const ROW_SLIDER As Single = 500
Private Const ROW_LEGEND As Single = 530

' Speed presets (ms per animation step)
Private Const SPEED_SLOW As Long = 600
Private Const SPEED_NORMAL As Long = 200
Private Const SPEED_FAST As Long = 60
Private Const SPEED_INSTANT As Long = 0

' =====================================================================
' BuildDashboard - rebuild the Simulator sheet from scratch.
' =====================================================================
Public Sub BuildDashboard()
    Dim ws As Worksheet
    Dim win As Window

    On Error GoTo EH
    Set ws = EnsureSimSheet()
    Application.ScreenUpdating = False

    ClearShapes ws
    PaintBackground ws
    EnsureLogo ws

    On Error Resume Next
    Set win = ActiveWindow
    If Not win Is Nothing Then
        win.DisplayGridlines = False
        win.DisplayHeadings = False
    End If
    On Error GoTo EH

    BuildTitleBar ws

    ' Frames first so empty panels still read as one composition
    BuildPanelFrames ws

    BuildCPUPanel ws          ' A9 (stub until panel lands)
    BuildMiniMemory ws        ' K10 (stub until panel lands)
    BuildIOPanel ws           ' M5 + M6 (stub until I/O lands)

    BuildButtons ws
    BuildSpeedSlider ws
    BuildSpeedPresets ws
    BuildLegend ws

    RefreshAll
    Application.ScreenUpdating = True
    FitToScreen
    Debug.Print "BuildDashboard: OK"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Debug.Print "BuildDashboard: ERROR " & Err.Number & " " & Err.Description
End Sub

Private Function EnsureSimSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_SIM
    End If
    Set EnsureSimSheet = ws
End Function

Private Sub PaintBackground(ByVal ws As Worksheet)
    ws.Cells.Interior.Color = CLR_BG
    ws.Range("A1:AZ40").Interior.Color = CLR_BG
End Sub

Private Sub ClearShapes(ByVal ws As Worksheet)
    Dim i As Integer
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name <> "img_Logo" Then
            On Error Resume Next
            ws.Shapes(i).Delete
            On Error GoTo 0
        End If
    Next i
    On Error Resume Next
    ws.Shapes("scr_Speed").Delete
    On Error GoTo 0
End Sub

' Real logo from docs/logo; the "BB" box is only a fallback when the
' PNG is missing (for example a workbook copied without the repo).
Private Sub EnsureLogo(ByVal ws As Worksheet)
    Dim s As Shape
    Dim path As String

    On Error Resume Next
    Set s = ws.Shapes("img_Logo")
    On Error GoTo 0
    path = LogoPath()

    If Not s Is Nothing Then
        If s.Type = msoPicture Or Len(path) = 0 Then
            s.Left = COL_L
            s.Top = 5
            Exit Sub
        End If
        s.Delete
    End If

    If Len(path) > 0 Then
        Set s = ws.Shapes.AddPicture(path, msoFalse, msoTrue, COL_L, 5, 36, 36)
        s.Name = "img_Logo"
        s.LockAspectRatio = msoTrue
    Else
        UI_AddBox ws, "img_Logo", COL_L, 5, 36, 36, "BB", CLR_ACTIVE
        Debug.Print "BuildDashboard: logo placeholder (missing docs/logo/" & LOGO_FILE & ")"
    End If
End Sub

Private Function LogoPath() As String
    Dim p As String
    p = ThisWorkbook.Path & Application.PathSeparator & ".." & Application.PathSeparator & _
        "docs" & Application.PathSeparator & "logo" & Application.PathSeparator & LOGO_FILE
    If Len(ThisWorkbook.Path) > 0 Then
        If Len(Dir$(p)) > 0 Then LogoPath = p
    End If
End Function

Private Sub BuildTitleBar(ByVal ws As Worksheet)
    Dim s As Shape
    Set s = AddLabel(ws, "lbl_Title", COL_L + 46, 8, 520, 34, "BYTE BAKERS  -  Architecture Simulator", _
                     15, True, msoAlignLeft)
    Set s = AddLabel(ws, "lbl_Clock", COL_R, 8, FRAME_W, 34, StatusText(), 10, False, msoAlignRight)
    s.TextFrame2.TextRange.Font.Name = "Consolas"
End Sub

' =====================================================================
' Frames: CLR_SURFACE card, 1 pt CLR_BORDER edge, header tab top-left
' =====================================================================
Private Sub BuildPanelFrames(ByVal ws As Worksheet)
    AddFrame ws, "CPU", COL_L, TOP_T, TOP_H, "PROCESSOR"
    AddFrame ws, "MEM", COL_R, TOP_T, TOP_H, "MEMORY"
    AddFrame ws, "IO", COL_L, BOT_T, BOT_H, "INPUT / OUTPUT"
    AddFrame ws, "CTRL", COL_R, BOT_T, BOT_H, "CONTROLS"
End Sub

Private Sub AddFrame(ByVal ws As Worksheet, ByVal key As String, ByVal l As Single, _
        ByVal t As Single, ByVal h As Single, ByVal title As String)
    Dim s As Shape
    Set s = UI_AddBox(ws, "frm_" & key, l, t, FRAME_W, h, "", CLR_SURFACE)
    s.Adjustments.Item(1) = 0.03
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = CLR_BORDER
    s.Line.Weight = 1

    ' The tab sits on the top edge and hides it, like a fieldset legend
    Set s = UI_AddBox(ws, "hdr_" & key, l + 12, t - HDR_H / 2, Len(title) * 7.5 + 20, HDR_H, title, CLR_BG)
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
    End With
End Sub

' =====================================================================
' Controls
' =====================================================================
Private Sub BuildButtons(ByVal ws As Worksheet)
    AddButton ws, "btn_Step", CTL_X, ROW_BTN, "STEP", "BtnStep"
    AddButton ws, "btn_Run", CTL_X + 100, ROW_BTN, "RUN", "BtnRun"
    AddButton ws, "btn_Pause", CTL_X + 200, ROW_BTN, "PAUSE", "BtnPause"
    AddButton ws, "btn_Reset", CTL_X + 300, ROW_BTN, "RESET", "BtnReset"
    AddButton ws, "btn_Demo", CTL_X + 400, ROW_BTN, "DEMO", "BtnDemo"
    AddButton ws, "btn_Present", CTL_X + 500, ROW_BTN, PresentCaption(), "BtnPresent"
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal txt As String, ByVal macro As String)
    Dim s As Shape
    Set s = UI_AddBox(ws, nm, l, t, 92, 32, txt, ButtonFill(nm))
    s.TextFrame2.TextRange.Font.Name = "Segoe UI"
    s.TextFrame2.TextRange.Font.Size = 10
    s.TextFrame2.TextRange.Font.Bold = msoTrue
    s.Line.Visible = msoTrue
    s.Line.Weight = 1.5
    s.Line.ForeColor.RGB = IIf(ButtonFill(nm) = CLR_SOFT, CLR_ACTIVE, CLR_BORDER)
    s.OnAction = macro
End Sub

' Primary actions on the coral tint (navy text, 4.5:1+); secondary on surface.
Private Function ButtonFill(ByVal nm As String) As Long
    Select Case nm
        Case "btn_Step", "btn_Run", "btn_Pause": ButtonFill = CLR_SOFT
        Case "btn_Reset": ButtonFill = CLR_HIGHZ
        Case Else: ButtonFill = CLR_SURFACE
    End Select
End Function

Private Sub BuildSpeedSlider(ByVal ws As Worksheet)
    Dim s As Shape
    Static primed As Boolean

    On Error Resume Next
    ws.Shapes("scr_Speed").Delete
    ws.Shapes("lbl_Speed").Delete
    On Error GoTo 0

    If Not primed Then
        If AnimDelay < 0 Or AnimDelay > 1000 Or AnimDelay = 0 Then AnimDelay = SPEED_NORMAL
        primed = True
    Else
        If AnimDelay < 0 Then AnimDelay = 0
        If AnimDelay > 1000 Then AnimDelay = 1000
    End If

    AddLabel ws, "lbl_Speed", CTL_X, ROW_SPEED, 120, 24, SpeedText(), 10, True, msoAlignLeft
    Set s = ws.Shapes.AddFormControl(xlScrollBar, CTL_X, ROW_SLIDER, 600, 18)
    s.Name = "scr_Speed"
    With s.ControlFormat
        .Min = 0
        .Max = 1000
        .SmallChange = 10
        .LargeChange = 50
        .Value = AnimDelay
    End With
    s.OnAction = "SpeedChanged"
End Sub

Private Sub BuildSpeedPresets(ByVal ws As Worksheet)
    AddPreset ws, "spd_Slow", 0, "SLOW", "BtnSpeedSlow"
    AddPreset ws, "spd_Normal", 1, "NORMAL", "BtnSpeedNormal"
    AddPreset ws, "spd_Fast", 2, "FAST", "BtnSpeedFast"
    AddPreset ws, "spd_Instant", 3, "INSTANT", "BtnSpeedInstant"
    PaintPresets
End Sub

Private Sub AddPreset(ByVal ws As Worksheet, ByVal nm As String, ByVal idx As Integer, _
        ByVal txt As String, ByVal macro As String)
    Dim s As Shape
    Set s = UI_AddBox(ws, nm, CTL_X + 130 + idx * 82, ROW_SPEED, 76, 24, txt, CLR_SURFACE)
    s.TextFrame2.TextRange.Font.Name = "Segoe UI"
    s.TextFrame2.TextRange.Font.Size = 9
    s.TextFrame2.TextRange.Font.Bold = msoTrue
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = CLR_BORDER
    s.Line.Weight = 1
    s.OnAction = macro
End Sub

' The preset matching AnimDelay is amber; any other slider value lights none.
Private Sub PaintPresets()
    PaintPreset "spd_Slow", SPEED_SLOW
    PaintPreset "spd_Normal", SPEED_NORMAL
    PaintPreset "spd_Fast", SPEED_FAST
    PaintPreset "spd_Instant", SPEED_INSTANT
End Sub

Private Sub PaintPreset(ByVal nm As String, ByVal ms As Long)
    UI_SetFill nm, IIf(AnimDelay = ms, CLR_LOAD, CLR_SURFACE)
End Sub

' Swatch + cream label: the label never sits on a saturated color.
Private Sub BuildLegend(ByVal ws As Worksheet)
    Dim keys As Variant, labels As Variant, colors As Variant
    Dim i As Integer
    Dim x As Single
    Dim s As Shape

    keys = Array("idle", "active", "load", "hit", "miss", "fault")
    labels = Array("IDLE / Hi-Z", "ACTIVE (EN)", "LOAD (LD)", "HIT", "MISS", "FAULT")
    colors = Array(CLR_HIGHZ, CLR_ACTIVE, CLR_LOAD, CLR_HIT, CLR_MISS, CLR_FAULT)

    For i = 0 To UBound(keys)
        x = CTL_X + i * 100
        Set s = UI_AddBox(ws, "leg_sw_" & keys(i), x, ROW_LEGEND + 3, 12, 12, "", CLng(colors(i)))
        s.Adjustments.Item(1) = 0.25
        AddLabel ws, "leg_" & keys(i), x + 16, ROW_LEGEND, 82, 18, CStr(labels(i)), 9, False, msoAlignLeft
    Next i
End Sub

' Text straight on the frame surface (no fill), cream ink.
Private Function AddLabel(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal txt As String, _
        ByVal size As Single, ByVal bold As Boolean, ByVal align As MsoParagraphAlignment) As Shape
    Dim s As Shape
    Set s = UI_AddBox(ws, nm, l, t, w, h, txt, CLR_BG)
    s.Fill.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .WordWrap = msoFalse
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = CLR_TEXT
        .TextRange.ParagraphFormat.Alignment = align
    End With
    Set AddLabel = s
End Function

' =====================================================================
' Fit to screen / presentation mode
' =====================================================================

' Zoom so every dashboard shape is visible in the current window.
Public Sub FitToScreen()
    Dim ws As Worksheet
    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    If Not ActiveSheet Is ws Then Exit Sub
    ws.Range(ws.Cells(1, 1), LastLayoutCell(ws)).Select
    ActiveWindow.Zoom = True
    ws.Cells(1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
Done:
End Sub

' First cell past the right and bottom edge of every shape.
Private Function LastLayoutCell(ByVal ws As Worksheet) As Range
    Dim s As Shape
    Dim maxR As Single, maxB As Single
    Dim r As Long, c As Long

    For Each s In ws.Shapes
        If s.Left + s.Width > maxR Then maxR = s.Left + s.Width
        If s.Top + s.Height > maxB Then maxB = s.Top + s.Height
    Next s
    c = 1
    Do While ws.Cells(1, c).Left + ws.Cells(1, c).Width < maxR + 8 And c < 200
        c = c + 1
    Loop
    r = 1
    Do While ws.Cells(r, 1).Top + ws.Cells(r, 1).Height < maxB + 8 And r < 500
        r = r + 1
    Loop
    Set LastLayoutCell = ws.Cells(r, c)
End Function

Private Function PresentCaption() As String
    PresentCaption = IIf(Application.DisplayFullScreen, "EXIT FULL", "PRESENT")
End Function

' =====================================================================
' Button handlers
' =====================================================================
Public Sub BtnStep()
    If RunFlag Then BlockWhileRunning "btn_Step": Exit Sub
    StepMicro
    RefreshAll
End Sub

Public Sub BtnRun()
    RunProgram
    RefreshAll
End Sub

Public Sub BtnPause()
    RunFlag = False
    RefreshAll
End Sub

Public Sub BtnReset()
    If RunFlag Then BlockWhileRunning "btn_Reset": Exit Sub
    ResetMachine
    ResetControlUnit
    ClearLogSheet
    LogClear
    RefreshAll
End Sub

' Re-entering Step/Reset while RUN loops would corrupt the micro-step.
Private Sub BlockWhileRunning(ByVal btn As String)
    UI_SetText "lbl_Clock", "Pause first - RUN is active"
    UI_SetFill btn, CLR_MISS
    UI_Wait 250
    UI_SetFill btn, ButtonFill(btn)
End Sub

' One fetch on the real CPU panel: PC -> MAR, memory -> MDR, MDR -> IR.
Public Sub BtnDemo()
    If RunFlag Then BlockWhileRunning "btn_Demo": Exit Sub
    UI_AnimateTransfer "reg_PC", "reg_MAR", "wire_INT_ADDR", "PC"
    UI_AnimateTransfer "buf_MEM", "reg_MDR", "wire_DATA", "M[MAR]"
    UI_AnimateTransfer "reg_MDR", "reg_IR", "wire_INT_DATA", "MDR"
    UI_Wait 600
    UI_ClearTransfer
    RefreshAll
End Sub

Public Sub BtnPresent()
    Application.DisplayFullScreen = Not Application.DisplayFullScreen
    UI_SetText "btn_Present", PresentCaption()
    FitToScreen
End Sub

Public Sub SpeedChanged()
    Dim ws As Worksheet
    Dim v As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    v = ws.Shapes("scr_Speed").ControlFormat.Value
    On Error GoTo 0
    If v < 0 Then v = 0
    If v > 1000 Then v = 1000
    AnimDelay = v
    UI_SetText "lbl_Speed", SpeedText()
    PaintPresets
End Sub

Public Sub BtnSpeedSlow()
    SetSpeed SPEED_SLOW
End Sub

Public Sub BtnSpeedNormal()
    SetSpeed SPEED_NORMAL
End Sub

Public Sub BtnSpeedFast()
    SetSpeed SPEED_FAST
End Sub

Public Sub BtnSpeedInstant()
    SetSpeed SPEED_INSTANT
End Sub

' Presets move the slider, so both controls always agree.
Private Sub SetSpeed(ByVal ms As Long)
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SIM).Shapes("scr_Speed").ControlFormat.Value = ms
    On Error GoTo 0
    AnimDelay = ms
    UI_SetText "lbl_Speed", SpeedText()
    PaintPresets
End Sub

Private Function SpeedText() As String
    If AnimDelay = 0 Then
        SpeedText = "SPEED  instant"
    Else
        SpeedText = "SPEED  " & AnimDelay & " ms"
    End If
End Function

' =====================================================================
' RefreshAll - single entry point (Step, Run, Reset, Load).
' =====================================================================
Public Sub RefreshAll()
    On Error Resume Next
    RefreshCPUPanel          ' A9
    RefreshMemoryPanel       ' K10
    RefreshPageTableView     ' K3
    RefreshStats             ' K9
    RefreshKeyboard          ' M5
    RefreshLCD               ' M6
    On Error GoTo 0
    UI_SetText "lbl_Clock", StatusText()
End Sub

Private Function StatusText() As String
    Dim detail As String
    On Error Resume Next
    detail = CurrentMicroText()
    If Err.Number <> 0 Or Len(detail) = 0 Then
        Err.Clear
        detail = "idle"
    End If
    On Error GoTo 0
    StatusText = PhaseWord() & "   " & detail & "    CLK " & Clock
End Function

Private Function PhaseWord() As String
    If MicroStep >= 1 And MicroStep <= 6 Then
        PhaseWord = "FETCH"
    ElseIf MicroStep = 7 Then
        PhaseWord = "DECODE"
    ElseIf MicroStep > 7 Then
        PhaseWord = "EXECUTE"
    Else
        PhaseWord = "IDLE"
    End If
End Function

' =====================================================================
' Test_Dashboard - Immediate-window self-check for M3 / M14.
' Returns failed checks (0 = ALL OK).
' =====================================================================
Public Function Test_Dashboard() As Long
    Dim fails As Long
    Dim ws As Worksheet
    Dim savedDelay As Long
    Dim savedRun As Boolean
    Dim clk0 As Long

    savedDelay = AnimDelay
    savedRun = RunFlag
    ' A check that raises (missing shape, bad property) counts as a failure
    On Error GoTo Broken
    ThisWorkbook.Worksheets(SH_SIM).Activate
    BuildDashboard

    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    fails = fails + DashCheck("Simulator sheet exists", Not ws Is Nothing)
    fails = fails + DashCheck("buttons exist", ShapeExists(ws, "btn_Step") And ShapeExists(ws, "btn_Run") _
        And ShapeExists(ws, "btn_Pause") And ShapeExists(ws, "btn_Reset") _
        And ShapeExists(ws, "btn_Demo") And ShapeExists(ws, "btn_Present"))
    fails = fails + DashCheck("scr_Speed exists", ShapeExists(ws, "scr_Speed"))
    fails = fails + DashCheck("lbl_Clock exists", ShapeExists(ws, "lbl_Clock"))

    ' Frames
    fails = fails + DashCheck("frames use CLR_SURFACE + 1 pt CLR_BORDER", FramesStyled(ws))
    fails = fails + DashCheck("frame headers have no task codes", HeadersClean(ws))
    If Len(LogoPath()) > 0 Then
        fails = fails + DashCheck("logo is the real PNG", ws.Shapes("img_Logo").Type = msoPicture)
    End If
    fails = fails + DashCheck("legend shows 6 states", LegendComplete(ws))

    ' Speed: slider and presets agree
    AnimDelay = 350
    ws.Shapes("scr_Speed").ControlFormat.Value = 350
    SpeedChanged
    fails = fails + DashCheck("speed slider -> AnimDelay", AnimDelay = 350)
    fails = fails + DashCheck("no preset lit at 350 ms", ws.Shapes("spd_Normal").Fill.ForeColor.RGB = CLR_SURFACE)
    BtnSpeedFast
    fails = fails + DashCheck("FAST preset -> 60 ms + slider", AnimDelay = SPEED_FAST _
        And ws.Shapes("scr_Speed").ControlFormat.Value = SPEED_FAST _
        And ws.Shapes("spd_Fast").Fill.ForeColor.RGB = CLR_LOAD)
    BtnSpeedInstant
    fails = fails + DashCheck("INSTANT preset -> 0 ms", AnimDelay = 0 _
        And ws.Shapes("scr_Speed").ControlFormat.Value = 0)

    ' STEP / RESET refused while RUN is active
    clk0 = Clock
    RunFlag = True
    BtnStep
    fails = fails + DashCheck("STEP while RUN shows Pause first", _
        InStr(1, ws.Shapes("lbl_Clock").TextFrame2.TextRange.Text, "Pause first") > 0 And Clock = clk0)
    BtnReset
    fails = fails + DashCheck("RESET while RUN is refused", Clock = clk0)
    fails = fails + DashCheck("button color restored", ws.Shapes("btn_Reset").Fill.ForeColor.RGB = CLR_HIGHZ)
    RunFlag = savedRun

    ' Readability and layout
    fails = fails + DashCheck("dashboard text is " & MIN_PT & " pt or larger", MinFontOk(ws))
    fails = fails + DashCheck("dashboard labels overlap nothing", NoOverlaps(ws))
    fails = fails + DashCheck("panels stay inside their frames", PanelsInsideFrames(ws))
    FitToScreen
    fails = fails + DashCheck("whole layout visible (fitted zoom)", _
        Not Intersect(ActiveWindow.VisibleRange, LastLayoutCell(ws).Offset(-1, -1)) Is Nothing)

    RefreshAll
    On Error GoTo 0

    RunFlag = savedRun
    SetSpeed savedDelay
    Debug.Print IIf(fails = 0, "Test_Dashboard: ALL OK", "Test_Dashboard: " & fails & " FAIL(S)")
    Test_Dashboard = fails
    Exit Function
Broken:
    fails = fails + DashCheck("no runtime error (" & Err.Number & " " & Err.Description & ")", False)
    Resume Next
End Function

Private Function ShapeExists(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(nm)
    ShapeExists = Not s Is Nothing
    On Error GoTo 0
End Function

Private Function DashCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    DashCheck = IIf(passed, 0, 1)
End Function

Private Function FramesStyled(ByVal ws As Worksheet) As Boolean
    Dim k As Variant
    FramesStyled = True
    For Each k In Array("CPU", "MEM", "IO", "CTRL")
        With ws.Shapes("frm_" & k)
            If .Fill.ForeColor.RGB <> CLR_SURFACE Or .Line.ForeColor.RGB <> CLR_BORDER _
                Or .Line.Weight <> 1 Or Not .Line.Visible Then FramesStyled = False
        End With
    Next k
End Function

Private Function HeadersClean(ByVal ws As Worksheet) As Boolean
    Dim k As Variant
    Dim t As String
    HeadersClean = True
    For Each k In Array("CPU", "MEM", "IO", "CTRL")
        t = ws.Shapes("hdr_" & k).TextFrame2.TextRange.Text
        If t Like "*(*" Or t Like "*[AKM]#*" Then HeadersClean = False
    Next k
End Function

Private Function LegendComplete(ByVal ws As Worksheet) As Boolean
    Dim k As Variant
    LegendComplete = True
    For Each k In Array("idle", "active", "load", "hit", "miss", "fault")
        If Not ShapeExists(ws, "leg_sw_" & k) Or Not ShapeExists(ws, "leg_" & k) Then LegendComplete = False
    Next k
End Function

' Shapes this module draws (everything else belongs to a panel).
Private Function IsDashShape(ByVal nm As String) As Boolean
    Select Case True
        Case nm = "lbl_Title", nm = "lbl_Clock", nm = "img_Logo", nm = "lbl_Speed", nm = "scr_Speed"
            IsDashShape = True
        Case nm Like "hdr_*", nm Like "spd_*", nm Like "leg_*", nm Like "frm_*"
            IsDashShape = True
        Case nm = "btn_Step", nm = "btn_Run", nm = "btn_Pause", nm = "btn_Reset", _
             nm = "btn_Demo", nm = "btn_Present"
            IsDashShape = True
    End Select
End Function

Private Function IsTransient(ByVal nm As String) As Boolean
    IsTransient = nm Like "anim_*" Or nm Like "demo_*" Or nm Like "tst_*" Or nm Like "ev_*"
End Function

Private Function MinFontOk(ByVal ws As Worksheet) As Boolean
    Dim s As Shape
    MinFontOk = True
    For Each s In ws.Shapes
        If IsDashShape(s.Name) And s.Type = msoAutoShape Then
            If s.TextFrame2.HasText Then
                If s.TextFrame2.TextRange.Font.Size < MIN_PT Then
                    Debug.Print "    small text: " & s.Name
                    MinFontOk = False
                End If
            End If
        End If
    Next s
End Function

' Every dashboard label/control is clear of every other shape (frames excluded).
Private Function NoOverlaps(ByVal ws As Worksheet) As Boolean
    Dim a As Shape, b As Shape
    NoOverlaps = True
    For Each a In ws.Shapes
        If IsDashShape(a.Name) And Not a.Name Like "frm_*" Then
            For Each b In ws.Shapes
                If b.Name <> a.Name And Not b.Name Like "frm_*" And Not IsTransient(b.Name) Then
                    If Overlap(a, b) Then
                        Debug.Print "    overlap: " & a.Name & " / " & b.Name
                        NoOverlaps = False
                    End If
                End If
            Next b
        End If
    Next a
End Function

' A panel shape whose center is in a frame must fit entirely in it.
Private Function PanelsInsideFrames(ByVal ws As Worksheet) As Boolean
    Dim s As Shape, f As Shape
    Dim k As Variant
    Dim cx As Single, cy As Single
    PanelsInsideFrames = True
    For Each s In ws.Shapes
        If Not IsDashShape(s.Name) And Not IsTransient(s.Name) Then
            cx = s.Left + s.Width / 2
            cy = s.Top + s.Height / 2
            For Each k In Array("CPU", "MEM", "IO", "CTRL")
                Set f = ws.Shapes("frm_" & k)
                If cx > f.Left And cx < f.Left + f.Width And cy > f.Top And cy < f.Top + f.Height Then
                    If s.Left < f.Left Or s.Top < f.Top + HDR_H / 2 Or _
                       s.Left + s.Width > f.Left + f.Width Or s.Top + s.Height > f.Top + f.Height Then
                        Debug.Print "    outside frm_" & k & ": " & s.Name
                        PanelsInsideFrames = False
                    End If
                End If
            Next k
        End If
    Next s
End Function

Private Function Overlap(ByVal a As Shape, ByVal b As Shape) As Boolean
    Const TOL As Single = 0.5
    Overlap = a.Left < b.Left + b.Width - TOL And b.Left < a.Left + a.Width - TOL And _
              a.Top < b.Top + b.Height - TOL And b.Top < a.Top + a.Height - TOL
End Function
