Attribute VB_Name = "modDashboard"
' =====================================================================
' modDashboard - Main Simulator sheet layout and controls
' Owner: Misael  |  Task: #26 (M3)
'
' Builds the dark dashboard (G7): logo, title, panel slots, Step/Run/
' Pause/Reset, speed slider (AnimDelay 0..1000) and status bar.
' RefreshAll is the single entry point used by Step, Run, Reset and Load.
' Layout targets 1366 x 768 with no gridlines / headings.
' =====================================================================
Option Explicit

Private Const LOGO_FILE As String = "logo-transparent.png"

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
        win.Zoom = 100
    End If
    On Error GoTo EH

    UI_AddBox ws, "lbl_Title", 90, 8, 520, 32, "BYTE BAKERS - Architecture Simulator", CLR_BG
    UI_AddBox ws, "lbl_Clock", 620, 8, 680, 32, StatusText(), CLR_BG

    ' Frames first so empty panels still read as one composition
    BuildPanelFrames ws

    BuildCPUPanel ws          ' A9 (stub until panel lands)
    BuildMiniMemory ws        ' K10 (stub until panel lands)
    BuildIOPanel ws           ' M5 + M6 (stub until I/O lands)

    BuildButtons ws
    BuildSpeedSlider ws

    RefreshAll
    Application.ScreenUpdating = True
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

Private Sub EnsureLogo(ByVal ws As Worksheet)
    Dim s As Shape

    On Error Resume Next
    Set s = ws.Shapes("img_Logo")
    On Error GoTo 0
    If Not s Is Nothing Then
        s.Left = 12
        s.Top = 6
        Exit Sub
    End If

    ' Keep name img_Logo so a real PNG from docs/logo (M1) can replace it
    ' via Insert → Pictures without breaking ClearShapes.
    UI_AddBox ws, "img_Logo", 12, 6, 64, 36, "BB", CLR_ACTIVE
    Debug.Print "BuildDashboard: img_Logo placeholder ready (optional PNG: docs/logo/" & LOGO_FILE & ")"
End Sub

Private Sub BuildPanelFrames(ByVal ws As Worksheet)
    UI_AddBox ws, "frm_CPU", 20, 50, 640, 320, "CPU / BUS (A9)", CLR_BG
    UI_AddBox ws, "frm_MEM", 680, 50, 640, 320, "MEMORY (K10)", CLR_BG
    UI_AddBox ws, "frm_IO", 20, 390, 640, 160, "I/O KBD + LCD (M5/M6)", CLR_BG
    UI_AddBox ws, "frm_CTRL", 680, 390, 640, 160, "CONTROLS", CLR_BG
End Sub

Private Sub BuildButtons(ByVal ws As Worksheet)
    AddButton ws, "btn_Step", 700, 430, "STEP", "BtnStep"
    AddButton ws, "btn_Run", 810, 430, "RUN", "BtnRun"
    AddButton ws, "btn_Pause", 920, 430, "PAUSE", "BtnPause"
    AddButton ws, "btn_Reset", 1030, 430, "RESET", "BtnReset"
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal l As Single, _
        ByVal t As Single, ByVal txt As String, ByVal macro As String)
    Dim s As Shape
    Set s = UI_AddBox(ws, nm, l, t, 95, 34, txt, CLR_ACTIVE)
    s.OnAction = macro
End Sub

Private Sub BuildSpeedSlider(ByVal ws As Worksheet)
    Dim s As Shape
    Static primed As Boolean

    On Error Resume Next
    ws.Shapes("scr_Speed").Delete
    ws.Shapes("lbl_Speed").Delete
    On Error GoTo 0

    If Not primed Then
        If AnimDelay < 0 Or AnimDelay > 1000 Or AnimDelay = 0 Then AnimDelay = 200
        primed = True
    Else
        If AnimDelay < 0 Then AnimDelay = 0
        If AnimDelay > 1000 Then AnimDelay = 1000
    End If

    UI_AddBox ws, "lbl_Speed", 700, 480, 220, 24, "SPEED " & AnimDelay & " ms", CLR_BG
    Set s = ws.Shapes.AddFormControl(xlScrollBar, 700, 510, 420, 18)
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

' =====================================================================
' Button handlers
' =====================================================================
Public Sub BtnStep()
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
    ResetMachine
    RefreshAll
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
    UI_SetText "lbl_Speed", "SPEED " & AnimDelay & " ms"
End Sub

' =====================================================================
' RefreshAll - single entry point (Step, Run, Reset, Load).
' =====================================================================
Public Sub RefreshAll()
    On Error Resume Next
    RefreshCPUPanel          ' A9
    RefreshMemoryPanel       ' K10
    RefreshStats             ' K9
    RefreshKeyboard          ' M5
    RefreshLCD               ' M6
    On Error GoTo 0
    UI_SetText "lbl_Clock", StatusText()
End Sub

Private Function StatusText() As String
    Dim disasm As String
    Dim irVal As Long
    On Error Resume Next
    irVal = CPU.IR
    disasm = Disassemble(irVal)
    If Err.Number <> 0 Or Len(disasm) = 0 Then
        Err.Clear
        disasm = "IR " & Right$("0000" & Hex$(irVal And &HFFFF&), 4)
    End If
    On Error GoTo 0
    StatusText = "CLK " & Clock & "  |  " & disasm & "  |  micro-step " & MicroStep
End Function

' =====================================================================
' Test_Dashboard - Immediate-window self-check for M3.
' Returns failed checks (0 = ALL OK).
' =====================================================================
Public Function Test_Dashboard() As Long
    Dim fails As Long
    Dim ws As Worksheet
    Dim savedDelay As Long

    savedDelay = AnimDelay
    On Error Resume Next
    BuildDashboard
    fails = fails + DashCheck("BuildDashboard no raise", Err.Number = 0)
    Err.Clear

    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    fails = fails + DashCheck("Simulator sheet exists", Not ws Is Nothing)
    fails = fails + DashCheck("btn_Step exists", ShapeExists(ws, "btn_Step"))
    fails = fails + DashCheck("btn_Run exists", ShapeExists(ws, "btn_Run"))
    fails = fails + DashCheck("btn_Pause exists", ShapeExists(ws, "btn_Pause"))
    fails = fails + DashCheck("btn_Reset exists", ShapeExists(ws, "btn_Reset"))
    fails = fails + DashCheck("scr_Speed exists", ShapeExists(ws, "scr_Speed"))
    fails = fails + DashCheck("lbl_Clock exists", ShapeExists(ws, "lbl_Clock"))

    AnimDelay = 350
    ws.Shapes("scr_Speed").ControlFormat.Value = 350
    SpeedChanged
    fails = fails + DashCheck("speed slider -> AnimDelay", AnimDelay = 350)

    RefreshAll
    fails = fails + DashCheck("RefreshAll no raise", Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    AnimDelay = savedDelay
    Debug.Print IIf(fails = 0, "Test_Dashboard: ALL OK", "Test_Dashboard: " & fails & " FAIL(S)")
    Test_Dashboard = fails
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
