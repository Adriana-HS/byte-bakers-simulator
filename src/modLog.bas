Attribute VB_Name = "modLog"
Option Explicit
' =====================================================================
' modLog - Rich micro-step log on the Log sheet (#77, K13)
' Owner: Karen  |  Rubric: G1, G7
'
' modLog owns the Log sheet. It only READS CPU and Clock (modState) and
' the colors of modConfig. It does not touch modControlUnit.
'
' LogClear      : empties the Log sheet and writes the header row
'                 (frozen, with AutoFilter).
' LogMicroStep  : appends one row per micro-step:
'   Step, CLK, Phase, Transfer, Signals, CW hex, PC, MAR, PADDR, MDR,
'   IR, AX, BX, CX, DX, ZF, CF, SF, Line
'   - registers in fixed-width hex (PC/MAR 3 digits, IR 4, others 2)
'   - flags as 0/1
'   - every register/flag cell that changed from the previous row is
'     amber (CLR_LOAD)
'   - the Phase cell is colored by phase (FETCH / DECODE / EXECUTE)
'   - "Line" is a one-line text: [Step 001] FETCH: PC -> MAR | PC=000 ...
' All text written to the sheet is ASCII.
' =====================================================================

Private Const COL_STEP As Long = 1
Private Const COL_CLK As Long = 2
Private Const COL_PHASE As Long = 3
Private Const COL_TRANSFER As Long = 4
Private Const COL_SIGNALS As Long = 5
Private Const COL_CW As Long = 6
Private Const COL_PC As Long = 7
Private Const COL_MAR As Long = 8
Private Const COL_PADDR As Long = 9
Private Const COL_MDR As Long = 10
Private Const COL_IR As Long = 11
Private Const COL_AX As Long = 12
Private Const COL_BX As Long = 13
Private Const COL_CX As Long = 14
Private Const COL_DX As Long = 15
Private Const COL_ZF As Long = 16
Private Const COL_CF As Long = 17
Private Const COL_SF As Long = 18
Private Const COL_LINE As Long = 19
Private Const LAST_COL As Long = 19

' ---------------------------------------------------------------------
' LogClear: empties the Log sheet and writes the header row.
' ---------------------------------------------------------------------
Public Sub LogClear()
    Dim ws As Worksheet
    Dim oldSU As Boolean

    oldSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo fail

    Set ws = GetLogSheet()
    ws.AutoFilterMode = False
    ws.Cells.Clear

    ' Text format first, so "000" stays text and is not turned into 0
    ws.Range(ws.Columns(COL_PHASE), ws.Columns(LAST_COL)).NumberFormat = "@"

    ws.Range(ws.Cells(1, 1), ws.Cells(1, LAST_COL)).value = _
        Array("Step", "CLK", "Phase", "Transfer", "Signals", "CW hex", _
              "PC", "MAR", "PADDR", "MDR", "IR", "AX", "BX", "CX", "DX", _
              "ZF", "CF", "SF", "Line")
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, LAST_COL))
        .Font.bold = True
        .Font.Color = CLR_TEXT
        .Interior.Color = CLR_BG
    End With

    ws.Columns(COL_STEP).ColumnWidth = 6
    ws.Columns(COL_CLK).ColumnWidth = 6
    ws.Columns(COL_PHASE).ColumnWidth = 10
    ws.Columns(COL_TRANSFER).ColumnWidth = 30
    ws.Columns(COL_SIGNALS).ColumnWidth = 34
    ws.Columns(COL_CW).ColumnWidth = 11
    ws.Range(ws.Columns(COL_PC), ws.Columns(COL_SF)).ColumnWidth = 6
    ws.Columns(COL_LINE).ColumnWidth = 100

    ws.Range(ws.Cells(1, 1), ws.Cells(1, LAST_COL)).AutoFilter
    FreezeHeader ws

    Application.ScreenUpdating = oldSU
    Exit Sub
fail:
    Application.ScreenUpdating = oldSU
    Debug.Print "LogClear ERROR: " & Err.Number & " " & Err.description
End Sub

' ---------------------------------------------------------------------
' LogMicroStep: appends one row with the current CPU and Clock.
'   phase    : "FETCH", "DECODE" or "EXECUTE"
'   cw       : control word of the micro-step
'   cwText   : active signals as text
'   transfer : register transfer as text, for example "PC -> MAR"
' ---------------------------------------------------------------------
Public Sub LogMicroStep(ByVal phase As String, ByVal cw As Long, _
        ByVal cwText As String, ByVal transfer As String)
    Dim ws As Worksheet
    Dim rowN As Long, c As Long, stepNo As Long
    Dim vals(COL_PC To COL_SF) As String
    Dim oldSU As Boolean

    oldSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo fail

    Set ws = GetLogSheet()
    ' Another module may have rewritten the sheet with a different header
    If CStr(ws.Cells(1, COL_STEP).value) <> "Step" Or CStr(ws.Cells(1, COL_LINE).value) <> "Line" Then
        LogClear
    End If

    rowN = LogLastRow(ws) + 1
    stepNo = rowN - 1

    vals(COL_PC) = HexFixed(CPU.PC And &H1FF, 3)
    vals(COL_MAR) = HexFixed(CPU.MAR And &H1FF, 3)
    vals(COL_PADDR) = HexFixed(CPU.paddr And BYTE_MAX, 2)
    vals(COL_MDR) = HexFixed(CPU.MDR And BYTE_MAX, 2)
    vals(COL_IR) = HexFixed(CPU.ir And &HFFFF&, 4)
    vals(COL_AX) = HexFixed(CPU.r(R_AX) And BYTE_MAX, 2)
    vals(COL_BX) = HexFixed(CPU.r(R_BX) And BYTE_MAX, 2)
    vals(COL_CX) = HexFixed(CPU.r(R_CX) And BYTE_MAX, 2)
    vals(COL_DX) = HexFixed(CPU.r(R_DX) And BYTE_MAX, 2)
    vals(COL_ZF) = IIf(CPU.ZF, "1", "0")
    vals(COL_CF) = IIf(CPU.CF, "1", "0")
    vals(COL_SF) = IIf(CPU.SF, "1", "0")

    ws.Cells(rowN, COL_STEP).value = stepNo
    ws.Cells(rowN, COL_CLK).value = Clock
    ws.Cells(rowN, COL_PHASE).value = phase
    ws.Cells(rowN, COL_TRANSFER).value = transfer
    ws.Cells(rowN, COL_SIGNALS).value = cwText
    ws.Cells(rowN, COL_CW).value = "0x" & HexFixed(cw, 7)
    For c = COL_PC To COL_SF
        ws.Cells(rowN, c).value = vals(c)
    Next c

    ' Amber where the value changed from the previous row
    If rowN > 2 Then
        For c = COL_PC To COL_SF
            If CStr(ws.Cells(rowN - 1, c).value) <> vals(c) Then
                ws.Cells(rowN, c).Interior.Color = CLR_LOAD
            End If
        Next c
    End If

    PaintPhase ws.Cells(rowN, COL_PHASE), phase

    ws.Cells(rowN, COL_LINE).value = "[Step " & Format$(stepNo, "000") & "] " & phase & ": " & _
        transfer & " | PC=" & vals(COL_PC) & " MAR=" & vals(COL_MAR) & _
        " PADDR=" & vals(COL_PADDR) & " MDR=" & vals(COL_MDR) & _
        " IR=" & vals(COL_IR) & " AX=" & vals(COL_AX) & " BX=" & vals(COL_BX) & _
        " CX=" & vals(COL_CX) & " DX=" & vals(COL_DX) & _
        " ZF=" & vals(COL_ZF) & " CF=" & vals(COL_CF) & " SF=" & vals(COL_SF)

    Application.ScreenUpdating = oldSU
    Exit Sub
fail:
    Application.ScreenUpdating = oldSU
    Debug.Print "LogMicroStep ERROR: " & Err.Number & " " & Err.description
End Sub

' ---------- helpers ----------

Private Function GetLogSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
    End If
    Set GetLogSheet = ws
End Function

' Last used row of the Step column (1 = only the header)
Private Function LogLastRow(ByVal ws As Worksheet) As Long
    LogLastRow = ws.Cells(ws.Rows.Count, COL_STEP).End(xlUp).row
    If LogLastRow < 1 Then LogLastRow = 1
End Function

' Upper-case hex of v with a fixed number of digits
Private Function HexFixed(ByVal v As Long, ByVal digits As Integer) As String
    HexFixed = Right$(String$(digits, "0") & Hex$(v), digits)
End Function

' Freezes row 1 of the Log sheet, then gives the user back the sheet
' that was active.
Private Sub FreezeHeader(ByVal ws As Worksheet)
    Dim prev As Object
    On Error Resume Next
    Set prev = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    If Not prev Is Nothing Then prev.Activate
    On Error GoTo 0
End Sub

' Phase cell color, only with the modConfig palette
Private Sub PaintPhase(ByVal cel As Range, ByVal phase As String)
    Select Case UCase$(Trim$(phase))
        Case "FETCH"
            cel.Interior.Color = CLR_ACTIVE
            cel.Font.Color = CLR_TEXT
        Case "DECODE"
            cel.Interior.Color = CLR_HIGHZ
            cel.Font.Color = CLR_BG
        Case "EXECUTE"
            cel.Interior.Color = CLR_HIT
            cel.Font.Color = CLR_TEXT
    End Select
End Sub

' ---------- TESTS ----------

Private Function LogCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    LogCheck = IIf(passed, 0, 1)
End Function

' Writes rows from a fake CPU and checks columns, highlight and LogClear.
' Returns the number of failed checks (0 = ALL OK). It empties the Log sheet.
Public Function Test_Log() As Long
    Dim fails As Long
    Dim ws As Worksheet
    Dim saved As TCPU, blank As TCPU
    Dim savedClk As Long
    Dim t As String

    saved = CPU
    savedClk = Clock
    On Error GoTo Broken

    CPU = blank
    LogClear
    Set ws = GetLogSheet()

    ' 1) Header
    fails = fails + LogCheck("header has Step / Phase / Line", CStr(ws.Cells(1, 1).value) = "Step" _
        And CStr(ws.Cells(1, 3).value) = "Phase" And CStr(ws.Cells(1, 19).value) = "Line")
    fails = fails + LogCheck("AutoFilter is on", ws.AutoFilterMode)
    fails = fails + LogCheck("no data rows after LogClear", LogLastRow(ws) = 1)

    ' 2) Four rows from a fake CPU
    Clock = 1
    LogMicroStep "FETCH", &H5&, "EN_PC LD_MAR", "PC -> MAR"
    Clock = 2
    CPU.MAR = &H1A5: CPU.paddr = &HA5
    LogMicroStep "FETCH", &H12&, "MEM_R LD_MDR", "M[MAR] -> MDR"
    Clock = 3
    CPU.MDR = &H3C: CPU.ir = &H3C00&: CPU.ZF = True
    LogMicroStep "DECODE", 0, "none", "decode IR"
    Clock = 4
    CPU.r(R_AX) = &H7F: CPU.CF = True
    LogMicroStep "EXECUTE", 0, "none", "IR(imm) -> AX"

    fails = fails + LogCheck("4 data rows", LogLastRow(ws) = 5)
    fails = fails + LogCheck("Step and CLK columns", CStr(ws.Cells(2, 1).value) = "1" _
        And CStr(ws.Cells(5, 1).value) = "4" And CStr(ws.Cells(5, 2).value) = "4")
    fails = fails + LogCheck("Phase, Transfer and Signals text", CStr(ws.Cells(2, 3).value) = "FETCH" _
        And CStr(ws.Cells(2, 4).value) = "PC -> MAR" And CStr(ws.Cells(2, 5).value) = "EN_PC LD_MAR")
    fails = fails + LogCheck("CW hex has 7 digits", CStr(ws.Cells(2, 6).value) = "0x0000005")
    fails = fails + LogCheck("PC is 3 hex digits", CStr(ws.Cells(2, 7).value) = "000")
    fails = fails + LogCheck("MAR 3 digits, PADDR 2 digits", CStr(ws.Cells(3, 8).value) = "1A5" _
        And CStr(ws.Cells(3, 9).value) = "A5")
    fails = fails + LogCheck("IR is 4 hex digits, MDR 2", CStr(ws.Cells(4, 11).value) = "3C00" _
        And CStr(ws.Cells(4, 10).value) = "3C")
    fails = fails + LogCheck("AX is 2 hex digits", CStr(ws.Cells(5, 12).value) = "7F")
    fails = fails + LogCheck("flags are 0/1", CStr(ws.Cells(4, 16).value) = "1" _
        And CStr(ws.Cells(4, 17).value) = "0" And CStr(ws.Cells(5, 17).value) = "1")

    ' 3) Highlight of changed cells
    fails = fails + LogCheck("first row has no highlight", ws.Cells(2, 8).Interior.Pattern = xlNone)
    fails = fails + LogCheck("MAR changed -> amber", ws.Cells(3, 8).Interior.Color = CLR_LOAD)
    fails = fails + LogCheck("PC unchanged -> no highlight", ws.Cells(3, 7).Interior.Pattern = xlNone)
    fails = fails + LogCheck("MAR unchanged on the next row -> no highlight", ws.Cells(4, 8).Interior.Pattern = xlNone)
    fails = fails + LogCheck("IR and ZF changed -> amber", ws.Cells(4, 11).Interior.Color = CLR_LOAD _
        And ws.Cells(4, 16).Interior.Color = CLR_LOAD)
    fails = fails + LogCheck("AX and CF changed -> amber", ws.Cells(5, 12).Interior.Color = CLR_LOAD _
        And ws.Cells(5, 17).Interior.Color = CLR_LOAD)
    fails = fails + LogCheck("ZF unchanged on the last row -> no highlight", ws.Cells(5, 16).Interior.Pattern = xlNone)

    ' 4) Phase colors
    fails = fails + LogCheck("FETCH is coral", ws.Cells(2, 3).Interior.Color = CLR_ACTIVE)
    fails = fails + LogCheck("DECODE is gray", ws.Cells(4, 3).Interior.Color = CLR_HIGHZ)
    fails = fails + LogCheck("EXECUTE is green", ws.Cells(5, 3).Interior.Color = CLR_HIT)

    ' 5) One-line text
    t = CStr(ws.Cells(3, 19).value)
    fails = fails + LogCheck("line text starts with [Step 002] FETCH:", InStr(t, "[Step 002] FETCH: ") = 1)
    fails = fails + LogCheck("line text has the registers", InStr(t, "MAR=1A5") > 0 And InStr(t, "PADDR=A5") > 0)

    ' 6) LogClear empties the sheet again
    LogClear
    fails = fails + LogCheck("LogClear removes the rows", LogLastRow(ws) = 1 And CStr(ws.Cells(2, 1).value) = "")
    fails = fails + LogCheck("LogClear keeps the header", CStr(ws.Cells(1, 1).value) = "Step" And ws.AutoFilterMode)

    CPU = saved
    Clock = savedClk
    On Error GoTo 0

    Debug.Print IIf(fails = 0, "Test_Log: ALL OK", "Test_Log: " & fails & " FAIL(S)")
    Test_Log = fails
    Exit Function
Broken:
    fails = fails + LogCheck("no runtime error (" & Err.Number & " " & Err.description & ")", False)
    Resume Next
End Function

' Visual demo for the PR screenshot: a fetch and a MOV AX,7 with made-up
' CPU values (it does not run the machine). Leaves the rows on the Log
' sheet. Run Test_Log afterwards to clean the sheet.
Public Sub Demo_Log()
    Dim saved As TCPU, blank As TCPU
    Dim savedClk As Long, cw As Long

    saved = CPU
    savedClk = Clock
    CPU = blank
    LogClear

    Clock = 1
    cw = SIG_EN_PC Or SIG_LD_MAR
    LogMicroStep "FETCH", cw, CW_ToString(cw), "PC -> MAR (addr bus)"
    Clock = 2
    cw = SIG_MEM_R Or SIG_LD_MDR
    CPU.MDR = &H21
    LogMicroStep "FETCH", cw, CW_ToString(cw), "M[MAR] -> MDR (fetch)"
    Clock = 3
    cw = SIG_EN_MDR Or SIG_LD_IR Or SIG_INC_PC
    CPU.ir = &H2100&: CPU.PC = 1
    LogMicroStep "FETCH", cw, CW_ToString(cw), "MDR -> IR (hi), PC+1"
    Clock = 4
    cw = SIG_EN_PC Or SIG_LD_MAR
    CPU.MAR = 1
    LogMicroStep "FETCH", cw, CW_ToString(cw), "PC -> MAR (addr bus)"
    Clock = 5
    cw = SIG_MEM_R Or SIG_LD_MDR
    CPU.paddr = 1: CPU.MDR = &H7
    LogMicroStep "FETCH", cw, CW_ToString(cw), "M[MAR] -> MDR (fetch)"
    Clock = 6
    cw = SIG_EN_MDR Or SIG_LD_IR Or SIG_INC_PC
    CPU.ir = &H2107&: CPU.PC = 2
    LogMicroStep "FETCH", cw, CW_ToString(cw), "MDR -> IR (lo), PC+1"
    Clock = 7
    LogMicroStep "DECODE", 0, "none", "decode IR"
    Clock = 8
    cw = SIG_EN_IR_D Or SIG_LD_REG(R_AX)
    CPU.r(R_AX) = 7
    LogMicroStep "EXECUTE", cw, CW_ToString(cw), "IR(imm) -> AX"

    CPU = saved
    Clock = savedClk
    ThisWorkbook.Worksheets(SH_LOG).Activate
End Sub

