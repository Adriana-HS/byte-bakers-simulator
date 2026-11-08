Attribute VB_Name = "modLoader"
' =====================================================================
' modLoader - program catalog, editor and loader (M8 / #31)
' Owner: Misael
'
' The Programs sheet keeps the assembler status strip on rows 2-3.
' The catalog sits under it. Verify uses Assemble and paints the bad
' line. Load writes VMem at address 0 and leaves pages 0-14 absent.
' =====================================================================
Option Explicit

Private Const ED_ROW As Long = 12
Private Const ED_ROWS As Long = 40
Private Const CAT_ROW As Long = 6

Public Sub BuildProgramsSheet()
    Dim ws As Worksheet
    Dim i As Integer
    Dim band As Range
    Set ws = ProgSheet()
    Application.ScreenUpdating = False
    ws.Visible = xlSheetVisible
    On Error Resume Next
    ws.Tab.Color = CLR_ACTIVE
    On Error GoTo 0

    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 11
    ws.Cells.Font.Color = CLR_BG
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 78
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("A").Hidden = False
    ws.Columns("B").Hidden = False
    ws.Columns(26).Hidden = True

    Set band = ws.Range("A5:B5")
    band.Value = ""
    ws.Range("A5").Value = "1   CATALOG"
    band.Font.Bold = True
    band.Font.Color = CLR_TEXT
    band.Font.Size = 12
    band.Interior.Color = CLR_BG
    ws.Rows(5).RowHeight = 22

    For i = 1 To 3
        ws.Rows(CAT_ROW + i - 1).RowHeight = 22
        ws.Cells(CAT_ROW + i - 1, 1).Value = CatalogName(i)
        ws.Cells(CAT_ROW + i - 1, 2).Value = CatalogBlurb(i)
        ws.Cells(CAT_ROW + i - 1, 1).Font.Color = CLR_BG
        ws.Cells(CAT_ROW + i - 1, 2).Font.Color = CLR_BG
        ws.Cells(CAT_ROW + i - 1, 1).Font.Bold = True
        ws.Cells(CAT_ROW + i - 1, 1).Interior.Color = CLR_HIGHZ
        ws.Cells(CAT_ROW + i - 1, 2).Interior.Color = CLR_HIGHZ
    Next i

    ws.Rows(9).RowHeight = 20
    ws.Range("A9").Value = "saved"
    ws.Range("B9").Value = "A program you save shows up on this row"
    ws.Range("A9:B9").Font.Color = CLR_BG
    ws.Range("A9:B9").Font.Bold = False
    ws.Range("A9:B9").Interior.Color = CLR_TEXT

    ws.Rows(10).RowHeight = 40
    ws.Range("A10:B10").Interior.Color = CLR_BG
    AddLoaderButton ws, "btn_Use1", 12, ws.Rows(10).Top + 6, "USE 1", "BtnUse1"
    AddLoaderButton ws, "btn_Use2", 94, ws.Rows(10).Top + 6, "USE 2", "BtnUse2"
    AddLoaderButton ws, "btn_Use3", 176, ws.Rows(10).Top + 6, "USE 3", "BtnUse3"
    AddLoaderButton ws, "btn_LVerify", 258, ws.Rows(10).Top + 6, "VERIFY", "BtnVerify"
    AddLoaderButton ws, "btn_LLoad", 340, ws.Rows(10).Top + 6, "LOAD", "BtnLoad"
    AddLoaderButton ws, "btn_LSave", 422, ws.Rows(10).Top + 6, "SAVE", "BtnSave"
    AddLoaderButton ws, "btn_LSaved", 504, ws.Rows(10).Top + 6, "SAVED", "BtnUseSaved"

    ws.Range("A11").Value = "2   EDITOR"
    ws.Range("B11").Value = "One instruction per row. Verify, then Load."
    ws.Range("A11:B11").Font.Bold = True
    ws.Range("A11:B11").Font.Color = CLR_TEXT
    ws.Range("A11:B11").Font.Size = 12
    ws.Range("A11:B11").Interior.Color = CLR_BG
    ws.Rows(11).RowHeight = 22

    PaintEditorChrome ws
    OpenCatalog 1
    ws.Activate
    On Error Resume Next
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Public Sub BtnUse1()
    OpenCatalog 1
End Sub

Public Sub BtnUse2()
    OpenCatalog 2
End Sub

Public Sub BtnUse3()
    OpenCatalog 3
End Sub

Public Sub BtnVerify()
    VerifyEditor
End Sub

Public Sub BtnLoad()
    LoadEditor
End Sub

Public Sub BtnSave()
    Dim nm As String
    nm = InputBox("Name for this program", "Byte Bakers")
    If Len(Trim$(nm)) = 0 Then Exit Sub
    SaveProgram CStr(nm)
End Sub

Public Sub OpenCatalog(ByVal index As Integer)
    Dim ws As Worksheet
    Dim lines As Variant
    Dim i As Long
    Set ws = ProgSheet()
    If index < 1 Or index > 3 Then Exit Sub
    MarkCatalog ws, index
    lines = CatalogLines(index)
    ClearEditor ws
    For i = LBound(lines) To UBound(lines)
        ws.Cells(ED_ROW + i - LBound(lines), 2).Value = CStr(lines(i))
    Next i
End Sub

Public Sub VerifyEditor()
    Dim lines As Variant
    Dim i As Long
    Dim msg As String
    Dim lineNo As Long
    Dim ws As Worksheet
    Set ws = ProgSheet()
    ClearMarks ws
    lines = EditorLines(ws)
    Assemble lines
    If AsmErrors Is Nothing Then Exit Sub
    For i = 1 To AsmErrors.Count
        msg = CStr(AsmErrors(i))
        lineNo = LineNoOf(msg)
        If lineNo > 0 Then
            ws.Cells(ED_ROW + lineNo - 1, 2).Interior.Color = CLR_MISS
            ws.Cells(ED_ROW + lineNo - 1, 2).Font.Color = CLR_BG
            ws.Cells(ED_ROW + lineNo - 1, 2).Font.Bold = True
        End If
    Next i
End Sub

Public Sub LoadEditor()
    Dim lines As Variant
    Dim raw As Variant
    Dim bytes() As Integer
    Dim i As Long
    Dim n As Long
    Dim ws As Worksheet
    Set ws = ProgSheet()
    lines = EditorLines(ws)
    raw = Assemble(lines)
    If Not AsmErrors Is Nothing Then
        If AsmErrors.Count > 0 Then
            VerifyEditor
            Exit Sub
        End If
    End If
    If IsEmpty(raw) Then Exit Sub
    On Error Resume Next
    n = UBound(raw) - LBound(raw) + 1
    If Err.Number <> 0 Or n <= 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    ReDim bytes(LBound(raw) To UBound(raw))
    For i = LBound(raw) To UBound(raw)
        bytes(i) = CInt(raw(i)) And BYTE_MAX
    Next i
    ClearVMem
    VMemWriteBlock 0, bytes
    ResetMachine
    DemandPaging
    RefreshAll
    ThisWorkbook.Worksheets(SH_SIM).Activate
End Sub

Public Sub SaveProgram(ByVal nm As String)
    Dim ws As Worksheet
    Dim lines As Variant
    Dim i As Long
    Set ws = ProgSheet()
    ws.Range("A9").Value = nm
    ws.Range("B9").Value = "Saved from the editor"
    ws.Range("A9").Font.Color = CLR_BG
    ws.Range("B9").Font.Color = CLR_BG
    ws.Range("A9").Font.Bold = True
    ws.Range("A9").Interior.Color = CLR_LOAD
    ws.Range("B9").Interior.Color = CLR_LOAD
    lines = EditorLines(ws)
    ws.Range("Z12:Z80").ClearContents
    For i = LBound(lines) To UBound(lines)
        ws.Cells(12 + i - LBound(lines), 26).Value = CStr(lines(i))
    Next i
End Sub

Public Sub BtnUseSaved()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ProgSheet()
    ClearEditor ws
    For i = 12 To 80
        If Len(CStr(ws.Cells(i, 26).Value)) > 0 Then
            ws.Cells(ED_ROW + i - 12, 2).Value = ws.Cells(i, 26).Value
        End If
    Next i
End Sub

Public Function Test_Loader() As Long
    Dim fails As Long
    Dim ws As Worksheet
    Dim i As Integer
    Dim absent As Boolean

    BuildProgramsSheet
    Set ws = ProgSheet()
    fails = fails + LdCheck("three names", _
        ws.Cells(6, 1).Value = "Multiply" And _
        ws.Cells(7, 1).Value = "Keyboard echo" And _
        ws.Cells(8, 1).Value = "Page faults")
    fails = fails + LdCheck("three blurbs", _
        Len(CStr(ws.Cells(6, 2).Value)) > 0 And _
        Len(CStr(ws.Cells(7, 2).Value)) > 0 And _
        Len(CStr(ws.Cells(8, 2).Value)) > 0)

    OpenCatalog 1
    fails = fails + LdCheck("editor got program 1", ws.Cells(ED_ROW, 2).Value = "MOV AX, 0")

    ws.Cells(ED_ROW, 2).Value = "NOPE AX, 1"
    VerifyEditor
    fails = fails + LdCheck("bad line is red", ws.Cells(ED_ROW, 2).Interior.Color = CLR_MISS)
    fails = fails + LdCheck("bad line has a message", Not AsmErrors Is Nothing And AsmErrors.Count > 0)
    fails = fails + LdCheck("message is on the sheet", InStr(CStr(ws.Range("B3").Value), "Line ") = 1)

    OpenCatalog 2
    LoadEditor
    absent = True
    For i = 0 To NUM_VPAGES - 1
        If i <> IO_VPAGE Then
            If PageTable(i).Present Then absent = False
        End If
    Next i
    fails = fails + LdCheck("pages 0-14 are absent", absent)
    fails = fails + LdCheck("VMem holds the first opcode", VMem(0) <> 0 Or VMem(1) <> 0)

    ws.Cells(ED_ROW, 2).Value = "MOV AX, 9"
    SaveProgram "Mine"
    ws.Cells(ED_ROW, 2).Value = "HLT"
    BtnUseSaved
    fails = fails + LdCheck("saved program reloads", ws.Cells(ED_ROW, 2).Value = "MOV AX, 9")

    Debug.Print IIf(fails = 0, "Test_Loader: ALL OK", "Test_Loader: " & fails & " FAIL(S)")
    Test_Loader = fails
End Function

Private Sub DemandPaging()
    Dim i As Integer
    For i = 0 To NUM_VPAGES - 1
        If i <> IO_VPAGE Then PageTable(i).Present = False
    Next i
End Sub

Private Sub MarkCatalog(ByVal ws As Worksheet, ByVal index As Integer)
    Dim i As Integer
    For i = 1 To 3
        ws.Cells(CAT_ROW + i - 1, 1).Interior.Color = CLR_HIGHZ
        ws.Cells(CAT_ROW + i - 1, 2).Interior.Color = CLR_HIGHZ
    Next i
    ws.Cells(CAT_ROW + index - 1, 1).Interior.Color = CLR_LOAD
    ws.Cells(CAT_ROW + index - 1, 2).Interior.Color = CLR_LOAD
End Sub

Private Sub ClearEditor(ByVal ws As Worksheet)
    ws.Range(ws.Cells(ED_ROW, 2), ws.Cells(ED_ROW + ED_ROWS - 1, 2)).ClearContents
    ClearMarks ws
End Sub

Private Sub PaintEditorChrome(ByVal ws As Worksheet)
    Dim i As Long
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(ED_ROW, 1), ws.Cells(ED_ROW + ED_ROWS - 1, 2))
    rng.Interior.Color = CLR_TEXT
    rng.Font.Name = "Consolas"
    rng.Font.Size = 12
    rng.Font.Color = CLR_BG
    rng.Font.Bold = False
    For i = 1 To ED_ROWS
        ws.Rows(ED_ROW + i - 1).RowHeight = 18
        ws.Cells(ED_ROW + i - 1, 1).Value = Right$("0" & CStr(i), 2)
        ws.Cells(ED_ROW + i - 1, 1).HorizontalAlignment = xlRight
    Next i
End Sub

Private Sub ClearMarks(ByVal ws As Worksheet)
    PaintEditorChrome ws
End Sub

Private Function EditorLines(ByVal ws As Worksheet) As Variant
    Dim buf() As String
    Dim n As Long
    Dim i As Long
    Dim last As Long
    last = ED_ROW
    For i = ED_ROW To ED_ROW + ED_ROWS - 1
        If Len(Trim$(CStr(ws.Cells(i, 2).Value))) > 0 Then last = i
    Next i
    n = last - ED_ROW
    ReDim buf(0 To n)
    For i = 0 To n
        buf(i) = CStr(ws.Cells(ED_ROW + i, 2).Value)
    Next i
    EditorLines = buf
End Function

Private Function LineNoOf(ByVal msg As String) As Long
    Dim rest As String
    Dim p As Long
    If Left$(msg, 5) <> "Line " Then Exit Function
    rest = Mid$(msg, 6)
    p = InStr(rest, ":")
    If p < 2 Then Exit Function
    On Error Resume Next
    LineNoOf = CLng(Left$(rest, p - 1))
    On Error GoTo 0
End Function

Private Function ProgSheet() As Worksheet
    On Error Resume Next
    Set ProgSheet = ThisWorkbook.Worksheets(SH_PROG)
    On Error GoTo 0
    If ProgSheet Is Nothing Then
        Set ProgSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ProgSheet.Name = SH_PROG
    End If
End Function

Private Sub AddLoaderButton(ByVal ws As Worksheet, ByVal nm As String, _
        ByVal l As Single, ByVal t As Single, ByVal txt As String, ByVal macro As String)
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
    Set s = UI_AddBox(ws, nm, l, t, 76, 28, txt, CLR_SOFT)
    s.TextFrame2.TextRange.Font.Name = "Segoe UI"
    s.TextFrame2.TextRange.Font.Size = 10
    s.TextFrame2.TextRange.Font.Bold = msoTrue
    s.OnAction = macro
End Sub

Private Function CatalogName(ByVal index As Integer) As String
    Select Case index
        Case 1: CatalogName = "Multiply"
        Case 2: CatalogName = "Keyboard echo"
        Case 3: CatalogName = "Page faults"
    End Select
End Function

Private Function CatalogBlurb(ByVal index As Integer) As String
    Select Case index
        Case 1: CatalogBlurb = "7 x 5 in AX by repeated addition. Result stored at 0x080."
        Case 2: CatalogBlurb = "Read 8 keys and show each one on the LCD."
        Case 3: CatalogBlurb = "Scattered loads and stores that force page replacement."
    End Select
End Function

Private Function CatalogLines(ByVal index As Integer) As Variant
    Select Case index
        Case 1
            CatalogLines = Array( _
                "MOV AX, 0", "MOV BX, 7", "MOV CX, 5", _
                "mult: ADD AX, BX", "DEC CX", "JNZ mult", _
                "MOV DX, AX", "STORE [0x080], DX", "HLT")
        Case 2
            CatalogLines = Array( _
                "MOV CX, 8", "read: IN AX, [KBD]", "CMP AX, 0", _
                "JZ read", "OUT [LCD], AX", "DEC CX", "JNZ read", "HLT")
        Case 3
            CatalogLines = Array( _
                "MOV CX, 3", "pass: LOAD AX, [0x020]", "LOAD BX, [0x060]", _
                "ADD AX, BX", "STORE [0x0A0], AX", "LOAD DX, [0x0E0]", _
                "STORE [0x120], DX", "LOAD BX, [0x160]", "STORE [0x1A0], BX", _
                "LOAD AX, [0x1C0]", "DEC CX", "JNZ pass", "HLT")
    End Select
End Function

Private Function LdCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    LdCheck = IIf(passed, 0, 1)
End Function
