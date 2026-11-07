Attribute VB_Name = "modAssembler"
' =====================================================================
' modAssembler - Two-pass assembler with syntax validation (G5)
' Owner: Misael  |  Task: #27 (M4)
'
' Pass 1 records labels -> address. Pass 2 encodes 16-bit words via
' OpcodeInfo + Encode. Errors go into AsmErrors with line numbers.
' =====================================================================
Option Explicit

Public AsmErrors As Collection
Public AsmStatusText As String          ' soft UX status for Programs sheet

Private Labels As Object                ' Scripting.Dictionary

' ---------------------------------------------------------------------
' Assemble - lines() of source text -> byte array (Integer 0..255 each).
' On failure still returns whatever was encoded; check AsmErrors.Count.
' ---------------------------------------------------------------------
Public Function Assemble(lines As Variant) As Variant
    Dim i As Long
    Dim addr As Integer
    Dim out() As Integer
    Dim n As Long
    Dim code As String
    Dim word As Long
    Dim emptyArr() As Integer

    Set AsmErrors = New Collection
    Set Labels = CreateObject("Scripting.Dictionary")
    AsmStatusText = ""

    ' Pass 1: labels -> byte address (each instruction = 2 bytes)
    addr = 0
    For i = LBound(lines) To UBound(lines)
        code = StripLabel(CleanLine(CStr(lines(i))), True, addr)
        If Len(code) > 0 Then addr = addr + 2
    Next i

    If addr = 0 Then
        ReDim emptyArr(0 To -1)
        Assemble = emptyArr
        AsmStatusText = "Nothing to assemble"
        PaintAsmStatus False
        Exit Function
    End If

    ReDim out(0 To addr - 1)
    n = 0
    For i = LBound(lines) To UBound(lines)
        code = StripLabel(CleanLine(CStr(lines(i))), False, 0)
        If Len(code) > 0 Then
            word = EncodeLine(code, i + 1)
            out(n) = CInt((word \ 256) And &HFF&)
            out(n + 1) = CInt(word And &HFF&)
            n = n + 2
        End If
    Next i

    Assemble = out
    If AsmErrors.Count = 0 Then
        AsmStatusText = "OK - " & CStr(n) & " bytes"
        PaintAsmStatus True
    Else
        AsmStatusText = "ERROR - " & CStr(AsmErrors.Count) & " problem(s)"
        PaintAsmStatus False
    End If
End Function

Private Function CleanLine(ByVal s As String) As String
    Dim p As Long
    s = Replace(s, vbTab, " ")
    p = InStr(s, ";")
    If p > 0 Then s = Left$(s, p - 1)
    CleanLine = Trim$(s)
End Function

' If recordLabel, stores label -> addr. Always returns code after the label.
Private Function StripLabel(ByVal code As String, ByVal recordLabel As Boolean, ByVal addr As Integer) As String
    Dim p As Long
    Dim lab As String
    p = InStr(code, ":")
    If p = 0 Then
        StripLabel = code
        Exit Function
    End If
    lab = UCase$(Trim$(Left$(code, p - 1)))
    If recordLabel Then
        If Len(lab) = 0 Then
            ' ignore bare ":"
        ElseIf Labels.Exists(lab) Then
            AsmErrors.Add "Duplicate label '" & lab & "'"
        Else
            Labels.Add lab, addr
        End If
    End If
    StripLabel = Trim$(Mid$(code, p + 1))
End Function

Private Function EncodeLine(ByVal code As String, ByVal lineNo As Long) As Long
    Dim mn As String
    Dim rest As String
    Dim parts() As String
    Dim fmt As String
    Dim reg As Integer
    Dim operand As Integer
    Dim op As Integer
    Dim ok As Boolean

    mn = UCase$(Trim$(Split(code, " ")(0)))
    rest = Trim$(Mid$(code, Len(mn) + 1))
    ok = BuildOperands(rest, fmt, reg, operand, lineNo)
    If Not ok Then Exit Function

    op = OpcodeInfo(mn, fmt)
    If op < 0 Then
        AddErr lineNo, "'" & code & "' is not a valid instruction (mnemonic/operands)"
        Exit Function
    End If

    ' Range checks after format is known
    Select Case fmt
        Case "RI", "R"
            If fmt = "RI" Then
                If operand < 0 Or operand > BYTE_MAX Then
                    AddErr lineNo, "immediate out of range 0..255: " & operand
                    Exit Function
                End If
            End If
        Case "RM", "MR", "A"
            If operand < 0 Or operand > VADDR_MAX Then
                AddErr lineNo, "address out of range 0..511: " & operand
                Exit Function
            End If
    End Select

    EncodeLine = Encode(op, reg, operand)
End Function

' Parses operand list into fmt / reg / operand for OpcodeInfo + Encode.
Private Function BuildOperands(ByVal rest As String, ByRef fmt As String, _
        ByRef reg As Integer, ByRef operand As Integer, ByVal lineNo As Long) As Boolean
    Dim a As String, b As String
    Dim ta As String, tb As String
    Dim va As Long, vb As Long
    Dim p As Long

    reg = 0
    operand = 0
    fmt = ""
    BuildOperands = True

    If Len(rest) = 0 Then
        fmt = ""
        Exit Function
    End If

    p = InStr(rest, ",")
    If p = 0 Then
        a = Trim$(rest)
        ta = OpKind(a, va, lineNo)
        If Len(ta) = 0 Then BuildOperands = False: Exit Function
        Select Case ta
            Case "R"
                fmt = "R"
                reg = CInt(va)
            Case "I", "A"
                ' bare number/label as jump/address target
                fmt = "A"
                operand = CInt(va)
            Case Else
                AddErr lineNo, "unexpected single operand: " & a
                BuildOperands = False
        End Select
        Exit Function
    End If

    a = Trim$(Left$(rest, p - 1))
    b = Trim$(Mid$(rest, p + 1))
    ta = OpKind(a, va, lineNo)
    tb = OpKind(b, vb, lineNo)
    If Len(ta) = 0 Or Len(tb) = 0 Then BuildOperands = False: Exit Function

    If ta = "R" And tb = "R" Then
        fmt = "RR": reg = CInt(va): operand = CInt(vb)
    ElseIf ta = "R" And (tb = "I" Or tb = "A") Then
        fmt = "RI": reg = CInt(va): operand = CInt(vb)
    ElseIf ta = "R" And tb = "M" Then
        fmt = "RM": reg = CInt(va): operand = CInt(vb)
    ElseIf ta = "M" And tb = "R" Then
        fmt = "MR": reg = CInt(vb): operand = CInt(va)
    ElseIf ta = "R" And tb = "K" Then
        fmt = "RK": reg = CInt(va): operand = 0
    ElseIf ta = "L" And tb = "R" Then
        fmt = "LR": reg = CInt(vb): operand = 0
    Else
        AddErr lineNo, "wrong operands: " & a & ", " & b
        BuildOperands = False
    End If
End Function

' Returns kind: R / I / A / M / K / L, and numeric value in v. "" on hard error.
Private Function OpKind(ByVal s As String, ByRef v As Long, ByVal lineNo As Long) As String
    Dim inner As String
    Dim ri As Integer
    s = Trim$(s)
    If Len(s) = 0 Then
        AddErr lineNo, "missing operand"
        Exit Function
    End If

    If Left$(s, 1) = "[" And Right$(s, 1) = "]" Then
        inner = Trim$(Mid$(s, 2, Len(s) - 2))
        If UCase$(inner) = "KBD" Then
            OpKind = "K": v = 0: Exit Function
        End If
        If UCase$(inner) = "LCD" Then
            OpKind = "L": v = 0: Exit Function
        End If
        v = ParseNumber(inner)
        If v < 0 Then
            AddErr lineNo, "undefined label or bad address '" & inner & "'"
            Exit Function
        End If
        OpKind = "M"
        Exit Function
    End If

    ri = RegIndex(s)
    If ri >= 0 Then
        OpKind = "R": v = ri: Exit Function
    End If

    v = ParseNumber(s)
    If v < 0 Then
        AddErr lineNo, "undefined label or bad value '" & s & "'"
        Exit Function
    End If
    ' Numbers/chars are immediates; labels resolved by ParseNumber are addresses
    If Labels.Exists(UCase$(Trim$(s))) Then
        OpKind = "A"
    Else
        OpKind = "I"
    End If
End Function

' "25", "0x1A", "'F'" or a label -> number. -1 if invalid / unknown label.
Public Function ParseNumber(ByVal s As String) As Long
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then ParseNumber = -1: Exit Function

    If Left$(t, 1) = "'" And Right$(t, 1) = "'" Then
        If Len(t) = 3 Then
            ParseNumber = Asc(Mid$(t, 2, 1))
        Else
            ParseNumber = -1
        End If
        Exit Function
    End If

    If LCase$(Left$(t, 2)) = "0x" Then
        On Error GoTo badHex
        ParseNumber = CLng("&H" & Mid$(t, 3))
        On Error GoTo 0
        Exit Function
    End If

    If IsNumeric(t) Then
        ParseNumber = CLng(t)
        Exit Function
    End If

    If Not Labels Is Nothing Then
        If Labels.Exists(UCase$(t)) Then
            ParseNumber = CLng(Labels(UCase$(t)))
            Exit Function
        End If
    End If

    ParseNumber = -1
    Exit Function
badHex:
    ParseNumber = -1
End Function

Private Sub AddErr(ByVal lineNo As Long, ByVal msg As String)
    AsmErrors.Add "Line " & lineNo & ": " & msg
End Sub

' Status strip + first error on Programs sheet (never crashes).
' Brand palette only: navy chip, cream strip, OK/ERROR word plus a
' green or red left edge so the state does not depend on color alone.
Private Sub PaintAsmStatus(ByVal ok As Boolean)
    Dim ws As Worksheet
    Dim banner As Range
    Dim detail As Range
    Dim tip As String
    Dim ink As Long
    Dim edge As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_PROG)
    If ws Is Nothing Then Exit Sub

    ws.Range("A1").Value = "ASSEMBLER"
    ws.Range("A1").Font.Name = "Segoe UI"
    ws.Range("A1").Font.Size = 12
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = CLR_TEXT
    ws.Range("A1").Interior.Color = CLR_BG
    ws.Range("A1").HorizontalAlignment = xlLeft
    ws.Range("A1").IndentLevel = 1
    If ws.Columns("A").ColumnWidth < 15 Then ws.Columns("A").ColumnWidth = 15

    Set banner = ws.Range("B2:F2")
    banner.Merge
    banner.Value = AsmStatusText
    banner.Font.Name = "Consolas"
    banner.Font.Size = 12
    banner.Font.Bold = True
    banner.HorizontalAlignment = xlLeft
    banner.VerticalAlignment = xlCenter
    ws.Rows(2).RowHeight = 22

    Set detail = ws.Range("B3:F3")
    detail.Merge
    detail.Font.Name = "Segoe UI"
    detail.Font.Size = 10
    detail.Font.Bold = False
    detail.HorizontalAlignment = xlLeft

    If ok Then
        ink = CLR_BG                 ' navy on cream 12.7:1
        edge = CLR_HIT
        tip = "Ready to load into memory"
    Else
        ink = CLR_FAULT              ' dark red on cream 6.2:1
        edge = CLR_MISS
        If Not AsmErrors Is Nothing Then
            If AsmErrors.Count > 0 Then
                tip = CStr(AsmErrors(1))
            Else
                tip = "Fix the source and assemble again"
            End If
        Else
            tip = "Fix the source and assemble again"
        End If
    End If
    detail.Value = tip

    banner.Interior.Color = CLR_TEXT
    detail.Interior.Color = CLR_TEXT
    banner.Font.Color = ink
    detail.Font.Color = ink
    With ws.Range("B2:B3").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = edge
    End With
    On Error GoTo 0
End Sub

' =====================================================================
' Test_Assembler - Immediate self-check for M4 / #27.
' =====================================================================
Public Function Test_Assembler() As Long
    Dim fails As Long
    Dim bytes As Variant
    Dim lines As Variant

    ' --- Program 1 (multiply) ---
    lines = Array( _
        "; Program 1", _
        "MOV AX, 0", _
        "MOV BX, 7", _
        "MOV CX, 5", _
        "mult: ADD AX, BX", _
        "DEC CX", _
        "JNZ mult", _
        "MOV DX, AX", _
        "STORE [0x080], DX", _
        "HLT")
    bytes = Assemble(lines)
    fails = fails + AsmCheck("Program 1 assembles", AsmErrors.Count = 0)
    fails = fails + AsmCheck("Program 1 byte count 18", Not IsEmpty(bytes) And UBound(bytes) - LBound(bytes) + 1 = 18)

    ' --- Program 2 (echo) ---
    lines = Array( _
        "MOV CX, 8", _
        "read: IN AX, [KBD]", _
        "CMP AX, 0", _
        "JZ read", _
        "OUT [LCD], AX", _
        "DEC CX", _
        "JNZ read", _
        "HLT")
    bytes = Assemble(lines)
    fails = fails + AsmCheck("Program 2 assembles", AsmErrors.Count = 0)

    ' --- Program 3 (scattered) ---
    lines = Array( _
        "MOV CX, 3", _
        "pass: LOAD AX, [0x020]", _
        "LOAD BX, [0x060]", _
        "ADD AX, BX", _
        "STORE [0x0A0], AX", _
        "LOAD DX, [0x0E0]", _
        "STORE [0x120], DX", _
        "LOAD BX, [0x160]", _
        "STORE [0x1A0], BX", _
        "LOAD AX, [0x1C0]", _
        "DEC CX", _
        "JNZ pass", _
        "HLT")
    bytes = Assemble(lines)
    fails = fails + AsmCheck("Program 3 assembles", AsmErrors.Count = 0)

    ' --- literals ---
    bytes = Assemble(Array("MOV AX, 'F'", "HLT"))
    fails = fails + AsmCheck("char literal 'F'", AsmErrors.Count = 0 And bytes(1) = Asc("F"))

    ' --- error: unknown opcode ---
    bytes = Assemble(Array("FOO AX, 1"))
    fails = fails + AsmCheck("unknown opcode errors", AsmErrors.Count > 0)

    ' --- error: bad register ---
    bytes = Assemble(Array("MOV ZZ, 1"))
    fails = fails + AsmCheck("invalid register errors", AsmErrors.Count > 0)

    ' --- error: undefined label ---
    bytes = Assemble(Array("JMP nowhere"))
    fails = fails + AsmCheck("undefined label errors", AsmErrors.Count > 0)

    ' --- error: wrong operands ---
    bytes = Assemble(Array("ADD AX"))
    fails = fails + AsmCheck("wrong operands errors", AsmErrors.Count > 0)

    ' --- error: imm out of range ---
    bytes = Assemble(Array("MOV AX, 300"))
    fails = fails + AsmCheck("imm out of range errors", AsmErrors.Count > 0)

    Debug.Print IIf(fails = 0, "Test_Assembler: ALL OK", "Test_Assembler: " & fails & " FAIL(S)")
    Test_Assembler = fails
End Function

Private Function AsmCheck(ByVal what As String, ByVal passed As Boolean) As Long
    Debug.Print IIf(passed, "  OK    ", "  FAIL  ") & what
    AsmCheck = IIf(passed, 0, 1)
End Function
