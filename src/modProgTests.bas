Attribute VB_Name = "modProgTests"
Option Explicit
' =====================================================================
' modProgTests - Checks that the catalog programs assemble correctly.
' Reads the real .asm file from programs\ (the workbook lives in excel\).
' =====================================================================

Public Sub Test_Program3()
    Dim fails As Integer, i As Long, n As Long, f As Integer
    Dim path As String, ln As String
    Dim src() As String
    Dim code As Variant, nBytes As Long

    path = ThisWorkbook.path & "\..\programs\program3_page_fault.asm"
    If Dir(path) = "" Then
        Debug.Print "  FAIL  1: file not found: " & path
        Debug.Print "Test_Program3: 1 FAIL(S)"
        Exit Sub
    End If

    ' 1) Read every line of the file
    f = FreeFile
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ReDim Preserve src(0 To n)
        src(n) = ln
        n = n + 1
    Loop
    Close #f
    If n = 0 Then
        Debug.Print "  FAIL  1: the file is empty"
        Debug.Print "Test_Program3: 1 FAIL(S)"
        Exit Sub
    End If

    ' 2) Assemble: no syntax errors allowed
    code = Assemble(src)
    If AsmErrors.Count > 0 Then
        fails = fails + 1
        Debug.Print "  FAIL  2: the assembler reported " & AsmErrors.Count & " error(s):"
        For i = 1 To AsmErrors.Count
            Debug.Print "    " & AsmErrors(i)
        Next i
    End If

    ' 3) Size: 13 instructions x 2 bytes = 26 bytes, all in page 0 (< 32 bytes)
    If Not IsArray(code) Then
        fails = fails + 1
        Debug.Print "  FAIL  3: Assemble did not return an array"
    Else
        nBytes = UBound(code) - LBound(code) + 1
        If nBytes <> 26 Then fails = fails + 1: Debug.Print "  FAIL  3: expected 26 bytes, got " & nBytes
        For i = LBound(code) To UBound(code)
            If code(i) < 0 Or code(i) > BYTE_MAX Then fails = fails + 1: Debug.Print "  FAIL  3: byte " & i & " out of range: " & code(i)
        Next i
    End If

    Debug.Print IIf(fails = 0, "Test_Program3: ALL OK (" & nBytes & " bytes)", "Test_Program3: " & fails & " FAIL(S)")
End Sub
