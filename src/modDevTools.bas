Attribute VB_Name = "modDevTools"
Option Explicit
' =====================================================================
' modDevTools - Development tools (K1)
' Exports/imports the VBA modules to the src\ folder of the repository so
' they can be versioned with Git. Requires opening the workbook from the
' excel\ folder of the repo and enabling "Trust access to the VBA project
' object model" (Trust Center > Macro Settings).
' =====================================================================

Private Function SrcPath() As String
    SrcPath = ThisWorkbook.path & "\..\src\"
End Function

Private Function SrcExists() As Boolean
    SrcExists = (Dir(SrcPath(), vbDirectory) <> "")
End Function

Private Sub ShowTrustError()
    MsgBox "Excel no permite acceder al proyecto VBA." & vbCrLf & vbCrLf & _
           "Activa: Archivo > Opciones > Centro de confianza > Configuracion del " & _
           "Centro de confianza > Configuracion de macros > " & _
           "'Confiar en el acceso al modelo de objetos de proyectos de VBA'.", _
           vbExclamation, "Byte Bakers"
End Sub

' ---------- EXPORT ----------
' Exports the modules to destPath (by default, src\ of the repo).
' Returns how many modules it exported; -1 if there was an error
Private Function ExportModulesCore(Optional ByVal destPath As String = "") As Integer
    Dim comp As Object, ext As String, n As Integer, dest As String
    If destPath = "" Then
        If Not SrcExists() Then
            MsgBox "No existe la carpeta src. Abre el libro desde la carpeta excel\ del repositorio.", _
                   vbExclamation, "Byte Bakers"
            ExportModulesCore = -1
            Exit Function
        End If
        dest = SrcPath()
    Else
        dest = destPath
    End If
    On Error GoTo EH
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1: ext = ".bas"      ' standard module
            Case 2: ext = ".cls"      ' class module
            Case Else: ext = ""       ' sheets and ThisWorkbook: not exported
        End Select
        If ext <> "" Then
            comp.Export dest & comp.Name & ext
            n = n + 1
        End If
    Next comp
    ExportModulesCore = n
    Exit Function
EH:
    If Err.Number = 1004 Then
        ShowTrustError
    Else
        MsgBox "Error al exportar: " & Err.description, vbCritical, "Byte Bakers"
    End If
    ExportModulesCore = -1
End Function

Public Sub ExportModules()
    Dim n As Integer
    n = ExportModulesCore()
    If n >= 0 Then MsgBox n & " modulos exportados a " & SrcPath(), vbInformation, "Byte Bakers"
End Sub

' ---------- IMPORT ----------
' Returns how many modules it imported (not counting modDevTools); -1 if there was an error
Private Function ImportModulesCore() As Integer
    Dim vbp As Object, old As Object, f As String, nm As String
    Dim files As New Collection, item As Variant, cnt As Integer
    If Not SrcExists() Then
        MsgBox "No existe la carpeta src. Abre el libro desde la carpeta excel\ del repositorio.", _
               vbExclamation, "Byte Bakers"
        ImportModulesCore = -1
        Exit Function
    End If
    On Error GoTo EH
    Set vbp = ThisWorkbook.VBProject
    ' 1) first list the files (Dir cannot be nested with other calls to Dir)
    f = Dir(SrcPath() & "*.bas")
    Do While f <> ""
        files.Add f
        f = Dir()
    Loop
    ' 2) then replace each module (except modDevTools, which is running)
    For Each item In files
        nm = Left(item, Len(item) - 4)
        If nm <> "modDevTools" Then
            Set old = Nothing
            On Error Resume Next
            Set old = vbp.VBComponents(nm)
            On Error GoTo EH
            If Not old Is Nothing Then
                old.Name = "zz_old_" & nm   ' keeps the imported one from being named "modConfig1"
                vbp.VBComponents.Remove old
            End If
            vbp.VBComponents.Import SrcPath() & item
            cnt = cnt + 1
        End If
    Next item
    ImportModulesCore = cnt
    Exit Function
EH:
    If Err.Number = 1004 Then
        ShowTrustError
    Else
        MsgBox "Error al importar: " & Err.description, vbCritical, "Byte Bakers"
    End If
    ImportModulesCore = -1
End Function

Public Sub ImportModules()
    Dim n As Integer
    n = ImportModulesCore()
    If n >= 0 Then MsgBox n & " modulos importados. Guarda el libro (Ctrl+S).", vbInformation, "Byte Bakers"
End Sub

' ---------- TEST ----------
' Run from the VBA editor (F5 with the cursor inside) and watch the Immediate window (Ctrl+G).
' Does NOT write to src\: imports from src\ and exports only to a temporary folder.
Public Sub Test_DevTools()
    Dim fails As Integer, nBefore As Integer, nAfter As Integer
    Dim comp As Object, nExp As Integer, nImp As Integer, tmp As String

    ' 1) Sync: brings the modules from src\ into this workbook
    If ImportModulesCore() < 0 Then
        fails = fails + 1: Debug.Print "  FAIL  la sincronizacion inicial dio error"
    End If

    ' 2) Round trip: importing again must not change the number of components
    nBefore = ThisWorkbook.VBProject.VBComponents.Count
    nImp = ImportModulesCore()
    nAfter = ThisWorkbook.VBProject.VBComponents.Count
    If nImp < 0 Then fails = fails + 1: Debug.Print "  FAIL  ImportModules dio error"
    If nBefore <> nAfter Then
        fails = fails + 1: Debug.Print "  FAIL  cambio la cantidad de componentes: " & nBefore & " -> " & nAfter
    End If
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If Left(comp.Name, 7) = "zz_old_" Then
            fails = fails + 1: Debug.Print "  FAIL  quedo un modulo viejo: " & comp.Name
        End If
    Next comp

    ' 3) Export to a temporary folder (so src\ is not touched)
    tmp = Environ("TEMP") & "\bb_devtools_test\"
    On Error Resume Next
    If Dir(tmp, vbDirectory) = "" Then MkDir tmp
    On Error GoTo 0
    nExp = ExportModulesCore(tmp)
    If nExp < 1 Then
        fails = fails + 1: Debug.Print "  FAIL  ExportModules no exporto nada"
    End If
    If Dir(tmp & "modDevTools.bas") = "" Then
        fails = fails + 1: Debug.Print "  FAIL  no se exporto modDevTools.bas"
    End If

    ' 4) Cleanup of the temporary folder
    On Error Resume Next
    Kill tmp & "*.bas"
    Kill tmp & "*.cls"
    RmDir tmp
    On Error GoTo 0

    Debug.Print "Importados: " & nImp & " | Exportados (carpeta temporal): " & nExp
    Debug.Print IIf(fails = 0, "Test_DevTools: ALL OK", "Test_DevTools: " & fails & " FAIL(S)")
End Sub

