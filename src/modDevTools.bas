Attribute VB_Name = "modDevTools"
Option Explicit
' =====================================================================
' modDevTools - Herramientas de desarrollo (K1)
' Exporta/importa los modulos VBA a la carpeta src\ del repositorio para
' poder versionarlos con Git. Requiere abrir el libro desde la carpeta
' excel\ del repo y activar "Confiar en el acceso al modelo de objetos
' de proyectos de VBA" (Centro de confianza > Configuracion de macros).
' =====================================================================

Private Function SrcPath() As String
    SrcPath = ThisWorkbook.Path & "\..\src\"
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

' ---------- EXPORTAR ----------
' Exporta los modulos a destPath (por defecto, src\ del repo).
' Devuelve cuantos modulos exporto; -1 si hubo error
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
            Case 1: ext = ".bas"      ' modulo estandar
            Case 2: ext = ".cls"      ' modulo de clase
            Case Else: ext = ""       ' hojas y ThisWorkbook: no se exportan
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
        MsgBox "Error al exportar: " & Err.Description, vbCritical, "Byte Bakers"
    End If
    ExportModulesCore = -1
End Function

Public Sub ExportModules()
    Dim n As Integer
    n = ExportModulesCore()
    If n >= 0 Then MsgBox n & " modulos exportados a " & SrcPath(), vbInformation, "Byte Bakers"
End Sub

' ---------- IMPORTAR ----------
' Devuelve cuantos modulos importo (sin contar modDevTools); -1 si hubo error
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
    ' 1) primero se lista (Dir no se puede anidar con otras llamadas a Dir)
    f = Dir(SrcPath() & "*.bas")
    Do While f <> ""
        files.Add f
        f = Dir()
    Loop
    ' 2) luego se reemplaza cada modulo (menos modDevTools, que esta en ejecucion)
    For Each item In files
        nm = Left(item, Len(item) - 4)
        If nm <> "modDevTools" Then
            Set old = Nothing
            On Error Resume Next
            Set old = vbp.VBComponents(nm)
            On Error GoTo EH
            If Not old Is Nothing Then
                old.Name = "zz_old_" & nm   ' evita que el importado quede como "modConfig1"
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
        MsgBox "Error al importar: " & Err.Description, vbCritical, "Byte Bakers"
    End If
    ImportModulesCore = -1
End Function

Public Sub ImportModules()
    Dim n As Integer
    n = ImportModulesCore()
    If n >= 0 Then MsgBox n & " modulos importados. Guarda el libro (Ctrl+S).", vbInformation, "Byte Bakers"
End Sub

' ---------- PRUEBA ----------
' Ejecutar desde el editor VBA (F5 con el cursor dentro) y mirar la ventana Inmediato (Ctrl+G).
' NO escribe en src\: importa desde src\ y exporta solo a una carpeta temporal.
Public Sub Test_DevTools()
    Dim fails As Integer, nBefore As Integer, nAfter As Integer
    Dim comp As Object, nExp As Integer, nImp As Integer, tmp As String

    ' 1) Sincronizar: trae a este libro los modulos de src\
    If ImportModulesCore() < 0 Then
        fails = fails + 1: Debug.Print "  FAIL  la sincronizacion inicial dio error"
    End If

    ' 2) Ida y vuelta: importar de nuevo no debe cambiar la cantidad de componentes
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

    ' 3) Exportar a una carpeta temporal (para no tocar src\)
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

    ' 4) Limpieza de la carpeta temporal
    On Error Resume Next
    Kill tmp & "*.bas"
    Kill tmp & "*.cls"
    RmDir tmp
    On Error GoTo 0

    Debug.Print "Importados: " & nImp & " | Exportados (carpeta temporal): " & nExp
    Debug.Print IIf(fails = 0, "Test_DevTools: ALL OK", "Test_DevTools: " & fails & " FAIL(S)")
End Sub

