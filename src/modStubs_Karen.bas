Attribute VB_Name = "modStubs_Karen"
Option Explicit
' =====================================================================
' modStubs_Karen - Firmas vacias de mis funciones aun no implementadas.
' REGLA: al implementar una funcion, se BORRA su stub en el MISMO PR
' (si no, VBA da "nombre ambiguo").
' =====================================================================

' --- modMMU (K5) ---
Public Sub HandlePageFault(ByVal vpn As Integer): End Sub

' --- modCache (K6, K7, K8) ---
Public Sub InvalidateAllCaches(): End Sub
Public Function CacheReadL1I(ByVal paddr As Integer) As Integer: CacheReadL1I = RAM(paddr): End Function
Public Function CacheReadL1D(ByVal paddr As Integer) As Integer: CacheReadL1D = RAM(paddr): End Function
Public Sub CacheWriteL1D(ByVal paddr As Integer, ByVal value As Integer): RAM(paddr) = value: End Sub
Public Sub InvalidateFrame(ByVal pfn As Integer): End Sub

' --- Estadisticas (K9) ---
Public Sub RefreshStats(): End Sub

' --- modPanelMemory (K10) ---
Public Sub ShowTranslation(ByVal vpn As Integer, ByVal offset As Integer): End Sub
Public Sub UI_PageFaultFlash(ByVal vpn As Integer): End Sub
Public Sub UI_AnimateSwapIn(ByVal vpn As Integer, ByVal pfn As Integer): End Sub
Public Sub UI_AnimateLineFill(ByVal idx As Integer): End Sub
Public Sub RefreshMemoryPanel(): End Sub
