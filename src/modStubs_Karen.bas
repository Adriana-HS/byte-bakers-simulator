Attribute VB_Name = "modStubs_Karen"
Option Explicit
' =====================================================================
' modStubs_Karen - Empty signatures of my functions not implemented yet.
' RULE: when a function is implemented, its stub is DELETED in the SAME
' PR (otherwise VBA raises "Ambiguous name detected").
' =====================================================================

' --- modCache (K7, K8) ---
Public Sub InvalidateAllCaches(): End Sub
Public Function CacheReadL1D(ByVal paddr As Integer) As Integer: CacheReadL1D = RAM(paddr): End Function
Public Sub CacheWriteL1D(ByVal paddr As Integer, ByVal value As Integer): RAM(paddr) = value: End Sub
Public Sub InvalidateFrame(ByVal pfn As Integer): End Sub

' --- Statistics (K9) ---
Public Sub RefreshStats(): End Sub

' --- modPanelMemory (K10) ---
Public Sub ShowTranslation(ByVal vpn As Integer, ByVal offset As Integer): End Sub
Public Sub UI_PageFaultFlash(ByVal vpn As Integer): End Sub
Public Sub UI_AnimateSwapIn(ByVal vpn As Integer, ByVal pfn As Integer): End Sub
Public Sub UI_AnimateLineFill(ByVal idx As Integer): End Sub
Public Sub RefreshMemoryPanel(): End Sub
