Attribute VB_Name = "modStubs_Adriana"
' =====================================================================
' modStubs_Adriana - TEMPORARY empty signatures (not in the release)
' Owner: Adriana  |  Task: #2 (A2)
'
' Public functions of Adriana's modules that other members call before
' they are implemented, so everybody's code compiles from day one.
' Rule: when a function is implemented, its stub is deleted in the SAME
' pull request; otherwise VBA raises "Ambiguous name detected".
' =====================================================================
Option Explicit

' --- modCPU (#4) ---
Public Sub ResetCPU(): End Sub
Public Function RegName(ByVal i As Integer) As String: RegName = Choose(i + 1, "AX", "BX", "CX", "DX"): End Function

' --- modBus (#6) ---
Public Sub ReleaseBuses(): End Sub
Public Sub DriveDataBus(ByVal source As String, ByVal value As Integer, ByVal isInternal As Boolean): End Sub
