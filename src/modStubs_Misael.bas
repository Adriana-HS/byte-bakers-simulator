Attribute VB_Name = "modStubs_Misael"
' =====================================================================
' modStubs_Misael - TEMPORARY empty signatures (not in the release)
' Owner: Misael  |  Task: #25 (M2) and later M3/M7
'
' Public functions of Misael's modules that other members call before
' they are implemented, so everybody's code compiles from day one.
' Rule: when a function is implemented, its stub is deleted in the SAME
' pull request; otherwise VBA raises "Ambiguous name detected".
'
' UI_* signatures for M2 live in modUI (this PR). Stubs left here are
' for modules still pending (M3 dashboard, M7 I/O).
' =====================================================================
Option Explicit

' --- modDashboard (#26 / M3) ---
Public Sub RefreshAll(): End Sub

' --- modIO (#30 / M7) ---
Public Function IoRead(ByVal destReg As Integer) As Integer: IoRead = 0: End Function
Public Sub IoWrite(ByVal v As Integer): End Sub
