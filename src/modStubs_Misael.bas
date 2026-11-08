Attribute VB_Name = "modStubs_Misael"
' =====================================================================
' modStubs_Misael - TEMPORARY empty signatures (not in the release)
' Owner: Misael
'
' Rule: when a function is implemented, its stub is deleted in the SAME
' pull request; otherwise VBA raises "Ambiguous name detected".
'
' RefreshAll lives in modDashboard (#26). UI_* live in modUI (#25).
' =====================================================================
Option Explicit

' --- modLCD (#29 / M6). BuildIOPanel + RefreshKeyboard live in modKeyboard (#28). ---
Public Sub RefreshLCD()
End Sub

' --- modIO (#30 / M7) ---
Public Function IoRead(ByVal destReg As Integer) As Integer
    IoRead = 0
End Function

Public Sub IoWrite(ByVal v As Integer)
End Sub

' --- K10. A8 StepMicro/RunProgram/MicroStep live in modControlUnit. ---
' --- A9 BuildCPUPanel/RefreshCPUPanel live in modPanelCPU. ---
Public Sub BuildMiniMemory(ByVal ws As Worksheet)
End Sub
' RefreshMemoryPanel / RefreshStats already exist in modStubs_Karen.
