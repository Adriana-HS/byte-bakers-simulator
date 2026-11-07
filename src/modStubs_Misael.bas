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

' --- modKeyboard / modLCD (#28/#29 / M5+M6) ---
Public Sub BuildIOPanel(ByVal ws As Worksheet)
End Sub

Public Sub RefreshKeyboard()
End Sub

Public Sub RefreshLCD()
End Sub

' --- modIO (#30 / M7) ---
Public Function IoRead(ByVal destReg As Integer) As Integer
    IoRead = 0
End Function

Public Sub IoWrite(ByVal v As Integer)
End Sub

' --- A9 / K10 bridges (A8 StepMicro/RunProgram/MicroStep → modControlUnit) ---
Public Sub BuildCPUPanel(ByVal ws As Worksheet)
End Sub

Public Sub RefreshCPUPanel()
End Sub

Public Sub BuildMiniMemory(ByVal ws As Worksheet)
End Sub
' RefreshMemoryPanel / RefreshStats already exist in modStubs_Karen.
