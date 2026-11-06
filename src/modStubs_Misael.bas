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

' ---------------------------------------------------------------------
' TEMP compile bridges for M3 until A8 / A9 / K10 publish their stubs.
' Must stay ABOVE all Subs (VBA forbids module vars after procedures).
' When those modules land, DELETE these lines (same pattern as the old
' UI_Highlight stub that lived here until M2).
' ---------------------------------------------------------------------
Public MicroStep As Integer   ' A8
Public RunFlag As Boolean     ' A8

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

' --- A8 / A9 / K10 bridges (delete when real modules land) ---
Public Sub StepMicro()
    If MicroStep < 1 Then MicroStep = 1 Else MicroStep = MicroStep + 1
End Sub

Public Sub RunProgram()
    RunFlag = True
End Sub

Public Sub BuildCPUPanel(ByVal ws As Worksheet)
End Sub

Public Sub RefreshCPUPanel()
End Sub

Public Sub BuildMiniMemory(ByVal ws As Worksheet)
End Sub
' RefreshMemoryPanel / RefreshStats already exist in modStubs_Karen.
