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

' --- modLCD (#29 / M6). RefreshLCD lives in modLCD. ---
' --- modIO (#30 / M7). IoRead and IoWrite live in modIO. ---
' --- K10. A8 StepMicro/RunProgram/MicroStep live in modControlUnit. ---
' --- A9 BuildCPUPanel/RefreshCPUPanel live in modPanelCPU. ---
' RefreshMemoryPanel / RefreshStats already exist in modStubs_Karen.
