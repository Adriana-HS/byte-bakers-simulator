Attribute VB_Name = "modStubs_Misael"
' =====================================================================
' modStubs_Misael - TEMPORARY empty signatures (not in the release)
' Owner: Misael  |  Expanded by task #25 (M2)
'
' Empty UI_* signatures so Adriana's buses (A6) and Karen's panels
' compile before the animation engine lands. Each stub is deleted in
' the same PR that implements it (ambiguous-name rule).
' =====================================================================
Option Explicit

' --- modUI (#25) ---
Public Sub UI_Highlight(ByVal nm As String, ByVal clr As Long): End Sub
