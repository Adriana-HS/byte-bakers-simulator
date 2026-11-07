Attribute VB_Name = "modLCD"
' =====================================================================
' modLCD - 8x1 video buffer latch (G4)
' Owner: Misael  |  Task: #29 (M6)
'
' The eight cells always show RAM(&HF8)-RAM(&HFF). The cursor lives at
' &HE2. LcdWrite is the latch: it stores a byte and advances. 0x08 steps
' back and clears that cell. The glyphs are CLR_LCD on navy, monospace.
' =====================================================================
Option Explicit

' Reserved column in frm_IO (same rectangle modKeyboard used to hold).
Private Const LCD_L As Single = 466
Private Const LCD_T As Single = 448
Private Const LCD_W As Single = 182
Private Const LCD_H As Single = 62
Private Const CELL_W As Single = 18
Private Const CELL_H As Single = 32
Private Const CELL_GAP As Single = 3

Private mLast As Integer

Public Sub LcdInit()
    Dim i As Integer
    For i = 0 To IO_BUF_SIZE - 1
        RAM(LCD_BASE + i) = 0
    Next i
    RAM(LCD_CURSOR) = 0
    mLast = -1
End Sub

' Latch one byte at the cursor. 0x08 clears the previous cell.
Public Sub LcdWrite(ByVal ascii As Integer)
    Dim cur As Integer
    ascii = ascii And &HFF&
    cur = RAM(LCD_CURSOR) And &HFF&
    If cur < 0 Or cur > 7 Then cur = 0
    If ascii = 8 Then
        If cur > 0 Then cur = cur - 1
        RAM(LCD_BASE + cur) = 0
        RAM(LCD_CURSOR) = cur
    Else
        RAM(LCD_BASE + cur) = ascii
        RAM(LCD_CURSOR) = (cur + 1) Mod IO_BUF_SIZE
    End If
    mLast = cur
    RefreshLCD
    BlinkCell cur
End Sub

Public Sub ClearLCD()
    LcdInit
    RefreshLCD
End Sub

Public Sub LcdBuild(ByVal ws As Worksheet)
    Dim glass As Shape
    Dim i As Integer
    Dim x As Single
    Dim y As Single
    On Error Resume Next
    ws.Shapes("lcd_placeholder").Delete
    ws.Shapes("lcd_glass").Delete
    For i = 0 To IO_BUF_SIZE - 1
        ws.Shapes("lcd_c" & i).Delete
    Next i
    On Error GoTo 0
    Set glass = ws.Shapes.AddShape(msoShapeRoundedRectangle, LCD_L, LCD_T, LCD_W, LCD_H)
    glass.Name = "lcd_glass"
    glass.Fill.ForeColor.RGB = CLR_BG
    glass.Line.Visible = msoTrue
    glass.Line.ForeColor.RGB = CLR_LCD
    glass.Line.Weight = 1.25
    glass.Adjustments.Item(1) = 0.08
    y = LCD_T + (LCD_H - CELL_H) / 2
    x = LCD_L + 8
    For i = 0 To IO_BUF_SIZE - 1
        Set glass = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, CELL_W, CELL_H)
        glass.Name = "lcd_c" & i
        glass.Fill.ForeColor.RGB = CLR_BG
        glass.Line.Visible = msoFalse
        With glass.TextFrame2
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Name = "Consolas"
            .TextRange.Font.Size = 12
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = CLR_LCD
        End With
        x = x + CELL_W + CELL_GAP
    Next i
    If RAM(LCD_CURSOR) < 0 Or RAM(LCD_CURSOR) > 7 Then LcdInit
    RefreshLCD
End Sub

Public Sub RefreshLCD()
    Dim ws As Worksheet
    Dim s As Shape
    Dim i As Integer
    Dim cur As Integer
    Dim v As Integer
    Dim shown As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SIM)
    If ws Is Nothing Then Exit Sub
    Set s = ws.Shapes("lcd_c0")
    If s Is Nothing Then Exit Sub
    On Error GoTo 0
    cur = RAM(LCD_CURSOR) And &HFF&
    If cur < 0 Or cur > 7 Then cur = 0
    For i = 0 To IO_BUF_SIZE - 1
        Set s = ws.Shapes("lcd_c" & i)
        v = RAM(LCD_BASE + i) And &HFF&
        If v >= 32 And v < 127 Then
            shown = Chr$(v)
        Else
            shown = ""
        End If
        With s.TextFrame2.TextRange
            .Text = shown
            .Font.Name = "Consolas"
            .Font.Size = 12
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = CLR_LCD
        End With
        s.Fill.ForeColor.RGB = CLR_BG
        If i = cur Then
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = CLR_LCD
            s.Line.Weight = 1.5
        Else
            s.Line.Visible = msoFalse
        End If
    Next i
End Sub

' The cell that was just latched inverts for one animation beat.
Private Sub BlinkCell(ByVal i As Integer)
    Dim s As Shape
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(SH_SIM).Shapes("lcd_c" & i)
    If s Is Nothing Then Exit Sub
    s.Fill.ForeColor.RGB = CLR_LCD
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_BG
    On Error GoTo 0
    UI_Wait
    RefreshLCD
End Sub

