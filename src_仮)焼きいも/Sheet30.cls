VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Target.Row = 4 Then Exit Sub
    If Not Target.Column = 7 Then Exit Sub
    Cancel = True
    Me.Cells(Target.Row, Target.Column) = OpenVbaFormAllSheets()
End Sub

Private Sub cmdCopy87_Click()
    Call CopyOrClearRow(Me, 87, True)
End Sub

Private Sub cmdCopy104_Click()
    Call CopyOrClearRow(Me, 104, True)
End Sub

Private Sub cmdClear87_Click()
    Call CopyOrClearRow(Me, 87, False)
End Sub

Private Sub cmdClear104_Click()
    Call CopyOrClearRow(Me, 104, False)
End Sub

'焼減率の計算（I108セルに0.25刻みで四捨五入した値をセット：計算式はThisWorkbook.SetI108WithRounding）
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHanlder
    If Not Intersect(Target, Range("I106,I107")) Is Nothing Then
        ThisWorkbook.SetI108WithRounding Me
    End If
    Exit Sub
ErrorHanlder:
    Debug.Print "[" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "] [Error] Sheet29.Worksheet_Change: " & Err.Number & " - " & Err.Description
    Call MsgBox("焼減率の計算に失敗しました。システム管理者へ連絡してください｡ ", vbCritical)
End Sub
