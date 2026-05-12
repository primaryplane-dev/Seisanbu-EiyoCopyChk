VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet29"
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
    Call subCopyClear(87, True)
End Sub

Private Sub cmdCopy104_Click()
    Call subCopyClear(104, True)
End Sub

Private Sub cmdClear87_Click()
    Call subCopyClear(87, False)
End Sub

Private Sub cmdClear104_Click()
    Call subCopyClear(104, False)
End Sub

' I106とI107の値が変更されたときにI108を更新する
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("I106,I107")) Is Nothing Then
        Dim val As Double
        Dim roundedVal As Double

        ' 計算式
        val = ((Range("I107").Value / Range("I106").Value) * 100 - 100) * (-1)

        ' 0.25刻みで四捨五入
        roundedVal = WorksheetFunction.Floor(val * 2, 1) / 2
        If val - roundedVal >= 0.25 Then
            roundedVal = roundedVal + 0.5
        End If

        Application.EnableEvents = False
        Range("I108").Value = roundedVal
        Application.EnableEvents = True
    End If
End Sub
