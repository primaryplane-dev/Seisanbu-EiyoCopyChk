VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    CheckCopyButtonAllSheets
End Sub

Sub SetI108WithRounding(ws As Worksheet)
    On Error GoTo ErrorHandler
    Dim Val As Double
    Dim RoundedVal As Double

    ' 0除算チェック
    If CDbl(ws.Range("I106").Value) = 0 Then Exit Sub

    ' 計算式の値を取得
    Val = ((ws.Range("I107").Value / ws.Range("I106").Value) * 100 - 100) * -1

    ' 0.25刻みで四捨五入
    RoundedVal = WorksheetFunction.Floor(Val * 2, 1) / 2
    If Val - RoundedVal >= 0.25 Then
        RoundedVal = RoundedVal + 0.5
    End If

    ws.Range("I108").Value = RoundedVal
    Exit Sub
ErrorHandler:
    Debug.Print "[" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "] [Error] ThisWorkbook.SetI108WithRounding: " & Err.Number & " - " & Err.Description
    Call MsgBox("焼減率の計算に失敗しました。システム管理者へ連絡してください｡ ", vbCritical)
End Sub