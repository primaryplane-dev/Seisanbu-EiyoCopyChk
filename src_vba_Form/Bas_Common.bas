Attribute VB_Name = "Bas_Common"
Option Explicit

Sub Auto_Close()
    ThisWorkbook.Saved = True   'ïœçXÇï€ë∂ÇµÇ»Ç¢
End Sub

Public Sub subBeforeEdit()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Public Sub subAfterEdit()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
