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
