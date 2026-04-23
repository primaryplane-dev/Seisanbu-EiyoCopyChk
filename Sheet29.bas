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
Private Const cnFile = "\\srv0103\ALL\SEISAN\栄養計算\新表示用　計算表\製品登録申請用\栄養計算フォーマット用分類一覧\vba_Form.xlsm"

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Target.Row = 4 Then Exit Sub
    If Not Target.Column = 7 Then Exit Sub
    Cancel = True
    Me.Cells(Target.Row, Target.Column) = fncOpen
End Sub

Private Function fncOpen() As String
    Dim oWB As Workbook
    Set oWB = Workbooks.Open(cnFile, False, True)
    Me.Copy oWB.Sheets(1)
    Application.Run "'" & cnFile & "'!Bas_Main.subMain"
    fncOpen = Trim(oWB.Sheets("VBA").Cells(1, 1))
    oWB.Close False
    Set oWB = Nothing
End Function

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

Private Sub subCopyClear(ByVal lRow As Long, ByVal bMode As Boolean)
    Dim lCol    As Long
    For lCol = 3 To 12
        If Not lCol = 6 Then
            If bMode Then Me.Cells(lRow, lCol) = Me.Cells(lRow - 1, lCol)
            If Not bMode Then Me.Cells(lRow, lCol) = ""
        End If
    Next
End Sub
