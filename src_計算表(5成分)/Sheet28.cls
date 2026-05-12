VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet28"
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
'------------------------------------------------------
'  ■■■■■■ 出力なし誤差シート作成    ■■■■■■
'------------------------------------------------------
Private Sub cmdCopy_Click()
    If Not MsgBox("誤差シートを作成しますか？", vbQuestion Or vbYesNo) = vbYes Then Exit Sub
    Call subCopyAndEdit
End Sub
Private Sub subCopyAndEdit()
    Dim ST          As Worksheet
    Me.Copy After:=Me
    Set ST = ThisWorkbook.Sheets(Me.Index + 1)
    ST.Range(ST.Cells(2, 27), ST.Cells(3, 29)).Merge
    ST.Range(ST.Cells(2, 27), ST.Cells(3, 29)).Interior.Color = RGB(255, 204, 0)
    'ST.Range(ST.Cells(2, 27), ST.Cells(3, 29)).Font.Color = RGB(255, 255, 255)
    ST.Range(ST.Cells(2, 27), ST.Cells(3, 29)).VerticalAlignment = xlCenter
    ST.Range(ST.Cells(2, 27), ST.Cells(3, 29)).HorizontalAlignment = xlCenter
    ST.Cells(2, 27) = "誤差範囲確認用"
    
    ST.Columns(38).Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromRightOrBelow
    ST.Columns(38).Insert Shift:=xlToLeft, CopyOrigin:=xlFormatFromRightOrBelow
    ST.Columns(38).NumberFormatLocal = "G/標準"
    ST.Columns(39).NumberFormatLocal = "G/標準"
    ST.Columns(38).HorizontalAlignment = xlCenter
    ST.Cells(13, 38) = "±20％範囲"
    ST.Range(ST.Cells(14, 38), ST.Cells(25, 38)).Formula = "=fncHantei('" & Me.Name & "'!RC[-5],RC[-5])"
    ST.Range(ST.Cells(14, 39), ST.Cells(25, 39)).Formula = "=fncCalc('" & Me.Name & "'!RC[-6],RC[-6])"
    
    ST.Columns(46).NumberFormatLocal = "G/標準"
    ST.Columns(47).NumberFormatLocal = "G/標準"
    ST.Columns(46).HorizontalAlignment = xlCenter
    ST.Cells(13, 46) = "±20％範囲"
    ST.Range(ST.Cells(17, 46), ST.Cells(19, 46)).Formula = "=fncHantei('" & Me.Name & "'!RC[-4],RC[-2])"
    ST.Range(ST.Cells(17, 47), ST.Cells(19, 47)).Formula = "=fncCalc('" & Me.Name & "'!RC[-5],RC[-3])"
End Sub
