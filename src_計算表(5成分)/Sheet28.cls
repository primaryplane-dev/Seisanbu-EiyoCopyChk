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
