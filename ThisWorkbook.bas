Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Dim ws As Worksheet
  Dim checkRows As Variant

  checkRows = Array(87, 104)
  Dim r As Variant
  Dim c As Long
  Dim isMiss As Boolean
  Dim msg As String

'すべてのシートの指定行をチェック
  For Each ws In ThisWorkbook.Worksheets
    For Each r In checkRows
      Dim missInRow As Boolean
      For c = 3 To 12
        If c <> 6 Then
          If CStr(ws.Cells(r, c).Value) <> CStr(ws.Cells(r - 1, c).Value) Then
            ws.Cells(r, c).Interior.Color = RGB(255, 199, 206) ' 薄いピンク（Excel標準）
            isMiss = True
            missInRow = True
          Else
            ws.Cells(r, c).Interior.ColorIndex = xlColorIndexNone ' 色を戻す
          End If
        End If
      Next c
      If missInRow Then
        msg = msg & ws.Name & "：行" & r & vbCrLf
      End If
    Next r
  Next ws

  If isMiss Then
    MsgBox "コピーボタンの押し忘れがある可能性があります。"  & vbCrLf & "薄いピンクのセルを確認してください。", vbExclamation
  End If
End Sub

