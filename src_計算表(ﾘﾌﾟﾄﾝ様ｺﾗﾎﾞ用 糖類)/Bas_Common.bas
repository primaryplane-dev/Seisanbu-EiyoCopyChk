Option Explicit

'=============================
' 共通定数
'=============================
Public Const cnFile As String = "\\srv0103\ALL\SEISAN\栄養計算\新表示用　計算表\製品登録申請用\栄養計算フォーマット用分類一覧\vba_Form.xlsm"

'=============================
' コピー・クリア共通処理
'=============================
'--- 既存ボタン呼び出し用のエイリアス（後方互換） ---
Public Sub subCopyClear(lRow As Long, bMode As Boolean)
    CopyOrClearRow ActiveSheet, lRow, bMode
End Sub

'--- コピー・クリア共通処理 ---
Public Sub CopyOrClearRow(ws As Worksheet, lRow As Long, bMode As Boolean)
    Dim lCol As Long
    Dim lastCol As Long

    ' 最終列を取得（ただし14列目までは処理する）
    lastCol = ws.Cells(lRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 14 Then lastCol = 14
    For lCol = 3 To lastCol
        If lCol <> 6 Then
            If bMode Then
                ws.Cells(lRow, lCol) = ws.Cells(lRow - 1, lCol)
            Else
                ws.Cells(lRow, lCol) = ""
            End If
        End If
    Next
End Sub

'=============================
' コピーボタン押し忘れチェック共通処理
'=============================
Public Function CheckCopyButtonAllSheets() As Boolean
    Dim ws As Worksheet
    Dim checkRows As Variant
    checkRows = Array(87, 104)
    Dim r As Variant
    Dim c As Long
    Dim isMiss As Boolean
    Dim msg As String
    Dim excludeSheets As Variant

    ' 除外シートの設定（必要に応じて変更）
    excludeSheets = Array() ' なし
    'excludeSheets = Array("ふわりｼﾌｫﾝいちご") ' あり

    For Each ws In ThisWorkbook.Worksheets
        Dim doCheck As Boolean
        If UBound(excludeSheets) >= 0 Then
            If IsError(Application.Match(ws.Name, excludeSheets, 0)) Then
                doCheck = True
            End If
        Else
            doCheck = True
        End If
        If doCheck Then
            For Each r In checkRows
                Dim missInRow As Boolean
                Dim lastCol As Long, maxCol As Long
                lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
                maxCol = 14
                For c = 3 To Application.WorksheetFunction.Max(lastCol, maxCol)
                    If c <> 6 Then
                        ' 通常時はピンク（実データの最終列まで）
                        If c <= lastCol Then
                            ws.Cells(r, c).Interior.Color = RGB(255, 199, 206)
                        End If
                        ' 差異があれば薄い黄色（14列目まで）
                        If c <= maxCol Then
                            If CStr(ws.Cells(r, c).Value) <> CStr(ws.Cells(r - 1, c).Value) Then
                                ws.Cells(r, c).Interior.Color = RGB(255, 255, 153) ' 薄い黄色
                                isMiss = True
                                missInRow = True
                            End If
                        End If
                    End If
                Next c
                If missInRow Then
                    msg = msg & ws.Name & "：行" & r & vbCrLf
                End If
            Next r
        End If
    Next ws
    If isMiss Then
        MsgBox "コピーボタンの押し忘れがある可能性があります。" & vbCrLf & "薄い黄色のセルを確認してください。", vbExclamation
    End If
    CheckCopyButtonAllSheets = isMiss
End Function

'=============================
' vba_Form.xlsm連携共通関数
'=============================
Public Function OpenVbaFormAllSheets() As String
    Dim oWB As Workbook
    Dim ws As Worksheet

    Set oWB = Workbooks.Open(cnFile, False, True)
    For Each ws In ThisWorkbook.Worksheets
        ws.Copy After:=oWB.Sheets(oWB.Sheets.Count)
    Next ws
    
    Application.Run "'" & cnFile & "'!Bas_Main.subMain"
    OpenVbaFormAllSheets = Trim(oWB.Sheets("VBA").Cells(1, 1))
    oWB.Close False
    
    Set oWB = Nothing
End Function



