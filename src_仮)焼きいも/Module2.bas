Attribute VB_Name = "Module2"
Option Explicit

' 共通定数・関数をまとめる標準モジュール
Public Const cnFile As String = "\\srv0103\ALL\SEISAN\栄養計算\新表示用　計算表\製品登録申請用\栄養計算フォーマット用分類一覧\vba_Form.xlsm"

'--- 既存ボタン呼び出し用のエイリアス（後方互換） ---
Public Sub subCopyClear(lRow As Long, bMode As Boolean)
    CopyOrClearRow ActiveSheet, lRow, bMode
End Sub

'--- コピー・クリア共通処理 ---
Public Sub CopyOrClearRow(ws As Worksheet, lRow As Long, bMode As Boolean)
    Dim lCol As Long
    For lCol = 3 To 12
        If lCol <> 6 Then
            If bMode Then
                ws.Cells(lRow, lCol) = ws.Cells(lRow - 1, lCol)
            Else
                ws.Cells(lRow, lCol) = ""
            End If
        End If
    Next
End Sub

'--- コピーボタン押し忘れチェック共通処理 ---
Public Sub CheckCopyButtonAllSheets()
    Dim ws As Worksheet
    Dim checkRows As Variant
    checkRows = Array(87, 104)
    Dim r As Variant
    Dim c As Long
    Dim isMiss As Boolean
    Dim msg As String
    For Each ws In ThisWorkbook.Worksheets
        For Each r In checkRows
            Dim missInRow As Boolean
            For c = 3 To 12
                If c <> 6 Then
                    ' 通常時はピンク
                    ws.Cells(r, c).Interior.Color = RGB(255, 199, 206)
                    ' 差異があれば薄い黄色
                    If CStr(ws.Cells(r, c).Value) <> CStr(ws.Cells(r - 1, c).Value) Then
                        ws.Cells(r, c).Interior.Color = RGB(255, 255, 153) ' 薄い黄色
                        isMiss = True
                        missInRow = True
                    End If
                End If
            Next c
            If missInRow Then
                msg = msg & ws.Name & "：行" & r & vbCrLf
            End If
        Next r
    Next ws
    If isMiss Then
        MsgBox "コピーボタンの押し忘れがある可能性があります。" & vbCrLf & "薄い黄色のセルを確認してください。", vbExclamation
    End If
End Sub

' vba_Form.xlsmを開いて全シートをコピーし、処理を実行して結果を返す共通関数
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

