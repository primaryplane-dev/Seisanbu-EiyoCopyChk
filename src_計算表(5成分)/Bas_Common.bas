
Attribute VB_Name = "Common"
Option Explicit

'=============================
' 共通定数
'=============================
Public Const cnFile As String = "\\srv0103\ALL\SEISAN\�?養計算\新表示用　計算表\製品登録申請用\�?養計算フォーマット用�?類一覧\vba_Form.xlsm"

'=============================
' コピ�?�・クリア共通�?��?
'=============================
'--- 既存�?�タン呼び出し用のエイリアス?��後方互換?�? ---
Public Sub subCopyClear(lRow As Long, bMode As Boolean)
    CopyOrClearRow ActiveSheet, lRow, bMode
End Sub

'--- コピ�?�・クリア共通�?��? ---
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

'=============================
' コピ�?�ボタン押し忘れチェ�?ク共通�?��?
' �߂�l: True=�����Y�ꂠ��(�ۑ�/�I���𒆎~), False=���Ȃ�
'=============================
Public Function CheckCopyButtonAllSheets() As Boolean
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
                    ' 通常時�?�ピンク
                    ws.Cells(r, c).Interior.Color = RGB(255, 199, 206)
                    ' 差異があれ�?��?�?�?色
                    If CStr(ws.Cells(r, c).Value) <> CStr(ws.Cells(r - 1, c).Value) Then
                        ws.Cells(r, c).Interior.Color = RGB(255, 255, 153) ' �?�?�?色
                        isMiss = True
                        missInRow = True
                    End If
                End If
            Next c
            If missInRow Then
                msg = msg & ws.Name & "?���?" & r & vbCrLf
            End If
        Next r
    Next ws
    If isMiss Then
        MsgBox "コピ�?�ボタンの押し忘れがある可能性があります�?" & vbCrLf & "�?�?�?色のセルを確認してください�?", vbExclamation
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
