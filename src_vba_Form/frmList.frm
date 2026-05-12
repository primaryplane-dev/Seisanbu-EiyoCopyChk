VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmList 
   Caption         =   "分類リスト"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7725
   OleObjectBlob   =   "frmList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dic As Dictionary

Private Sub cmdCancel_Click()
    If Not dic Is Nothing Then Set dic = Nothing
    Unload Me
End Sub

Private Sub cmdRegist_Click()
    If lstData.ListIndex = -1 Then Exit Sub
    If Not dic Is Nothing Then Set dic = Nothing
    stSheet.Cells(1, 1) = lstData.Text
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Call subBeforeEdit
    Call subSetDictionary
    Call subAfterEdit
    txtWord.Text = ""
    lstData.Clear
    Call cmdSearch_Click
End Sub
Private Sub subSetDictionary()
    Dim oWB     As Workbook
    Dim i       As Long
    Dim sKey    As String
    Set oWB = Workbooks.Open(cnFile, False, True)
    Set dic = New Dictionary
    For i = 2 To 200
        sKey = Trim(oWB.Sheets(1).Cells(i, 1))
        If Not sKey = "" Then
            If Not dic.Exists(sKey) Then dic.Add sKey, "DMY"
        End If
    Next
    oWB.Close True
    Set oWB = Nothing
End Sub

Private Sub cmdSearch_Click()
    txtWord.Text = Trim(txtWord.Text)
    Call subMakeList
End Sub

Private Sub subMakeList()
    Dim vKey        As Variant
    Dim sWord       As String: sWord = StrConv(StrConv(txtWord.Text, vbNarrow), vbUpperCase)
    Dim sKey        As String
    lstData.Clear
    For Each vKey In dic
        sKey = StrConv(StrConv(vKey, vbNarrow), vbUpperCase)
        If Not InStr(sKey, sWord) = 0 Then
            lstData.AddItem vKey
        End If
    Next
End Sub
