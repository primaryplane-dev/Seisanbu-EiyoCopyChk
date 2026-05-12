Attribute VB_Name = "modCalc"
Option Explicit

Public Function fncHantei(ByVal para1 As String, ByVal para2 As String) As String
    fncHantei = ""
    para1 = Trim(para1)
    para2 = Trim(para2)
    If para1 = "" Or para2 = "" Then Exit Function
    If Not IsNumeric(para1) Or Not IsNumeric(para1) Then Exit Function
    If fncCalc(para1, para2) = "" Then
        fncHantei = "Åõ"
    Else
        fncHantei = "Å~"
    End If
End Function

Public Function fncCalc(ByVal para1 As String, ByVal para2 As String) As String
    fncCalc = ""
    para1 = Trim(para1)
    para2 = Trim(para2)
    If para1 = "" Or para2 = "" Then Exit Function
    If Not IsNumeric(para1) Or Not IsNumeric(para1) Then Exit Function
    Dim cSuu1       As Currency: cSuu1 = para1
    Dim cSuu2       As Currency: cSuu2 = para2
    If cSuu2 < cSuu1 * 0.8 Then fncCalc = cSuu2 - cSuu1 * 0.8
    If cSuu2 > cSuu1 * 1.2 Then fncCalc = cSuu2 - cSuu1 * 1.2
End Function
