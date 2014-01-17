Attribute VB_Name = "VaridateModule"
Option Explicit
Private Function has_same_Mitumori(MitumoriNo As String) As Boolean

End Function
Private Function varidateData_is_input(data As String, Dtype As String) As String
'文字列データが入力されているか確認する
    If data Like "" Then
        varidateData_is_input = Dtype & "を記入してください" & Chr(13)
    End If
End Function
Private Function varidateDay_is_input(data As Date, Dtype As String) As String
'文字列データが入力されているか確認する
    If data Like "" Then
        varidateDay_is_input = Dtype & "を記入してください" & Chr(13)
    End If
End Function

Function varidateHyoudaiData(Hdata As HyoudaiData) As String
'表題データを検証する
'違反内容を文字列で返す
Dim Mno As MitumoriNumber
Dim ans As String
Set Mno = New MitumoriNumber
    If Mno.Push(Hdata.strMitumoriNo) = False Then _
        ans = "見積NOが不正規です" & Chr(13)
    ans = ans & varidateData_is_input(Hdata.strCustomer, "宛名")
    ans = ans & varidateData_is_input(Hdata.strContents, "工事内容")
    ans = ans & varidateData_is_input(Hdata.strSiharai, "支払い方法")
    ans = ans & varidateData_is_input(Hdata.strMaker, "作成者")
    ans = ans & varidateData_is_input(Hdata.strPublishRequestType, "発行申請")
    varidateHyoudaiData = Trim(ans)
End Function
