Attribute VB_Name = "VaridateModule"
Option Explicit
Private Function has_same_Mitumori(MitumoriNo As String) As Boolean

End Function
Private Function varidateData_is_input(data As String, Dtype As String) As String
'������f�[�^�����͂���Ă��邩�m�F����
    If data Like "" Then
        varidateData_is_input = Dtype & "���L�����Ă�������" & Chr(13)
    End If
End Function
Private Function varidateDay_is_input(data As Date, Dtype As String) As String
'������f�[�^�����͂���Ă��邩�m�F����
    If data Like "" Then
        varidateDay_is_input = Dtype & "���L�����Ă�������" & Chr(13)
    End If
End Function

Function varidateHyoudaiData(Hdata As HyoudaiData) As String
'�\��f�[�^�����؂���
'�ᔽ���e�𕶎���ŕԂ�
Dim Mno As MitumoriNumber
Dim ans As String
Set Mno = New MitumoriNumber
    If Mno.Push(Hdata.strMitumoriNo) = False Then _
        ans = "����NO���s���K�ł�" & Chr(13)
    ans = ans & varidateData_is_input(Hdata.strCustomer, "����")
    ans = ans & varidateData_is_input(Hdata.strContents, "�H�����e")
    ans = ans & varidateData_is_input(Hdata.strSiharai, "�x�������@")
    ans = ans & varidateData_is_input(Hdata.strMaker, "�쐬��")
    ans = ans & varidateData_is_input(Hdata.strPublishRequestType, "���s�\��")
    varidateHyoudaiData = Trim(ans)
End Function
