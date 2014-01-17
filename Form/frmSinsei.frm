VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSinsei 
   Caption         =   "�\���ύX"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   OleObjectBlob   =   "frmSinsei.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSinsei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRequestType As String

Private Sub cancel_Button_Click()
    Unload frmSinsei
End Sub

Private Sub M_buttun_Click()
    strRequestType = "����"
End Sub

Private Sub MS_buttun_Click()
    strRequestType = "���ρA����"
End Sub

Private Sub ok_Button_Click()
Dim Hdata As HyoudaiData
Dim MitumoriType As String
Dim SeikyuuType As String
Dim Mnos() As String
Dim i As Long
    Mnos = getMitumoriNoOnRangeAreas
    For i = 0 To UBound(Mnos)
        Hdata = getHyoudaiData(Mnos(i))
    '���Ϗ����̊m��
        If IsNull(lstMitumoriType) = False Then
            MitumoriType = lstMitumoriType.Value
        Else
            MitumoriType = Hdata.strFormat
        End If
    '�����^�C�v�̊m��
        If IsNull(lstSeikyuuType) = False Then
            SeikyuuType = lstSeikyuuType.Value
        Else
            SeikyuuType = Hdata.strSeikyuuType
        End If
    '�f�[�^�̐\���̔��s����я�������
        Call reWriteHyoudaiWithRequest(Hdata, ActiveWorkbook, strRequestType, MitumoriType, SeikyuuType)
    Next
    Unload frmSinsei
End Sub

Private Sub S_buttun_Click()
    strRequestType = "����"
End Sub

Private Sub UserForm_Initialize()
    Call initItemList
End Sub
Private Sub initItemList()
'�i�����X�g������������
Dim strMType As String
Dim strSType As String
Dim Hdata As HyoudaiData
Dim strIndex As String
    Hdata = getHyoudaiData(getMitumoriNo)
'���σ^�C�v�̏�����
    lstMitumoriType.MultiSelect = fmMultiSelectSingle
    lstMitumoriType.List = getLstMhyouki
    strIndex = findIndex(getLstMhyouki, Hdata.strFormat)
    If Not strIndex Like "" Then
        lstMitumoriType.Selected(CLng(strIndex)) = True
    End If
'���������̏�����
    lstSeikyuuType.MultiSelect = fmMultiSelectSingle
    lstSeikyuuType.List = getLstSeikyuuType
    strIndex = findIndex(getLstSeikyuuType, Hdata.strSeikyuuType)
    If Not strIndex Like "" Then
        lstSeikyuuType.Selected(CLng(strIndex)) = True
    End If
'�\�������̏�����
    strRequestType = Hdata.strPublishRequestType
'�f�[�^�̏�������
    selectData.Text = postSinseiCheckText
End Sub
Private Function findIndex(lst() As String, strTxt As String) As String
Dim i As Long
    For i = 0 To UBound(lst)
        If lst(i) Like strTxt Then
            findIndex = CStr(i)
            Exit Function
        End If
    Next
End Function
