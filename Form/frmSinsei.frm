VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSinsei 
   Caption         =   "申請変更"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   OleObjectBlob   =   "frmSinsei.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
    strRequestType = "見積"
End Sub

Private Sub MS_buttun_Click()
    strRequestType = "見積、請求"
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
    '見積書式の確定
        If IsNull(lstMitumoriType) = False Then
            MitumoriType = lstMitumoriType.Value
        Else
            MitumoriType = Hdata.strFormat
        End If
    '請求タイプの確定
        If IsNull(lstSeikyuuType) = False Then
            SeikyuuType = lstSeikyuuType.Value
        Else
            SeikyuuType = Hdata.strSeikyuuType
        End If
    'データの申請の発行および書き込み
        Call reWriteHyoudaiWithRequest(Hdata, ActiveWorkbook, strRequestType, MitumoriType, SeikyuuType)
    Next
    Unload frmSinsei
End Sub

Private Sub S_buttun_Click()
    strRequestType = "請求"
End Sub

Private Sub UserForm_Initialize()
    Call initItemList
End Sub
Private Sub initItemList()
'品名リストを初期化する
Dim strMType As String
Dim strSType As String
Dim Hdata As HyoudaiData
Dim strIndex As String
    Hdata = getHyoudaiData(getMitumoriNo)
'見積タイプの初期化
    lstMitumoriType.MultiSelect = fmMultiSelectSingle
    lstMitumoriType.List = getLstMhyouki
    strIndex = findIndex(getLstMhyouki, Hdata.strFormat)
    If Not strIndex Like "" Then
        lstMitumoriType.Selected(CLng(strIndex)) = True
    End If
'請求書式の初期化
    lstSeikyuuType.MultiSelect = fmMultiSelectSingle
    lstSeikyuuType.List = getLstSeikyuuType
    strIndex = findIndex(getLstSeikyuuType, Hdata.strSeikyuuType)
    If Not strIndex Like "" Then
        lstSeikyuuType.Selected(CLng(strIndex)) = True
    End If
'申請書式の初期化
    strRequestType = Hdata.strPublishRequestType
'データの書き込み
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
