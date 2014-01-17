Attribute VB_Name = "PostModule"
Option Explicit

Function postStrDate(dteMy As Date) As String
'date�f�[�^�𕶎���ŕԂ�
    If dteMy = 0 Then
        postStrDate = ""
    Else
        postStrDate = CStr(dteMy)
    End If
End Function
Function postSumAndCost(bokMy As Workbook, Hdata As HyoudaiData) As Boolean
    Hdata.dblCost = postGyousyaSum(Hdata)
    Hdata.dblSum = postSyousaiSum(Hdata)
End Function
Function postHdataFormat(Hdata As HyoudaiData) As String
'�\��f�[�^�̏����^�C�v���m�肵�ĕԂ�
    If Hdata.strFormat Like "" Then
        postHdataFormat = findCustomerFormat(Hdata.strCustomer)
    Else
        postHdataFormat = Hdata.strFormat
    End If
End Function
Function postHdataSeikyuuType(Hdata As HyoudaiData) As String
'�\��f�[�^�̐\���������m�肵�ĕԂ�
    If Hdata.strSeikyuuType Like "" Then
        postHdataSeikyuuType = findCustomerSeikyuuType(Hdata.strCustomer)
    Else
        postHdataSeikyuuType = Hdata.strSeikyuuType
    End If
End Function
Function postNO(strSrc As String) As String
'strSrc�擪��No.��t���ĕԂ�
    postNO = "No." & strSrc
End Function
Function postSagyouDate(Hdata As HyoudaiData) As String
'��Ɠ��i�\��j�̕������Ԃ�
Dim mitumoriDay As String
Dim SeikyuuDay As String
    mitumoriDay = "���ϓ� : " & postStrDate(Hdata.dteMitumoriDay)
    If Hdata.dteSeikyuuDay > 0 Then
        SeikyuuDay = Chr(13) & Chr(10) & "������ : " & postStrDate(Hdata.dteSeikyuuDay)
    End If
    postSagyouDate = mitumoriDay & SeikyuuDay
End Function
Function postGyousyaNames(Hdata As HyoudaiData) As String
'�Ǝ҂̖��O��Ԃ�
Dim i As Long
Dim ans As String
Dim Gdata() As GyousyaData
    If Hdata.strMitumoriNo Like "" Then Exit Function
    Gdata = getGyousyaData(Hdata.strMitumoriNo)
    For i = 0 To UBound(Gdata)
        ans = ans & " " & Gdata(i).strGyousya
    Next
    postGyousyaNames = Trim(ans)
End Function
Function postGyousyaSum(Hdata As HyoudaiData) As Double
'�Ǝ҂̌������v��Ԃ�
Dim i As Long
Dim ans As Double
Dim Gdata() As GyousyaData
postGyousyaSum = 0
    If Hdata.strMitumoriNo Like "" Then Exit Function
    Gdata = getGyousyaData(Hdata.strMitumoriNo)
    For i = 0 To UBound(Gdata)
        ans = ans + CDbl(Gdata(i).strCost)
    Next
    postGyousyaSum = ans
End Function
Function postGyousyaSumWithTax(Hdata As HyoudaiData) As Double
'�Ǝ҂̐ō��������v��Ԃ�
Dim i As Long
Dim ans As Double
Dim Gdata() As GyousyaData
postGyousyaSumWithTax = 0
    If Hdata.strMitumoriNo Like "" Then Exit Function
    Gdata = getGyousyaData(Hdata.strMitumoriNo)
    For i = 0 To UBound(Gdata)
        ans = ans + postWithTax(CDbl(Gdata(i).strCost), Hdata.dblTaxRate)
    Next
    postGyousyaSumWithTax = ans
End Function
Function postSyousaiSum(Hdata As HyoudaiData) As Double
'�ڍׂ̐ŕʍ��v��Ԃ�
Dim i As Long
Dim ans As Double
Dim Sdata() As SyousaiData
Dim dblSum As Double
postSyousaiSum = 0
    If Hdata.strMitumoriNo Like "" Then Exit Function
    Sdata = getSyousaiData(Hdata.strMitumoriNo)
    For i = 0 To UBound(Sdata)
        If IsNumeric(Sdata(i).strSum) = True Then
            dblSum = CDbl(Sdata(i).strSum)
        Else
            dblSum = 0
        End If
        ans = ans + dblSum
    Next
    If SinseiType_is_zeikomi(Hdata.strFormat) = True Then
        ans = postWtithOutTax(dblSum, Hdata.dblTaxRate)
    End If
    postSyousaiSum = ans
End Function
Function postSiharaiTypeOnCommision(Hdata As HyoudaiData) As String
'���\���ł̎x�������@�̕\�L��Ԃ�
    If Not Hdata.strSeikyuuType Like "" Then
        postSiharaiTypeOnCommision = Hdata.strSeikyuuType
        Exit Function
    End If
    If Not Hdata.strFormat Like "" Then
        postSiharaiTypeOnCommision = Hdata.strFormat
        Exit Function
    End If
    postSiharaiTypeOnCommision = Hdata.strSiharai
End Function
Function postSeikyuuDate(Hdata As HyoudaiData) As Date
'�\���^�C�v���琿�������m�肵�ĕԂ�
Dim dteNextDay As Date
Dim dteMDay As Date
Dim steSDay As Date
    dteMDay = Hdata.dteMitumoriDay
    dteNextDay = DateAdd("d", 1, dteMDay)
    If Hdata.dteSeikyuuDay > Hdata.dteMitumoriDay Then Exit Function
    If is_match_Seikyuu(Hdata.strPublishRequestType) = True Then
        postSeikyuuDate = endOfMonth(Now())
    End If
End Function
Function postCustomerName(strCustomer As String) As String
'���X�g�V�[�g�ɊY�����鐳�����̂�����΂����Ԃ��A
'�Ȃ��ꍇ�͈����̕������Ԃ�
Dim rngCustomer As Range
Dim strName As String
    Set rngCustomer = findCustomerName(strCustomer)
    If rngCustomer Is Nothing Then
        postCustomerName = strCustomer
    Else
        strName = rngCustomer.Cells(1, 3)
        If strName Like "" Then
            postCustomerName = strCustomer
        Else
            postCustomerName = strName
        End If
    End If
End Function
Function postWithTax(dblPrice As Double, dblRate As Double) As Double
'�ō����i��Ԃ�
    postWithTax = CDbl(Int(dblPrice * (1 + dblRate)))
End Function
Function postWtithOutTax(dblPrice As Double, dblRate As Double) As Double
'�Ŕ������i��Ԃ�
    postWtithOutTax = WorksheetFunction.RoundUp(dblPrice / (1 + dblRate), 0)
End Function
Function postTeikiToHyoudaiData(Hdata As HyoudaiData, strMonth As String) As HyoudaiData
'����\��f�[�^����\��f�[�^�ɓ]�L���鎞�̃f�[�^�ϊ������A���̒l��Ԃ�
Dim Hd As HyoudaiData
Dim Mno As MitumoriNumber
    Set Mno = New MitumoriNumber
    Hd = Hdata
    Call Mno.Push(Hdata.strMitumoriNo)
    Call Mno.to_teiki(strMonth)
    Hd.strMitumoriNo = Mno.Publish
    Hd.strSerial = ""
    Hd.dteMitumoriDay = Now()
    Hd.dteSeikyuuDay = endOfMonth(Hd.dteMitumoriDay)
    Hd.strUriageTuki = ""
    postTeikiToHyoudaiData = Hd
End Function
Function postTeikiToSyousaiData(Sdata() As SyousaiData, strMonth As String) As SyousaiData()
'����ڍ׃f�[�^����ڍ׃f�[�^�ɓ]�L���鎞�̕ϊ������ĕԂ�
Dim Sd() As SyousaiData
Dim i As Long
Dim Mno As MitumoriNumber
    Set Mno = New MitumoriNumber
    Sd = Sdata
    If Sd(0).strMitumoriNo Like "" Then
        ReDim postTeikiToSyousaiData(0)
        Exit Function
    End If
    Call Mno.Push(Sd(0).strMitumoriNo)
    Call Mno.to_teiki(strMonth)
    For i = 0 To UBound(Sdata)
        Sd(i).strMitumoriNo = Mno.Publish
    Next
    postTeikiToSyousaiData = Sd
End Function
Function postTeikiToGyousyaData(Gdata() As GyousyaData, strMonth As String) As GyousyaData()
'����Ǝ҃f�[�^����Ǝ҃f�[�^�ɓ]�L����Ƃ��̕ϊ������ĕԂ�
Dim Gd() As GyousyaData
Dim i As Long
Dim Mno As MitumoriNumber
    Set Mno = New MitumoriNumber
    Gd = Gdata
    If Gd(0).strMitumoriNo Like "" Then
        ReDim postTeikiToGyousyaData(0)
        Exit Function
    End If
    Call Mno.Push(Gd(0).strMitumoriNo)
    Call Mno.to_teiki(strMonth)
    For i = 0 To UBound(Gdata)
        Gd(i).strMitumoriNo = Mno.Publish
    Next
    postTeikiToGyousyaData = Gd
End Function
Function postSinseiCheckText() As String
'�I������Ă��錩�ς̃`�F�b�N�p�f�[�^���e�L�X�g�ŕԂ�
Dim Mnos() As String
Dim i As Long
Dim Hdata As HyoudaiData
Dim ans As String
    Mnos = getMitumoriNoOnRangeAreas
    For i = 0 To UBound(Mnos)
        Hdata = getHyoudaiData(Mnos(i))
        ans = Hdata.strMitumoriNo & " : " & Hdata.strCustomer & " : " & Hdata.strContents & vbCrLf
        postSinseiCheckText = postSinseiCheckText & ans
    Next
End Function
