Attribute VB_Name = "MainModule"
Option Explicit

Function rekeyMitumori(MitumoriNo As String) As String
'mitumorino�̌��σf�[�^����̓V�[�g�֓]�L����
'��肪���������ꍇ�͕Ԓl�ɃG���[���L�ڂ���
'����ɏI�������ꍇ�͋󕶎���Ԃ�
Dim Mno As MitumoriNumber
Dim Hdata As HyoudaiData
Dim Sdata() As SyousaiData
Dim Gdata() As GyousyaData
Dim Udata() As UtiwakeData
Set Mno = New MitumoriNumber
    If Mno.Push(MitumoriNo) = False Then
        rekeyMitumori = "����No������������܂���": Exit Function
    End If
    Call initInputSheet
    Hdata = getHyoudaiData(MitumoriNo)
    Sdata = getSyousaiData(MitumoriNo)
    Udata = getUtiwakeData(MitumoriNo)
    Gdata = getGyousyaData(MitumoriNo)
    rekeyMitumori = writeSheetInputData(Hdata, Sdata, Udata, Gdata)
End Function
Function SheetInput(bokMy As Workbook) As Boolean
'���̓V�[�g����K�v������荞�݁A����No��������U�������
'�e�V�[�g�֓]�L����
Dim ans As String
Dim Mno As String
Dim Hdata As HyoudaiData
Dim Sdata() As SyousaiData
Dim Udata() As UtiwakeData
Dim Gdata() As GyousyaData
Dim i As Long
'�I�[�g�t�B���^�[�ɂ���\���̃`�F�b�N
    ans = chkAllSheetFilter
    If Not ans Like "" Then
        MsgBox (ans & " �����𒆒f���܂�")
        Exit Function
    End If
    Set bokMy = ActiveWorkbook
'�\��f�[�^�擾
    Hdata = getSheetInputHyoudai(bokMy)
    ans = chkInputSheetMitumoriNumber(Hdata.strMitumoriNo, getSheetInputMitumoriType)
    If Not ans Like "" Then
        MsgBox (ans)
        Exit Function
    End If
'����No�̍쐬
    Mno = postMitumoriNo(Hdata.strMitumoriNo, getSheetInputMitumoriType)
    Hdata.strMitumoriNo = Mno
    ans = varidateHyoudaiData(Hdata)
    If Not ans Like "" Then
        MsgBox (ans)
        Exit Function
    End If
    If Hdata.strMitumoriNo Like "" Then
        MsgBox ("����No���쐬�ł��܂���ł���" & Chr(13) & "�Ǘ��҂ɑ��k���Ă�������"): Exit Function
    End If
    If getSheetInputMitumoriType Like "�V�K" Then Hdata.strSerial = CStr(Range("serial") + 1)
    If Hdata.dteMitumoriDay = 0 Then Hdata.dteMitumoriDay = Now()
    Hdata.strFormat = postHdataFormat(Hdata)
    Hdata.strSeikyuuType = postHdataSeikyuuType(Hdata)
    Hdata.dteSeikyuuDay = postSeikyuuDate(Hdata)
'�ڍ׃f�[�^�擾
    Sdata = getSheetInputSyousaiData(bokMy, Mno)
'�Ǝ҃f�[�^�擾
    Gdata = getSheetInputGyousyaData(bokMy, Mno)
'�\��f�[�^��������
    Call writeNewHyoudaiData(Hdata, bokMy)
'�ڍ׃f�[�^��������
    Call writeNewSyousaiData(Sdata, bokMy)
'�Ǝ҃f�[�^��������
    Call writeNewGyousyaData(Gdata, bokMy)
'����ڍ׃f�[�^��荞�݋y�я�������
    i = 1
    Do While i < 100
        Udata = getSheetInputUtiwakePage(bokMy, i, Mno)
        If Udata(0).strMitumoriNo Like "" Then Exit Do
        Call writeNewUtiwakeData(Udata, bokMy)
        i = i + 1
    Loop
    SheetInput = True
End Function
Function SinseiKakunin(MitumoriNo As String) As Boolean
'�\�����m�F����
Dim shtSrc As Worksheet
Dim shtDst As Worksheet
Dim Hdata As HyoudaiData
Dim shtMy As Worksheet
Set shtMy = ActiveWorkbook.ActiveSheet
    If MitumoriNo Like "" Then Exit Function
    If SinseiType_is_Mitumori(MitumoriNo) = True Then
        Set shtSrc = Sheets("���ό���")
        Set shtDst = Sheets("���Ϗ�")
        If publishMitumori(MitumoriNo, shtSrc, shtDst) = False Then Exit Function
    End If
    If SinseiType_is_Seikyuu(MitumoriNo) = True Then
        Set shtSrc = Sheets("��������")
        Set shtDst = Sheets("������")
        If publishSeikyuu(MitumoriNo, shtSrc, shtDst) = False Then Exit Function
    End If
    Hdata = getHyoudaiData(MitumoriNo)
    Call postSumAndCost(shtSrc.Parent, Hdata)
    Call reWriteHyoudaiData(Hdata, shtSrc.Parent)
    shtDst.Activate
    shtDst.Range("b3").Select
    shtMy.Activate
    SinseiKakunin = True
End Function
Function MakeCommissions(Mno() As String, _
                         strPrinter As String, _
                         srtBokName As String, _
                         signature As Boolean) As String
'���˗����쐬����
'strPrinter�ň������v�����^�[�̐؂�ւ�
'strBokName������ꍇ�̓��[�N�u�b�N�ւ̏o��
'signature�͈�ӂ�������邩�ǂ���
 Dim shtCommission As Worksheet
 Dim DefaultPrinter As String
 Dim OutPutPrinter As String
 Dim i As Long
 Dim j As Long
 Dim k As Long
 Dim Mitumorinos(3) As String
 Dim MNumbers() As String
 Dim sumOfBox As Long
 Dim addCount As Long
    sumOfBox = 4
    Set shtCommission = ActiveWorkbook.Sheets("���˗���")
    DefaultPrinter = getPrinterName
    OutPutPrinter = findPrinter(strPrinter)
'�\���̔��s
    j = 0
    Application.ActivePrinter = OutPutPrinter
    addCount = sumOfBox - ((UBound(Mno) + 1) Mod sumOfBox)
    If addCount = 4 Then addCount = 0
    MNumbers = Mno
    ReDim Preserve MNumbers(UBound(Mno) + addCount)
    For i = 0 To UBound(MNumbers)
        Mitumorinos(j) = MNumbers(i)
        If j = 3 Then
            shtCommission.Activate
            Call WriteCommission(shtCommission, Mitumorinos, signature)
            For k = 0 To UBound(Mitumorinos)
                Mitumorinos(k) = ""
            Next
            j = -1
            Call publishCommission(shtCommission, srtBokName)
        End If
        j = j + 1
    Next
    Call visibleCommisionSignature(shtCommission, False)
    Application.ActivePrinter = DefaultPrinter
End Function
Sub MakeSinseiToHQ(Mno As String)
'��������No�̌��Ϗ��܂��͐��������쐬��
'�{���w��t�H���_�̎w��t�@�C���ɃR�s�[�A�ۑ�����
'�����Ƃ��ď������e�Ə������ʂ�String�ŕԂ�
Dim bokMy As Workbook
Dim bokDst As Workbook
Dim Hdata As HyoudaiData
Dim strReqFolder As String, strReqFile As String '�\���f�B���N�g���A�\���t�@�C����
Dim ans As String
    Set bokMy = ActiveWorkbook
    Hdata = getHyoudaiData(Mno)
    If Hdata.strMitumoriNo Like "" Then
        Call MsgBox("�L���Ȍ��ςł͂���܂���")
        Exit Sub
    End If
'���ϐ\��
    If is_match_Mitumori(Hdata.strPublishRequestType) = True Then
        strReqFolder = Range("Mitumori_dir").Value
        strReqFile = Range("Mitumori_file").Value
        ans = MakeSinsei(strReqFolder, strReqFile, Hdata, bokMy, "m", 85)
        If Not ans Like "" Then
            Call MsgBox(ans)
            ans = ""
        Else
            Hdata = getHyoudaiData(Hdata.strMitumoriNo)
            Call writeSinseiData(bokMy, Hdata)
        End If
    End If
'�����\��
    If is_match_Seikyuu(Hdata.strPublishRequestType) = True Then
        strReqFolder = Range("Seikyuu_dir").Value
        strReqFile = Range("Seikyuu_file").Value
        ans = MakeSinsei(strReqFolder, strReqFile, Hdata, bokMy, "s", 85)
        If Not ans Like "" Then
            Call MsgBox(ans)
            ans = ""
        Else
            Hdata = getHyoudaiData(Hdata.strMitumoriNo)
            Call writeSinseiData(bokMy, Hdata)
        End If
    End If
End Sub
Function TorikomiSheetInput() As String
'���u�b�N�̓��̓V�[�g����̎�荞��
'�P�Ǝ�荞�ݗp�ŁA��荞�݌��ʂ�Ԃ�
Dim bokMyName As String
Dim bokTorikomi As Workbook
Dim bokMy As Workbook
Dim Hdata As HyoudaiData
Dim Sdata() As SyousaiData
Dim Udata() As UtiwakeData
Dim UdataPage() As UtiwakeData
Dim i As Long
Dim j As Long
Dim k As Long
Dim Gdata() As GyousyaData
Dim dammyMno As String
    dammyMno = "torikomi"
    Set bokMy = ActiveWorkbook
    If Not has_same_sheet("����", bokMy) = True Then
        TorikomiSheetInput = "���̓V�[�g�Ŏ��s���Ă�������"
        Exit Function
    End If
    If getTorikomiWorkbook(bokMy.Name) Is Nothing Then
        TorikomiSheetInput = "��荞�ރu�b�N������܂���"
        Exit Function
    End If
    Set bokTorikomi = getTorikomiWorkbook(bokMy.Name)
    Hdata = getSheetInputHyoudai(bokTorikomi)
    Sdata = getSheetInputSyousaiData(bokTorikomi, dammyMno)
    i = 1
    j = 0
    ReDim UdataPage(0)
    Do While i < 100
        Udata = getSheetInputUtiwakePage(bokTorikomi, i, dammyMno)
        If Udata(0).strMitumoriNo Like "" Then
            Exit Do
        End If
        For k = 0 To UBound(Udata)
            ReDim Preserve UdataPage(j)
            UdataPage(j) = Udata(k)
            j = j + 1
        Next
        i = i + 1
    Loop
    Gdata = getSheetInputGyousyaData(bokTorikomi, dammyMno)
    bokMy.Activate
    Call writeSheetInputData(Hdata, Sdata, UdataPage, Gdata)
    bokTorikomi.Close
    bokMy.Activate
End Function
Function makeTeiki(lngMonth As Long, bokSrc As Workbook) As String
'�w�茎�̒���f�[�^��\��A�ڍׁA�Ǝ҂֓]�L����
Dim strMno() As String
Dim Hdata() As HyoudaiData
Dim i As Long
Dim Hd As HyoudaiData
Dim Sdata() As SyousaiData
Dim Gdata() As GyousyaData
    strMno = findTeikiMitumoriNumbers(lngMonth, bokSrc)
    If UBound(strMno) < 0 Then
        makeTeiki = CStr(lngMonth) & "���ɊY���������f�[�^�͂���܂���"
        Exit Function
    End If
    Hdata = getTeikiHyoudaiDatas(strMno, bokSrc)
    For i = 0 To UBound(Hdata)
        Hd = postTeikiToHyoudaiData(Hdata(i), CStr(lngMonth))
        If same_MitumoriNo(Hd.strMitumoriNo) = True Then
            makeTeiki = makeTeiki & "���쐬 : " & Hd.strMitumoriNo & Chr(13)
        Else
            Sdata = getSyousaiData(Hdata(i).strMitumoriNo, bokSrc.Sheets("����ڍ�"))
            Gdata = getGyousyaData(Hdata(i).strMitumoriNo, bokSrc.Sheets("����Ǝ�"))
            Call replaceTeikiHyoudaiFormat(Hd)
            Call replaceTeikiSyousaiFormat(Sdata, Hd.dteSeikyuuDay)
            Call writeNewHyoudaiData(Hd, bokSrc)
            Call writeNewSyousaiData(postTeikiToSyousaiData(Sdata, CStr(lngMonth)), bokSrc)
            Call writeNewGyousyaData(postTeikiToGyousyaData(Gdata, CStr(lngMonth)), bokSrc)
            makeTeiki = makeTeiki & "�쐬 : " & Hd.strMitumoriNo & Chr(13)
        End If
    Next
    makeTeiki = Trim(makeTeiki)
End Function
