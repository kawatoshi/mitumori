Attribute VB_Name = "CheckModule"
Option Explicit
Function is_match(strPat As String, strTxt As String) As Boolean
Dim RegEx As Variant
Dim Matches As Variant
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .pattern = strPat
        .ignorecase = False
        Set Matches = .Execute(strTxt)
        If Matches.Count = 0 Then
            Exit Function
        End If
    End With
    is_match = True
End Function
Function chkFilter(strSheetname As String) As Boolean
'�����ɗ^����ꂽ�V�[�g�ɃI�[�g�t�B���^�[�ɂ���\���s��
'���邩��Ԃ�
Dim shtMy As Worksheet
Dim lngCount As Long, lngCounter As Long
    
    chkFilter = False
    lngCount = ActiveWorkbook.Worksheets.Count
    For lngCounter = 1 To lngCount
        If Worksheets(lngCounter).Name Like strSheetname Then
            Set shtMy = Worksheets(strSheetname)
            chkFilter = shtMy.FilterMode
        End If
    Next
End Function
Function chkAllSheetFilter() As String
Dim SheetNames() As String
Dim sName As Variant
    SheetNames = Split("�\��,�ڍ�,����,�Ǝ�", ",")
    For Each sName In SheetNames
        If chkFilter(CStr(sName)) = True Then
            chkAllSheetFilter = CStr(sName) & "�Ƀt�B���^�[���������Ă��܂��B"
            Exit Function
        End If
    Next
    chkAllSheetFilter = ""
End Function
Function Utiwake_is_empty(Udata() As UtiwakeData) As Boolean
'�^����ꂽ����Ƀf�[�^���܂܂�Ă��Ȃ�����Ԃ�
    If Udata(0).strMitumoriNo Like "" Then
        Utiwake_is_empty = True
    Else
        Utiwake_is_empty = False
    End If
End Function
Function same_MitumoriNo(MitumoriNo As String) As Boolean
'�^����ꂽ����No���\��ɑ��݂��邩��Ԃ�
Dim shtMy As Worksheet
Set shtMy = Sheets("�\��")
    If findMitumoriNo(MitumoriNo, shtMy) Is Nothing Then
        same_MitumoriNo = False
    Else
        same_MitumoriNo = True
    End If
End Function
Function chkInputSheetMitumoriNumber(MitumoriNo As String, MitumoriType As String)
Dim ans As String
Dim Mno As MitumoriNumber
Set Mno = New MitumoriNumber
    Select Case MitumoriType
    Case "�V�K"
        If MitumoriNo Like "" Then Exit Function
        ans = "����No�����͂���Ă��܂��B�V�K�̏ꍇ�͌���No�ɉ����L�����Ȃ��ł�������"
    Case "�Č���"
        If Mno.Push(MitumoriNo) = False Then ans = "�L���Ȍ���No�ł͂���܂���"
        If MitumoriNo Like "" Then ans = "����No�����͂���Ă��܂���B�L���Ȍ���No���L�����Ă�������"
        If same_MitumoriNo(Mno.MainNo & "*") = False Then ans = "�Č��ςɕK�v�ȃf�[�^��������܂���" & _
                                                            Chr(13) & "�\��V�[�g�Ƀf�[�^�̂��錩�ϔԍ����K�v�ł�"
    Case "���"
        If Mno.Push(MitumoriNo) = False Then ans = "�L���Ȍ���No�ł͂���܂���"
        If MitumoriNo Like "" Then ans = "����No�����͂���Ă��܂���B�L���Ȍ���No���L�����Ă�������"
        If same_MitumoriNo(Mno.MainNo & "*") = False Then ans = "������ςɕK�v�ȃf�[�^��������܂���" & _
                                                            Chr(13) & "�\��V�[�g�Ƀf�[�^�̂��錩�ϔԍ����K�v�ł�"
    Case "�]�L"
        If Mno.Push(MitumoriNo) = False Then ans = "�L���Ȍ���No�ł͂���܂���"
        If MitumoriNo Like "" Then ans = "����No�����͂���Ă��܂���B�L���Ȍ���No���L�����Ă�������"
        If same_MitumoriNo(Mno.Publish) = True Then ans = "���łɂ��̌���No�͎g�p����Ă��܂�"
    Case Else
        ans = "���σ^�C�v���s���ł�"
    End Select
    chkInputSheetMitumoriNumber = ans
End Function
Function is_match_Mitumori(strTxt As String) As Boolean
    is_match_Mitumori = is_match("����", strTxt)
End Function
Function SinseiType_is_Mitumori(MitumoriNo As String) As Boolean
Dim Hdata As HyoudaiData
    Hdata = getHyoudaiData(MitumoriNo)
    If Hdata.strMitumoriNo Like "" Then Exit Function
    SinseiType_is_Mitumori = is_match_Mitumori(Hdata.strPublishRequestType)
End Function
Function is_match_Seikyuu(strTxt As String) As Boolean
    is_match_Seikyuu = is_match("����", strTxt)
End Function
Function SinseiType_is_Seikyuu(MitumoriNo As String) As Boolean
Dim Hdata As HyoudaiData
    Hdata = getHyoudaiData(MitumoriNo)
    If Hdata.strMitumoriNo Like "" Then Exit Function
    SinseiType_is_Seikyuu = is_match_Seikyuu(Hdata.strPublishRequestType)
End Function
Function SinseiType_is_zeikomi(strFormatOrPublishRequestType As String) As Boolean
'���ς܂��͐������@���u�ō��v���܂�ł��邩��Ԃ�
    SinseiType_is_zeikomi = is_match("�ō�", strFormatOrPublishRequestType)
End Function
Function Book_is_opend(BokName As String) As Workbook
'�����̃u�b�N���J����Ă���΁A���̃u�b�N��Ԃ�
Dim bokMy As Variant
    For Each bokMy In Application.Workbooks
        If bokMy.Name Like BokName Then
            Set Book_is_opend = bokMy
            Exit Function
        End If
    Next
End Function
Function has_same_sheet(SheetName As String, bokCheck As Workbook) As Boolean
'sheetname�Ɠ����V�[�g�����݂��邩��Ԃ�
Dim shtMy As Worksheet
    For Each shtMy In bokCheck.Worksheets
        If shtMy.Name Like SheetName Then
            has_same_sheet = True
            Exit Function
        End If
    Next
End Function
