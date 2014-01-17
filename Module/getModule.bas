Attribute VB_Name = "getModule"
Private Function getEndRow(strColumns As String, shtMy As Worksheet) As Long
'���߂�������ŕ\������Ă���ŏI�s���擾����
'���͂��Ȃ��ꍇ��0��Ԃ�
'strColumns: "a" a�� : "c:e" c-e��
Dim rngColumns As Range
Dim rngSort As Range
    Set rngColumns = shtMy.Columns(strColumns)
    With rngColumns
        Set rngSort = .Find(what:="*", after:=.Cells(1), LookIn:=xlValues, searchorder:=xlByRows, searchdirection:=xlPrevious)
    End With
    If rngSort Is Nothing Then
        getEndRow = 0
    Else
        getEndRow = rngSort.row
    End If
End Function
Private Function getHyoudaiDataV2(rngMno As Range) As HyoudaiData
'rngMno����Ή�����\��f�[�^��Ԃ�
    With rngMno
        getHyoudaiDataV2.strSerial = .Cells(1, 0)         '���͔ԍ�
        getHyoudaiDataV2.strMitumoriNo = .Cells(1, 1)     '����No
        getHyoudaiDataV2.strCustomer = .Cells(1, 2)       '����
        getHyoudaiDataV2.dteMitumoriDay = .Cells(1, 3)      '���ϓ�
        getHyoudaiDataV2.strFormat = .Cells(1, 4)         '�����^�C�v
        getHyoudaiDataV2.strBumon = .Cells(1, 5)         '�S������
        getHyoudaiDataV2.strSite = .Cells(1, 6)      '���ݒn
        getHyoudaiDataV2.strLocation = .Cells(1, 7)      '�ʒu
        getHyoudaiDataV2.strKiHyouki = .Cells(1, 8)       '�M�\�L
        getHyoudaiDataV2.strName = .Cells(1, 9)      '���O
        getHyoudaiDataV2.strContents = .Cells(1, 10)      '���e
        getHyoudaiDataV2.strDeliveryPlace = .Cells(1, 11)  '�[���ꏊ
        getHyoudaiDataV2.strSiharai = .Cells(1, 12)        '�x��������
        getHyoudaiDataV2.strYuukoukikann = .Cells(1, 13)   '�L������
        getHyoudaiDataV2.dblProceeds = .Cells(1, 14)       '���z�i�ō��j
        getHyoudaiDataV2.dblSum = .Cells(1, 15)            '���z(�ŕʁj
        getHyoudaiDataV2.dblCost = .Cells(1, 16)           '�����i�ŕʁj
        getHyoudaiDataV2.strNotes = .Cells(1, 17)          '����
        getHyoudaiDataV2.strMaker = .Cells(1, 18)          '�쐬��
        getHyoudaiDataV2.dteSeikyuuDay = .Cells(1, 19)       '������
        getHyoudaiDataV2.strSeikyuuType = .Cells(1, 20)    '�������@
        getHyoudaiDataV2.strMsinsei = .Cells(1, 21)        '�\������
        getHyoudaiDataV2.dblTaxRate = .Cells(1, 22)            '�����
        getHyoudaiDataV2.strPublishRequestType = .Cells(1, 23)     '���s�\��
        getHyoudaiDataV2.strMitumoriPresentDay = .Cells(1, 24)     '���ϒ�o��
        getHyoudaiDataV2.strAccountsDate = .Cells(1, 25)       '���ϓ�
        getHyoudaiDataV2.strCheckOfAccounts = .Cells(1, 26)     '�󒍊m�F��
        getHyoudaiDataV2.strCheckOfFinishing = .Cells(1, 27)   '�����m�F��
        getHyoudaiDataV2.strWorkReport = .Cells(1, 28)         '��ƕ񍐏�
        getHyoudaiDataV2.strUriageTuki = .Cells(1, 29)         '���㌎
    End With
End Function
Private Function getUtiwakeDataV2(rngMnos() As Range) As UtiwakeData()
'rngMnos()�������ڍ׃f�[�^��Ԃ�
Dim varCell As Variant
Dim Udata() As UtiwakeData
Dim i As Long
    i = 0
    For Each varCell In rngMnos()
        ReDim Preserve Udata(i)
        With varCell
            Udata(i).strMitumoriNo = .Cells(1, 1)
            Udata(i).strHeader = .Cells(1, 2)
            Udata(i).strContents = .Cells(1, 3)
            Udata(i).strSpec = .Cells(1, 4)
            Udata(i).strNumber = .Cells(1, 5)
            Udata(i).strUnit = .Cells(1, 6)
            Udata(i).strPrice = .Cells(1, 7)
            Udata(i).strSum = .Cells(1, 8)
            Udata(i).strNote = .Cells(1, 9)
            Udata(i).strPage = .Cells(1, 10)
        End With
        i = i + 1
    Next
    getUtiwakeDataV2 = Udata
End Function
Private Function getGyousyaDataV2(rngMnos() As Range) As GyousyaData()
'rngMnos()����Ή�����Ǝ҃f�[�^��Ԃ�
Dim varCell As Variant
Dim Gdata() As GyousyaData
Dim i As Long
    i = 0
    For Each varCell In rngMnos()
        ReDim Preserve Gdata(i)
        With varCell
            Gdata(i).strMitumoriNo = .Cells(1, 1)
            Gdata(i).strGyousya = .Cells(1, 2)
            Gdata(i).strCost = .Cells(1, 3)
            Gdata(i).strCostWithTax = .Cells(1, 4)
            Gdata(i).strBillMonth = .Cells(1, 5)
        End With
        i = i + 1
    Next
    getGyousyaDataV2 = Gdata
End Function
Private Function getSyousaiDataV2(rngMnos() As Range) As SyousaiData()
'rngMnos()����Ή�����ڍ׃f�[�^��Ԃ�
Dim varCell As Variant
Dim Sdata() As SyousaiData
Dim i As Long
    i = 0
    For Each varCell In rngMnos()
        ReDim Preserve Sdata(i)
        With varCell
            Sdata(i).strMitumoriNo = .Cells(1, 1)
            Sdata(i).strHeader = .Cells(1, 2)
            Sdata(i).strContents = .Cells(1, 3)
            Sdata(i).strSpec = .Cells(1, 4)
            Sdata(i).strNumber = .Cells(1, 5)
            Sdata(i).strUnit = .Cells(1, 6)
            Sdata(i).strPrice = .Cells(1, 7)
            Sdata(i).strSum = .Cells(1, 8)
            Sdata(i).strNote = .Cells(1, 9)
        End With
        i = i + 1
    Next
    getSyousaiDataV2 = Sdata
End Function
Private Function Uni(strData() As String) As String
'�z��ɗ^����ꂽ�f�[�^���d�����Ȃ���Ԃɂ��ĕԂ�
    Dim i As Long, j As Long, k As Long
    Dim strUni() As String
    Dim lngArray As Long
    Dim lngUni As Long
    Dim lngSame As Long
    
    ReDim strUni(UBound(strData()))
    lngArray = UBound(strData())
    If UBound(strData()) = -1 Then: Uni = "": GoTo ending
    If lngArray = 0 Then Uni = "0": GoTo ending
    strUni(0) = strData(0)
    j = 1
    k = 1
    lngUni = 1
    For i = 1 To lngArray
        For j = 0 To i
            If strData(i) Like strUni(j) Then
                lngSame = lngSame + 1
            End If
        Next
        If lngSame = 0 Then
            strUni(k) = strData(i)
            k = k + 1
        End If
        lngSame = 0
    Next
    ReDim strData(k - 1)
    ReDim Preserve strUni(k - 1)
    strData = strUni
    Uni = "uni"
ending:
End Function
Private Function getMitumoriNoOnRow(shtMy As Worksheet, row As Long) As String
'�s���猩��No���擾����
Dim lngRow As Long
Dim lngcolumn As Long
    Select Case shtMy.Name
    Case "�\��"
        lngRow = row
        If row < 3 Then Exit Function
        lngcolumn = 2
    Case "�ڍ�", "����", "�Ǝ�"
        lngRow = row
        If row < 3 Then Exit Function
        lngcolumn = 1
    Case "����"
        lngRow = 2
        lngcolumn = 6
    Case Else
        Exit Function
    End Select
    getMitumoriNoOnRow = shtMy.Cells(lngRow, lngcolumn)
End Function

Function getMitumoriNoRange(sht As Worksheet) As Range
'����No�����p��range��Ԃ�
    Select Case sht.Name
    Case "�\��", "����\��"
        Set getMitumoriNoRange = sht.Range("b:b")
    Case "�ڍ�", "����", "�Ǝ�", "����ڍ�", "����Ǝ�"
        Set getMitumoriNoRange = sht.Range("a:a")
    End Select
End Function
Function getBumonNameRange() As Range
'�S�����匟���p��range��Ԃ�
    Set getBumonNameRange = Sheets("�S������").Range("a:a")
End Function
Function getDataVersion() As Long
'�f�[�^�o�[�W�������擾����
    getDataVersion = Range("data_version")
End Function
Function getMitumoriNo() As String
'�V�[�g���猩��No���擾����
Dim shtMy As Worksheet
Dim lngRow As Long
Dim lngcolumn As Long
    Set shtMy = ActiveSheet
    getMitumoriNo = getMitumoriNoOnRow(shtMy, Selection.row)
End Function
Function getHyoudaiData(MitumoriNo As String, _
                        Optional shtMy As Worksheet) As HyoudaiData
'�\��f�[�^�擾�̃t�����g�G���h
Dim rngMno As Range
    If shtMy Is Nothing Then Set shtMy = ActiveWorkbook.Sheets("�\��")
    Set rngMno = findMitumoriNo(MitumoriNo, shtMy)
    If rngMno Is Nothing Then Exit Function
    Select Case getDataVersion
    Case 2
        getHyoudaiData = getHyoudaiDataV2(rngMno)
    End Select
End Function
Function getTeikiHyoudaiDatas(MitumoriNo() As String, _
                              bokMy As Workbook) As HyoudaiData()
'�����ƕ\��f�[�^�擾
'����NO�����ׂĐ��������Ƃ�O��ɓ��삵�Ă���̂Œ���
Dim i As Long
Dim shtMy As Worksheet
Dim Hdata() As HyoudaiData
    Set shtMy = bokMy.Sheets("����\��")
    ReDim Hdata(UBound(MitumoriNo))
    For i = 0 To UBound(MitumoriNo)
        Hdata(i) = getHyoudaiData(MitumoriNo(i), shtMy)
    Next
    getTeikiHyoudaiDatas = Hdata
End Function
Function getSyousaiData(MitumoriNo As String, _
                        Optional shtMy As Worksheet) As SyousaiData()
'�ڍ׃f�[�^�擾�̃t�����g�G���h
'�f�[�^�������ꍇ�A�S�f�[�^���󔒂̈�s��Ԃ�
Dim rngMnos() As Range
Dim Sdata() As SyousaiData
    If shtMy Is Nothing Then Set shtMy = ActiveWorkbook.Sheets("�ڍ�")
    If findMitumoriNo(MitumoriNo, shtMy) Is Nothing Then
        ReDim getSyousaiData(0)
        Exit Function
    End If
    rngMnos() = findMitumoriNumRanges(MitumoriNo, shtMy)
    Select Case getDataVersion
    Case 2
        getSyousaiData = getSyousaiDataV2(rngMnos)
    End Select
End Function
Function getGyousyaData(MitumoriNo As String, _
                        Optional shtMy As Worksheet) As GyousyaData()
'�Ǝ҃f�[�^�擾�̃t�����g�G���h
'�f�[�^�������ꍇ�A�S�f�[�^���󔒂̈�s��Ԃ�
Dim rngMnos() As Range
    If shtMy Is Nothing Then Set shtMy = ActiveWorkbook.Sheets("�Ǝ�")
    If findMitumoriNo(MitumoriNo, shtMy) Is Nothing Then
        ReDim getGyousyaData(0)
        Exit Function
    End If
    rngMnos() = findMitumoriNumRanges(MitumoriNo, shtMy)
    Select Case getDataVersion
    Case 2
        getGyousyaData = getSyousaiDataV2(rngMnos)
    End Select
End Function
Function getUtiwakeData(MitumoriNo As String) As UtiwakeData()
'�Ǝ҃f�[�^�擾�̃t�����g�G���h
'�f�[�^�������ꍇ�A�S�f�[�^���󔒂̈�s��Ԃ�
Dim rngMnos() As Range
    If findMitumoriNo(MitumoriNo, ActiveWorkbook.Sheets("����")) Is Nothing Then
        ReDim getUtiwakeData(0)
        Exit Function
    End If
    rngMnos() = findMitumoriNumRanges(MitumoriNo, ActiveWorkbook.Sheets("����"))
    Select Case getDataVersion
    Case 2
        getUtiwakeData = getUtiwakeDataV2(rngMnos)
    End Select
End Function
Function getSheetInputMitumoriNo() As String
'���̓V�[�g���猩��No���擾����
Dim shtMy As Worksheet
Set shtMy = Sheets("����")
    getSheetInputMitumoriNo = shtMy.Range("d2")
End Function
Function getSheetInputMitumoriType() As String
'���̓V�[�g���猩�ϕ��@���擾����
Dim shtMy As Worksheet
Set shtMy = Sheets("����")
    getSheetInputMitumoriType = shtMy.Range("f2")
End Function
Function getSheetInputHyoudai(bokMy As Workbook) As HyoudaiData
'���̓V�[�g����\��f�[�^���擾����
Dim shtMy As Worksheet
    Set shtMy = bokMy.Worksheets("����")
    getSheetInputHyoudai.strCustomer = shtMy.Range("b2")
    getSheetInputHyoudai.strMitumoriNo = getSheetInputMitumoriNo
    getSheetInputHyoudai.strMaker = shtMy.Range("h2")
    getSheetInputHyoudai.dteMitumoriDay = shtMy.Range("c5")
    getSheetInputHyoudai.strFormat = shtMy.Range("d5")
    getSheetInputHyoudai.strBumon = shtMy.Range("b5")
    getSheetInputHyoudai.strSite = shtMy.Range("b8")
    getSheetInputHyoudai.strLocation = shtMy.Range("e8")
    getSheetInputHyoudai.strName = shtMy.Range("b11")
    getSheetInputHyoudai.strContents = shtMy.Range("b14")
    getSheetInputHyoudai.strKiHyouki = shtMy.Range("c11")
    getSheetInputHyoudai.strSeikyuuType = shtMy.Range("e14")
    getSheetInputHyoudai.strYuukoukikann = shtMy.Range("h14")
    getSheetInputHyoudai.strSiharai = shtMy.Range("g14")
    getSheetInputHyoudai.dblTaxRate = shtMy.Range("g5")
    getSheetInputHyoudai.strPublishRequestType = shtMy.Range("h8")
    getSheetInputHyoudai.dblSum = shtMy.Range("g35")
End Function
Function getSheetInputSyousaiStartRow() As Long
'���̓V�[�g�̏ڍ׃f�[�^�J�n�s��Ԃ�
    getSheetInputSyousaiStartRow = 17
End Function
Function getSyousaiRows() As Long
'�ڍ׃f�[�^�̍s����Ԃ�
    getSyousaiRows = 17
End Function
Function getSheetInputGyousyaStartRow() As Long
'���̓V�[�g�̋Ǝ҃f�[�^�J�n�s��Ԃ�
    getSheetInputGyousyaStartRow = getSheetInputSyousaiStartRow
End Function
Function getSheetInputSyousaiData(bokMy As Workbook, Mno As String) As SyousaiData()
'���̓V�[�g����ڍ׃f�[�^���擾����
'�ڍ׃f�[�^�������ꍇ�A����No���󔒂̈�s��Ԃ�
Dim shtMy As Worksheet
Dim Sdata() As SyousaiData
Dim startRow As Long
Dim i As Long
ReDim Sdata(getSyousaiRows)
    Set shtMy = bokMy.Worksheets("����")
    startRow = getSheetInputSyousaiStartRow
    For i = 0 To UBound(Sdata)
        Sdata(i).strMitumoriNo = Mno
        Sdata(i).strHeader = shtMy.Cells(startRow + i, 1)
        Sdata(i).strContents = shtMy.Cells(startRow + i, 2)
        Sdata(i).strSpec = shtMy.Cells(startRow + i, 3)
        Sdata(i).strNumber = shtMy.Cells(startRow + i, 4)
        Sdata(i).strUnit = shtMy.Cells(startRow + i, 5)
        Sdata(i).strPrice = shtMy.Cells(startRow + i, 6)
        Sdata(i).strSum = shtMy.Cells(startRow + i, 7)
        Sdata(i).strNote = shtMy.Cells(startRow + i, 8)
    Next
    For i = getSyousaiRows To 0 Step -1
        If Not Sdata(i).strHeader Like "" Then Exit For
        If Not Sdata(i).strContents Like "" Then Exit For
        If Not Sdata(i).strSpec Like "" Then Exit For
        If Not Sdata(i).strNumber Like "" Then Exit For
        If Not Sdata(i).strUnit Like "" Then Exit For
        If Not Sdata(i).strPrice Like "" Then Exit For
        If Not Sdata(i).strSum Like "" Then Exit For
        If Not Sdata(i).strNote Like "" Then Exit For
        If i = 0 Then
            Sdata(i).strMitumoriNo = ""
            Exit For
        End If
        ReDim Preserve Sdata(i - 1)
    Next
    getSheetInputSyousaiData = Sdata
End Function
Function getSheetInputGyousyaData(bokMy As Workbook, Mno As String) As GyousyaData()
'���̓V�[�g����Ǝ҃f�[�^���擾����
Dim shtMy As Worksheet
Dim Gdata() As GyousyaData
Dim i As Long
Dim j As Long
Dim startRow As Long
    startRow = getSheetInputGyousyaStartRow
    Set shtMy = bokMy.Worksheets("����")
    ReDim Gdata(17)
    j = 0
    For i = 0 To UBound(Gdata)
        If Not shtMy.Cells(startRow + i, 10) Like "" Then
            Gdata(j).strMitumoriNo = Mno
            Gdata(j).strGyousya = shtMy.Cells(startRow + i, 10)
            Gdata(j).strCost = shtMy.Cells(startRow + i, 11)
            j = j + 1
        End If
    Next
    If j > 0 Then
        ReDim Preserve Gdata(j - 1)
    Else
        ReDim Preserve Gdata(0)
    End If
    getSheetInputGyousyaData = Gdata
End Function
Function getUtiwakePageRows() As Long
'����ڍׂ̍s����Ԃ�
    getUtiwakePageRows = 39
End Function
Function getUtiwakeDataPage(strPage As String) As Long
'����f�[�^��strPage�f�[�^����y�[�W��Ԃ�
    getUtiwakeDataPage = CLng(Right(strPage, Len(strPage) - 1))
End Function
Function getSheetInputUtiwakeFeedRow(page As Long) As Long
'���̓V�[�g�̓���ڍ׎擾�J�n�s��Ԃ�
Dim startRow As Long
Dim pageRows As Long
    pageRows = getUtiwakePageRows
    startRow = 40
    getSheetInputUtiwakeFeedRow = startRow + ((page - 1) * (pageRows + 1)) + page
End Function
Function getSheetInputUtiwakePage(bokMy As Workbook, page As Long, Mno As String) As UtiwakeData()
'���̓V�[�g����Y���y�[�W�̓���ڍ׃f�[�^���擾����
'�y�[�W�ɓ��͂��Ȃ��ꍇ�ɂ͌��ϔԍ��̋�̈�s�f�[�^��Ԃ�
Dim shtMy As Worksheet
Dim row As Long
Dim FeedRow As Long
Dim Udata() As UtiwakeData
    Set shtMy = bokMy.Worksheets("����")
    FeedRow = getSheetInputUtiwakeFeedRow(page)
    ReDim Udata(getUtiwakePageRows)
    For row = 0 To UBound(Udata)
       Udata(row).strMitumoriNo = Mno
       Udata(row).strHeader = shtMy.Cells(FeedRow + row, 1)
       Udata(row).strContents = shtMy.Cells(FeedRow + row, 2)
       Udata(row).strSpec = shtMy.Cells(FeedRow + row, 3)
       Udata(row).strNumber = shtMy.Cells(FeedRow + row, 4)
       Udata(row).strUnit = shtMy.Cells(FeedRow + row, 5)
       Udata(row).strPrice = shtMy.Cells(FeedRow + row, 6)
       Udata(row).strSum = shtMy.Cells(FeedRow + row, 7)
       Udata(row).strNote = shtMy.Cells(FeedRow + row, 8)
       Udata(row).strPage = "P" & page
    Next
    For row = getUtiwakePageRows To 0 Step -1
        If Not Udata(row).strHeader Like "" Then Exit For
        If Not Udata(row).strContents Like "" Then Exit For
        If Not Udata(row).strSpec Like "" Then Exit For
        If Not Udata(row).strNumber Like "" Then Exit For
        If Not Udata(row).strUnit Like "" Then Exit For
        If Not Udata(row).strPrice Like "" Then Exit For
        If Not Udata(row).strSum Like "" Then Exit For
        If Not Udata(row).strNote Like "" Then Exit For
        If row = 0 Then
            Udata(row).strMitumoriNo = ""
            Exit For
        End If
        ReDim Preserve Udata(row - 1)
    Next
    getSheetInputUtiwakePage = Udata
End Function
Function getHyoudaiEndCell(bokMy As Workbook) As Range
Dim shtMy As Worksheet
Dim lngRow As Long
Set shtMy = bokMy.Sheets("�\��")
    lngRow = getHyoudaiEndRow(bokMy)
    Set getHyoudaiEndCell = shtMy.Cells(lngRow, 2)
End Function
Function getHyoudaiEndRow(bokMy As Workbook) As Long
'�\��V�[�g�̓��͍ŏI�s���擾����
    getHyoudaiEndRow = getEndRow("b", bokMy.Sheets("�\��"))
End Function
Function getSyousaiEndRow(bokMy As Workbook) As Long
'�ڍ׃V�[�g�̓��͍ŏI�s���擾����
    getSyousaiEndRow = getEndRow("a", bokMy.Sheets("�ڍ�"))
End Function
Function getGyousyaEndRow(bokMy As Workbook) As Long
'�Ǝ҃V�[�g�̓��͍ŏI�s���擾����
    getGyousyaEndRow = getEndRow("a", bokMy.Sheets("�Ǝ�"))
End Function
Function getUtiwakeEndRow(bokMy As Workbook) As Long
'����V�[�g�̓��͍ŏI�s���擾����
    getUtiwakeEndRow = getEndRow("a", bokMy.Sheets("����"))
End Function
Function getDeskTopPath() As String
    Dim WSH As Variant
    Set WSH = CreateObject("WScript.Shell")
    getDeskTopPath = WSH.specialfolders("Desktop")
End Function
Function getLstMhyouki() As String()
'���ϕ\�L���@�̃��X�g���擾����
Dim rngMy As Range
Dim cellValue() As String
Dim i As Long
    Set rngMy = Range("lstMhyouki")
    ReDim cellValue(rngMy.Count - 1)
    For i = 0 To rngMy.Count - 1
        cellValue(i) = rngMy.Cells(i + 1)
    Next
    getLstMhyouki = cellValue
End Function
Function getLstSeikyuuType() As String()
'�������@�̃��X�g���擾����
Dim rngMy As Range
Dim cellValue() As String
Dim i As Long
    Set rngMy = Range("lstSeikyuuType")
    ReDim cellValue(rngMy.Count - 1)
    For i = 0 To rngMy.Count - 1
        cellValue(i) = rngMy.Cells(i + 1)
    Next
    getLstSeikyuuType = cellValue
End Function
Function getMitumoriNoOnRangeAreas() As String()
'�I�����Ă���S�Ă̌���No���d���Ȃ��ŕԂ�
'�������A���Ԃ̓\�[�g���Ȃ�
Dim i As Long, j As Long, k As Long
Dim lngAreas As Long
Dim lngCellCount As Long
Dim strMno As String
Dim strMnos() As String
Dim lngERROR As Long
Dim shtMy As Worksheet
    
    On Error Resume Next
    lngAreas = Selection.Areas.Count
    On Error GoTo 0
    If Err.Number <> 0 Then getRngAreasRows = "0": GoTo ending
    Set shtMy = Selection.Worksheet
    For i = 1 To lngAreas
        lngCellCount = Selection.Areas(i).Rows.Count
        For j = 1 To lngCellCount
            strMno = strMno & " " & getMitumoriNoOnRow(shtMy, Selection.Areas(i).Rows(j).row)
            strMno = Trim(strMno)
        Next
    Next
    strMnos() = Split(strMno, " ")
    If UBound(strMnos) < 0 Then ReDim strMnos(0)
    Call Uni(strMnos)
    getMitumoriNoOnRangeAreas = strMnos
ending:
End Function
Function getTorikomiWorkbooks(bokMyName As String) As Workbook()
'bokMyName�̃u�b�N���ȊO�̓��̓V�[�g�����݂���workbook��z��ŕԂ�
'���݂��Ȃ��ꍇ�A��̔z�񂪋A��
Dim bokThis As Workbook
Dim bokMy() As Workbook
Dim i As Long
Dim j As Long
    j = 0
    For i = 1 To Application.Workbooks.Count
        Set bokThis = Application.Workbooks(i)
        If Not bokMyName Like bokThis.Name Then
            If has_same_sheet("����", bokThis) = True Then
                ReDim Preserve bokMy(j)
                Set bokMy(j) = bokThis
                j = j + 1
            End If
        End If
    Next
    getTorikomiWorkbooks = bokMy
End Function
Function getTorikomiWorkbook(bokMyName As String) As Workbook
'bokMyName�̃u�b�N���ȊO�̓��̓V�[�g������workbook��Ԃ�
'���݂��Ȃ��ꍇ�A��̃I�u�W�F�N�g��Ԃ�
Dim bokThis As Workbook
Dim i As Long
    For i = 1 To Application.Workbooks.Count
        Set bokThis = Application.Workbooks(i)
        If Not bokMyName Like bokThis.Name Then
            If has_same_sheet("����", bokThis) = True Then
                Set getTorikomiWorkbook = bokThis
                Exit Function
            End If
        End If
    Next
End Function
