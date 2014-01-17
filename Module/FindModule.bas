Attribute VB_Name = "FindModule"
Option Explicit
Private Function findCustomerListData(CustomerName As String, lngIndex As Long) As String
'�������X�g�ɊY������K��t�H�[�}�b�g�^�C�v��Ԃ�
Dim rngFind As Range
Dim ans As Range
Dim data As String
    Set rngFind = ActiveWorkbook.Sheets("���X�g").Range("a:a")
    Set ans = rngFind.Find(CustomerName, , , xlWhole, xlByColumns, xlNext, False, False)
    If ans Is Nothing Then Exit Function
    data = ans.Cells(1, 2)
    If data Like "" Then Exit Function
    findCustomerListData = Split(data, ",")(lngIndex)
End Function

Function findTeikiMitumoriNumbers(lngMonth As Long, bok As Workbook) As String()
'�^����ꂽ���ɊY����������ƌ���No��z��ŕԂ�
'�Y�����ς��Ȃ��ꍇ�ɂ͋�̔z�񂪋A��̂�for each���ŃG���[��������邱��
Dim rngFind As Range
Dim strMonth As String
Dim rngCell As Range
Dim firstAddress As String
Dim i As Long
Dim Mnos() As String
    findTeikiMitumoriNumbers = Split("")
    If lngMonth <= 0 Then Exit Function
    If lngMonth > 12 Then Exit Function
    Set rngFind = bok.Sheets("����\��").Range("ad:ad")
    strMonth = CStr(lngMonth)
    If Len(strMonth) = 1 Then strMonth = "0" & strMonth
    With rngFind
        Set rngCell = .Find(strMonth, , , xlPart, xlByColumns, xlNext, False, False)
        If Not rngCell Is Nothing Then
            i = 0
            firstAddress = rngCell.Address
            Do
                ReDim Preserve Mnos(i)
                Mnos(i) = rngCell.Cells(1, -27)
                Set rngCell = .FindNext(rngCell)
                i = i + 1
            Loop While Not rngCell Is Nothing And rngCell.Address <> firstAddress
            findTeikiMitumoriNumbers = Mnos()
        End If
    End With
End Function
Function findMitumoriNumRanges(MitumoriNo As String, sht As Worksheet) As Range()
'mitumorino�Ɠ������e�̃Z����z��ŕԂ�
Dim rngFind As Range
Dim rngCell As Range
Dim firstAddress As String
Dim i As Long
Dim rngCells() As Range
    Set rngFind = getMitumoriNoRange(sht)
    With rngFind
        Set rngCell = .Find(MitumoriNo, , , xlWhole, xlByColumns, xlNext, False, False)
        If Not rngCell Is Nothing Then
            i = 0
            firstAddress = rngCell.Address
            Do
                ReDim Preserve rngCells(i)
                Set rngCells(i) = rngCell
                Set rngCell = .FindNext(rngCell)
                i = i + 1
            Loop While Not rngCell Is Nothing And rngCell.Address <> firstAddress
            findMitumoriNumRanges = rngCells
        End If
    End With
End Function
Function findMitumoriNo(MitumoriNo As String, sht As Worksheet) As Range
'���ϔԍ��Ɠ������e�̃Z���Ԃ�
Dim rngFind As Range
    Set rngFind = getMitumoriNoRange(sht)
    Set findMitumoriNo = rngFind.Find(MitumoriNo, , , xlWhole, xlByColumns, xlNext, False, False)
End Function
Function findCustomerFormat(CustomerName As String) As String
'�������X�g�ɊY�����鏑���^�C�v��Ԃ�
    findCustomerFormat = findCustomerListData(CustomerName, 0)
End Function
Function findCustomerSeikyuuType(CustomerName As String) As String
'�������X�g�ɊY�����鐿��������Ԃ�
    findCustomerSeikyuuType = findCustomerListData(CustomerName, 1)
End Function
Function findBumonName(strBumon As String) As Range
'strBumon�Ɠ������e�̃Z����Ԃ�
Dim rngFind As Range
    Set rngFind = getBumonNameRange
    Set findBumonName = rngFind.Find(strBumon, , , xlWhole, xlByColumns, xlNext, False, False)
End Function
Function findCustomerName(strCustomer As String) As Range
'strCustomer�Ɠ������e�̃Z����Ԃ�
Dim rngFind As Range
    Set rngFind = Worksheets("���X�g").Range("a:a")
    Set findCustomerName = rngFind.Find(strCustomer, , , xlWhole, xlByColumns, xlNext, False, False)
End Function
