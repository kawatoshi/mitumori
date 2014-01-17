Attribute VB_Name = "FindModule"
Option Explicit
Private Function findCustomerListData(CustomerName As String, lngIndex As Long) As String
'宛名リストに該当する規定フォーマットタイプを返す
Dim rngFind As Range
Dim ans As Range
Dim data As String
    Set rngFind = ActiveWorkbook.Sheets("リスト").Range("a:a")
    Set ans = rngFind.Find(CustomerName, , , xlWhole, xlByColumns, xlNext, False, False)
    If ans Is Nothing Then Exit Function
    data = ans.Cells(1, 2)
    If data Like "" Then Exit Function
    findCustomerListData = Split(data, ",")(lngIndex)
End Function

Function findTeikiMitumoriNumbers(lngMonth As Long, bok As Workbook) As String()
'与えられた月に該当する定期作業見積Noを配列で返す
'該当見積がない場合には空の配列が帰るのでfor each文でエラーを回避すること
Dim rngFind As Range
Dim strMonth As String
Dim rngCell As Range
Dim firstAddress As String
Dim i As Long
Dim Mnos() As String
    findTeikiMitumoriNumbers = Split("")
    If lngMonth <= 0 Then Exit Function
    If lngMonth > 12 Then Exit Function
    Set rngFind = bok.Sheets("定期表題").Range("ad:ad")
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
'mitumorinoと同じ内容のセルを配列で返す
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
'見積番号と同じ内容のセル返す
Dim rngFind As Range
    Set rngFind = getMitumoriNoRange(sht)
    Set findMitumoriNo = rngFind.Find(MitumoriNo, , , xlWhole, xlByColumns, xlNext, False, False)
End Function
Function findCustomerFormat(CustomerName As String) As String
'宛名リストに該当する書式タイプを返す
    findCustomerFormat = findCustomerListData(CustomerName, 0)
End Function
Function findCustomerSeikyuuType(CustomerName As String) As String
'宛名リストに該当する請求書式を返す
    findCustomerSeikyuuType = findCustomerListData(CustomerName, 1)
End Function
Function findBumonName(strBumon As String) As Range
'strBumonと同じ内容のセルを返す
Dim rngFind As Range
    Set rngFind = getBumonNameRange
    Set findBumonName = rngFind.Find(strBumon, , , xlWhole, xlByColumns, xlNext, False, False)
End Function
Function findCustomerName(strCustomer As String) As Range
'strCustomerと同じ内容のセルを返す
Dim rngFind As Range
    Set rngFind = Worksheets("リスト").Range("a:a")
    Set findCustomerName = rngFind.Find(strCustomer, , , xlWhole, xlByColumns, xlNext, False, False)
End Function
