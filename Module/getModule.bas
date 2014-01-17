Attribute VB_Name = "getModule"
Private Function getEndRow(strColumns As String, shtMy As Worksheet) As Long
'求めたい列内で表示されている最終行を取得する
'入力がない場合は0を返す
'strColumns: "a" a列 : "c:e" c-e列
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
'rngMnoから対応する表題データを返す
    With rngMno
        getHyoudaiDataV2.strSerial = .Cells(1, 0)         '入力番号
        getHyoudaiDataV2.strMitumoriNo = .Cells(1, 1)     '見積No
        getHyoudaiDataV2.strCustomer = .Cells(1, 2)       '宛名
        getHyoudaiDataV2.dteMitumoriDay = .Cells(1, 3)      '見積日
        getHyoudaiDataV2.strFormat = .Cells(1, 4)         '書式タイプ
        getHyoudaiDataV2.strBumon = .Cells(1, 5)         '担当部門
        getHyoudaiDataV2.strSite = .Cells(1, 6)      '所在地
        getHyoudaiDataV2.strLocation = .Cells(1, 7)      '位置
        getHyoudaiDataV2.strKiHyouki = .Cells(1, 8)       '貴表記
        getHyoudaiDataV2.strName = .Cells(1, 9)      '名前
        getHyoudaiDataV2.strContents = .Cells(1, 10)      '内容
        getHyoudaiDataV2.strDeliveryPlace = .Cells(1, 11)  '納入場所
        getHyoudaiDataV2.strSiharai = .Cells(1, 12)        '支払い条件
        getHyoudaiDataV2.strYuukoukikann = .Cells(1, 13)   '有効期限
        getHyoudaiDataV2.dblProceeds = .Cells(1, 14)       '金額（税込）
        getHyoudaiDataV2.dblSum = .Cells(1, 15)            '金額(税別）
        getHyoudaiDataV2.dblCost = .Cells(1, 16)           '原価（税別）
        getHyoudaiDataV2.strNotes = .Cells(1, 17)          'メモ
        getHyoudaiDataV2.strMaker = .Cells(1, 18)          '作成者
        getHyoudaiDataV2.dteSeikyuuDay = .Cells(1, 19)       '請求日
        getHyoudaiDataV2.strSeikyuuType = .Cells(1, 20)    '請求方法
        getHyoudaiDataV2.strMsinsei = .Cells(1, 21)        '申請日時
        getHyoudaiDataV2.dblTaxRate = .Cells(1, 22)            '消費税
        getHyoudaiDataV2.strPublishRequestType = .Cells(1, 23)     '発行申請
        getHyoudaiDataV2.strMitumoriPresentDay = .Cells(1, 24)     '見積提出日
        getHyoudaiDataV2.strAccountsDate = .Cells(1, 25)       '決済日
        getHyoudaiDataV2.strCheckOfAccounts = .Cells(1, 26)     '受注確認者
        getHyoudaiDataV2.strCheckOfFinishing = .Cells(1, 27)   '完了確認日
        getHyoudaiDataV2.strWorkReport = .Cells(1, 28)         '作業報告書
        getHyoudaiDataV2.strUriageTuki = .Cells(1, 29)         '売上月
    End With
End Function
Private Function getUtiwakeDataV2(rngMnos() As Range) As UtiwakeData()
'rngMnos()から内訳詳細データを返す
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
'rngMnos()から対応する業者データを返す
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
'rngMnos()から対応する詳細データを返す
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
'配列に与えられたデータを重複がない状態にして返す
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
'行から見積Noを取得する
Dim lngRow As Long
Dim lngcolumn As Long
    Select Case shtMy.Name
    Case "表題"
        lngRow = row
        If row < 3 Then Exit Function
        lngcolumn = 2
    Case "詳細", "内訳", "業者"
        lngRow = row
        If row < 3 Then Exit Function
        lngcolumn = 1
    Case "入力"
        lngRow = 2
        lngcolumn = 6
    Case Else
        Exit Function
    End Select
    getMitumoriNoOnRow = shtMy.Cells(lngRow, lngcolumn)
End Function

Function getMitumoriNoRange(sht As Worksheet) As Range
'見積No検索用のrangeを返す
    Select Case sht.Name
    Case "表題", "定期表題"
        Set getMitumoriNoRange = sht.Range("b:b")
    Case "詳細", "内訳", "業者", "定期詳細", "定期業者"
        Set getMitumoriNoRange = sht.Range("a:a")
    End Select
End Function
Function getBumonNameRange() As Range
'担当部門検索用のrangeを返す
    Set getBumonNameRange = Sheets("担当部門").Range("a:a")
End Function
Function getDataVersion() As Long
'データバージョンを取得する
    getDataVersion = Range("data_version")
End Function
Function getMitumoriNo() As String
'シートから見積Noを取得する
Dim shtMy As Worksheet
Dim lngRow As Long
Dim lngcolumn As Long
    Set shtMy = ActiveSheet
    getMitumoriNo = getMitumoriNoOnRow(shtMy, Selection.row)
End Function
Function getHyoudaiData(MitumoriNo As String, _
                        Optional shtMy As Worksheet) As HyoudaiData
'表題データ取得のフロントエンド
Dim rngMno As Range
    If shtMy Is Nothing Then Set shtMy = ActiveWorkbook.Sheets("表題")
    Set rngMno = findMitumoriNo(MitumoriNo, shtMy)
    If rngMno Is Nothing Then Exit Function
    Select Case getDataVersion
    Case 2
        getHyoudaiData = getHyoudaiDataV2(rngMno)
    End Select
End Function
Function getTeikiHyoudaiDatas(MitumoriNo() As String, _
                              bokMy As Workbook) As HyoudaiData()
'定期作業表題データ取得
'見積NOがすべて正しいことを前提に動作しているので注意
Dim i As Long
Dim shtMy As Worksheet
Dim Hdata() As HyoudaiData
    Set shtMy = bokMy.Sheets("定期表題")
    ReDim Hdata(UBound(MitumoriNo))
    For i = 0 To UBound(MitumoriNo)
        Hdata(i) = getHyoudaiData(MitumoriNo(i), shtMy)
    Next
    getTeikiHyoudaiDatas = Hdata
End Function
Function getSyousaiData(MitumoriNo As String, _
                        Optional shtMy As Worksheet) As SyousaiData()
'詳細データ取得のフロントエンド
'データが無い場合、全データが空白の一行を返す
Dim rngMnos() As Range
Dim Sdata() As SyousaiData
    If shtMy Is Nothing Then Set shtMy = ActiveWorkbook.Sheets("詳細")
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
'業者データ取得のフロントエンド
'データが無い場合、全データが空白の一行を返す
Dim rngMnos() As Range
    If shtMy Is Nothing Then Set shtMy = ActiveWorkbook.Sheets("業者")
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
'業者データ取得のフロントエンド
'データが無い場合、全データが空白の一行を返す
Dim rngMnos() As Range
    If findMitumoriNo(MitumoriNo, ActiveWorkbook.Sheets("内訳")) Is Nothing Then
        ReDim getUtiwakeData(0)
        Exit Function
    End If
    rngMnos() = findMitumoriNumRanges(MitumoriNo, ActiveWorkbook.Sheets("内訳"))
    Select Case getDataVersion
    Case 2
        getUtiwakeData = getUtiwakeDataV2(rngMnos)
    End Select
End Function
Function getSheetInputMitumoriNo() As String
'入力シートから見積Noを取得する
Dim shtMy As Worksheet
Set shtMy = Sheets("入力")
    getSheetInputMitumoriNo = shtMy.Range("d2")
End Function
Function getSheetInputMitumoriType() As String
'入力シートから見積方法を取得する
Dim shtMy As Worksheet
Set shtMy = Sheets("入力")
    getSheetInputMitumoriType = shtMy.Range("f2")
End Function
Function getSheetInputHyoudai(bokMy As Workbook) As HyoudaiData
'入力シートから表題データを取得する
Dim shtMy As Worksheet
    Set shtMy = bokMy.Worksheets("入力")
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
'入力シートの詳細データ開始行を返す
    getSheetInputSyousaiStartRow = 17
End Function
Function getSyousaiRows() As Long
'詳細データの行数を返す
    getSyousaiRows = 17
End Function
Function getSheetInputGyousyaStartRow() As Long
'入力シートの業者データ開始行を返す
    getSheetInputGyousyaStartRow = getSheetInputSyousaiStartRow
End Function
Function getSheetInputSyousaiData(bokMy As Workbook, Mno As String) As SyousaiData()
'入力シートから詳細データを取得する
'詳細データが無い場合、見積Noが空白の一行を返す
Dim shtMy As Worksheet
Dim Sdata() As SyousaiData
Dim startRow As Long
Dim i As Long
ReDim Sdata(getSyousaiRows)
    Set shtMy = bokMy.Worksheets("入力")
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
'入力シートから業者データを取得する
Dim shtMy As Worksheet
Dim Gdata() As GyousyaData
Dim i As Long
Dim j As Long
Dim startRow As Long
    startRow = getSheetInputGyousyaStartRow
    Set shtMy = bokMy.Worksheets("入力")
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
'内訳詳細の行数を返す
    getUtiwakePageRows = 39
End Function
Function getUtiwakeDataPage(strPage As String) As Long
'内訳データのstrPageデータからページを返す
    getUtiwakeDataPage = CLng(Right(strPage, Len(strPage) - 1))
End Function
Function getSheetInputUtiwakeFeedRow(page As Long) As Long
'入力シートの内訳詳細取得開始行を返す
Dim startRow As Long
Dim pageRows As Long
    pageRows = getUtiwakePageRows
    startRow = 40
    getSheetInputUtiwakeFeedRow = startRow + ((page - 1) * (pageRows + 1)) + page
End Function
Function getSheetInputUtiwakePage(bokMy As Workbook, page As Long, Mno As String) As UtiwakeData()
'入力シートから該当ページの内訳詳細データを取得する
'ページに入力がない場合には見積番号の空の一行データを返す
Dim shtMy As Worksheet
Dim row As Long
Dim FeedRow As Long
Dim Udata() As UtiwakeData
    Set shtMy = bokMy.Worksheets("入力")
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
Set shtMy = bokMy.Sheets("表題")
    lngRow = getHyoudaiEndRow(bokMy)
    Set getHyoudaiEndCell = shtMy.Cells(lngRow, 2)
End Function
Function getHyoudaiEndRow(bokMy As Workbook) As Long
'表題シートの入力最終行を取得する
    getHyoudaiEndRow = getEndRow("b", bokMy.Sheets("表題"))
End Function
Function getSyousaiEndRow(bokMy As Workbook) As Long
'詳細シートの入力最終行を取得する
    getSyousaiEndRow = getEndRow("a", bokMy.Sheets("詳細"))
End Function
Function getGyousyaEndRow(bokMy As Workbook) As Long
'業者シートの入力最終行を取得する
    getGyousyaEndRow = getEndRow("a", bokMy.Sheets("業者"))
End Function
Function getUtiwakeEndRow(bokMy As Workbook) As Long
'内訳シートの入力最終行を取得する
    getUtiwakeEndRow = getEndRow("a", bokMy.Sheets("内訳"))
End Function
Function getDeskTopPath() As String
    Dim WSH As Variant
    Set WSH = CreateObject("WScript.Shell")
    getDeskTopPath = WSH.specialfolders("Desktop")
End Function
Function getLstMhyouki() As String()
'見積表記方法のリストを取得する
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
'請求方法のリストを取得する
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
'選択している全ての見積Noを重複なしで返す
'ただし、順番はソートしない
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
'bokMyNameのブック名以外の入力シートが存在するworkbookを配列で返す
'存在しない場合、空の配列が帰る
Dim bokThis As Workbook
Dim bokMy() As Workbook
Dim i As Long
Dim j As Long
    j = 0
    For i = 1 To Application.Workbooks.Count
        Set bokThis = Application.Workbooks(i)
        If Not bokMyName Like bokThis.Name Then
            If has_same_sheet("入力", bokThis) = True Then
                ReDim Preserve bokMy(j)
                Set bokMy(j) = bokThis
                j = j + 1
            End If
        End If
    Next
    getTorikomiWorkbooks = bokMy
End Function
Function getTorikomiWorkbook(bokMyName As String) As Workbook
'bokMyNameのブック名以外の入力シートを持つworkbookを返す
'存在しない場合、空のオブジェクトを返す
Dim bokThis As Workbook
Dim i As Long
    For i = 1 To Application.Workbooks.Count
        Set bokThis = Application.Workbooks(i)
        If Not bokMyName Like bokThis.Name Then
            If has_same_sheet("入力", bokThis) = True Then
                Set getTorikomiWorkbook = bokThis
                Exit Function
            End If
        End If
    Next
End Function
