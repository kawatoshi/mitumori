Attribute VB_Name = "WriteModule"
Option Explicit
Private Sub writeSheetInputUtiwakeData(shtMy As Worksheet, Udata() As UtiwakeData)
'内訳詳細データ書き込み
Dim page As String
Dim newPage As Long
Dim FeedRow As Long
Dim i As Long
Dim j As Long
    If Utiwake_is_empty(Udata) = True Then Exit Sub
    For i = 0 To UBound(Udata)
        If Not Udata(i).strPage Like page Then
            page = Udata(i).strPage
            newPage = getUtiwakeDataPage(Udata(i).strPage)
            FeedRow = getSheetInputUtiwakeFeedRow(newPage)
            j = 0
        End If
        shtMy.Cells(FeedRow + j, 1) = Udata(i).strHeader
        shtMy.Cells(FeedRow + j, 2) = Udata(i).strContents
        shtMy.Cells(FeedRow + j, 3) = Udata(i).strSpec
        shtMy.Cells(FeedRow + j, 4) = Udata(i).strNumber
        shtMy.Cells(FeedRow + j, 5) = Udata(i).strUnit
        shtMy.Cells(FeedRow + j, 6) = Udata(i).strPrice
        shtMy.Cells(FeedRow + j, 7) = Udata(i).strSum
        shtMy.Cells(FeedRow + j, 8) = Udata(i).strNote
        j = j + 1
    Next
End Sub
Private Function writeHyoudaiData(Hdata As HyoudaiData, bokMy As Workbook, postRow As Long) As Boolean
'表題データを書き込む
Dim shtMy As Worksheet
Set shtMy = bokMy.Sheets("表題")
    shtMy.Cells(postRow, 1) = Hdata.strSerial
    shtMy.Cells(postRow, 2) = Hdata.strMitumoriNo
    shtMy.Cells(postRow, 3) = Hdata.strCustomer
    shtMy.Cells(postRow, 4) = Hdata.dteMitumoriDay
    shtMy.Cells(postRow, 5) = Hdata.strFormat
    shtMy.Cells(postRow, 6) = Hdata.strBumon
    shtMy.Cells(postRow, 7) = Hdata.strSite
    shtMy.Cells(postRow, 8) = Hdata.strLocation
    shtMy.Cells(postRow, 9) = Hdata.strKiHyouki
    shtMy.Cells(postRow, 10) = Hdata.strName
    shtMy.Cells(postRow, 11) = Hdata.strContents
    shtMy.Cells(postRow, 12) = Hdata.strDeliveryPlace
    shtMy.Cells(postRow, 13) = Hdata.strSiharai
    shtMy.Cells(postRow, 14) = Hdata.strYuukoukikann
    shtMy.Cells(postRow, 15) = Hdata.dblProceeds
    shtMy.Cells(postRow, 16) = Hdata.dblSum
    shtMy.Cells(postRow, 17) = Hdata.dblCost
    shtMy.Cells(postRow, 18) = Hdata.strNotes
    shtMy.Cells(postRow, 19) = Hdata.strMaker
    shtMy.Cells(postRow, 20) = postStrDate(Hdata.dteSeikyuuDay)
    shtMy.Cells(postRow, 21) = Hdata.strSeikyuuType
    shtMy.Cells(postRow, 22) = Hdata.strMsinsei
    shtMy.Cells(postRow, 23) = Hdata.dblTaxRate
    shtMy.Cells(postRow, 24) = Hdata.strPublishRequestType
    shtMy.Cells(postRow, 25) = Hdata.strMitumoriPresentDay
    shtMy.Cells(postRow, 26) = Hdata.strAccountsDate
    shtMy.Cells(postRow, 27) = Hdata.strCheckOfAccounts
    shtMy.Cells(postRow, 28) = Hdata.strCheckOfFinishing
    shtMy.Cells(postRow, 29) = Hdata.strWorkReport
    shtMy.Cells(postRow, 30) = Hdata.strUriageTuki
    writeHyoudaiData = True
End Function

Function writeSheetInputData(Hdata As HyoudaiData, _
                            Sdata() As SyousaiData, _
                            Udata() As UtiwakeData, _
                            Gdata() As GyousyaData)
'入力シートに書き込む
Dim shtMy As Worksheet
Dim sRow As Long
Dim i As Long
    Set shtMy = Worksheets("入力")
    '表題データ書き込み
    shtMy.Range("b2") = Hdata.strCustomer
    shtMy.Range("d2") = Hdata.strMitumoriNo
    shtMy.Range("d5") = Hdata.strFormat
    shtMy.Range("b5") = Hdata.strBumon
    shtMy.Range("b8") = Hdata.strSite
    shtMy.Range("e8") = Hdata.strLocation
    shtMy.Range("b11") = Hdata.strName
    shtMy.Range("b14") = Hdata.strContents
    shtMy.Range("c11") = Hdata.strKiHyouki
    shtMy.Range("e14") = Hdata.strSeikyuuType
    shtMy.Range("h14") = Hdata.strYuukoukikann
    shtMy.Range("g14") = Hdata.strSiharai
    '詳細データ書き込み
    sRow = getSheetInputSyousaiStartRow
    For i = 0 To UBound(Sdata)
        shtMy.Cells(sRow + i, 1) = Sdata(i).strHeader
        shtMy.Cells(sRow + i, 2) = Sdata(i).strContents
        shtMy.Cells(sRow + i, 3) = Sdata(i).strSpec
        shtMy.Cells(sRow + i, 4) = Sdata(i).strNumber
        shtMy.Cells(sRow + i, 5) = Sdata(i).strUnit
        shtMy.Cells(sRow + i, 6) = Sdata(i).strPrice
        shtMy.Cells(sRow + i, 7) = Sdata(i).strSum
        shtMy.Cells(sRow + i, 8) = Sdata(i).strNote
    Next
    '業者データ書き込み
    sRow = getSheetInputGyousyaStartRow
    For i = 0 To UBound(Gdata)
        shtMy.Cells(sRow + i, 10) = Gdata(i).strGyousya
        shtMy.Cells(sRow + i, 11) = Gdata(i).strCost
    Next
    '内訳詳細データ書き込み
    Call writeSheetInputUtiwakeData(shtMy, Udata())
    
End Function
Function writeNewHyoudaiData(Hdata As HyoudaiData, bokMy As Workbook) As Boolean
'新規表題データを書き込む
    writeNewHyoudaiData = writeHyoudaiData(Hdata, bokMy, getHyoudaiEndRow(bokMy) + 1)
End Function
Function reWriteHyoudaiData(Hdata As HyoudaiData, bokMy As Workbook) As Boolean
'表題データを書き戻す
Dim rngMy As Range
    Set rngMy = findMitumoriNo(Hdata.strMitumoriNo, Sheets("表題"))
    If rngMy Is Nothing Then Exit Function
    reWriteHyoudaiData = writeHyoudaiData(Hdata, bokMy, rngMy.row)
End Function
Function reWriteHyoudaiWithRequest(Hdata As HyoudaiData, _
                                 bokMy As Workbook, _
                                 RequestType As String, _
                                 MitumoriType As String, _
                                 SeikyuuType As String) As Boolean
'申請書類を変更する
    Hdata.strPublishRequestType = RequestType
    Hdata.strFormat = MitumoriType
    Hdata.strSeikyuuType = SeikyuuType
    If Hdata.dteSeikyuuDay < 1 Then
        Hdata.dteSeikyuuDay = postSeikyuuDate(Hdata)
    End If
    reWriteHyoudaiWithRequest = reWriteHyoudaiData(Hdata, bokMy)
End Function
Function writeNewSyousaiData(Sdata() As SyousaiData, bokMy As Workbook) As Boolean
'新規詳細データを書き込む
Dim postRow As Long
Dim i As Long
Dim shtMy As Worksheet
Set shtMy = bokMy.Sheets("詳細")
    If Sdata(0).strMitumoriNo Like "" Then Exit Function
    postRow = getSyousaiEndRow(bokMy) + 1
    For i = 0 To UBound(Sdata)
        shtMy.Cells(postRow, 1) = Sdata(i).strMitumoriNo
        shtMy.Cells(postRow, 2) = Sdata(i).strHeader
        shtMy.Cells(postRow, 3) = Sdata(i).strContents
        shtMy.Cells(postRow, 4) = Sdata(i).strSpec
        shtMy.Cells(postRow, 5) = Sdata(i).strNumber
        shtMy.Cells(postRow, 6) = Sdata(i).strUnit
        shtMy.Cells(postRow, 7) = Sdata(i).strPrice
        shtMy.Cells(postRow, 8) = Sdata(i).strSum
        shtMy.Cells(postRow, 9) = Sdata(i).strNote
        postRow = postRow + 1
    Next
    writeNewSyousaiData = True
End Function
Function writeNewGyousyaData(Gdata() As GyousyaData, bokMy As Workbook) As Boolean
'新規業者データを書き込む
Dim postRow As Long
Dim i As Long
Dim shtMy As Worksheet
Set shtMy = bokMy.Sheets("業者")
    If Gdata(0).strMitumoriNo Like "" Then Exit Function
    postRow = getGyousyaEndRow(bokMy) + 1
    For i = 0 To UBound(Gdata)
        shtMy.Cells(postRow, 1) = Gdata(i).strMitumoriNo
        shtMy.Cells(postRow, 2) = Gdata(i).strGyousya
        shtMy.Cells(postRow, 3) = Gdata(i).strCost
        shtMy.Cells(postRow, 4) = Gdata(i).strCostWithTax
        shtMy.Cells(postRow, 5) = Gdata(i).strBillMonth
        postRow = postRow + 1
    Next
    writeNewGyousyaData = True
End Function
Function writeNewUtiwakeData(Udata() As UtiwakeData, bokMy As Workbook) As Boolean
'新規内訳データを書き込む
Dim postRow As Long
Dim i As Long
Dim shtMy As Worksheet
Set shtMy = bokMy.Sheets("内訳")
    If Udata(0).strMitumoriNo Like "" Then Exit Function
    postRow = getUtiwakeEndRow(bokMy) + 1
    For i = 0 To UBound(Udata)
        shtMy.Cells(postRow, 1) = Udata(i).strMitumoriNo
        shtMy.Cells(postRow, 2) = Udata(i).strHeader
        shtMy.Cells(postRow, 3) = Udata(i).strContents
        shtMy.Cells(postRow, 4) = Udata(i).strSpec
        shtMy.Cells(postRow, 5) = Udata(i).strNumber
        shtMy.Cells(postRow, 6) = Udata(i).strUnit
        shtMy.Cells(postRow, 7) = Udata(i).strPrice
        shtMy.Cells(postRow, 8) = Udata(i).strSum
        shtMy.Cells(postRow, 9) = Udata(i).strNote
        shtMy.Cells(postRow, 10) = Udata(i).strPage
        postRow = postRow + 1
    Next
    writeNewUtiwakeData = True
End Function
Function writeSinseiData(bokMy As Workbook, Hdata As HyoudaiData) As Boolean
'本部申請時のデータ更新書き込みを行う
    Call postSumAndCost(bokMy, Hdata)
    writeSinseiData = reWriteHyoudaiData(Hdata, bokMy)
End Function
