Attribute VB_Name = "PublishModule"
Option Explicit
Private Function postKenmei(Hdata As HyoudaiData) As String
Dim ans As String
    ans = Trim(Hdata.strSite & " " & Hdata.strLocation)
    If Not Hdata.strKiHyouki Like "無" Then
        ans = ans & "貴『" & Hdata.strName & "』"
    Else
        ans = ans & " " & Hdata.strName
    End If
    ans = ans & " " & Hdata.strContents
    postKenmei = Trim(ans)
End Function
Private Function postTaxInfo(Hdata As HyoudaiData) As String
Dim ans As String
    If SinseiType_is_zeikomi(Hdata.strFormat) = True Then
        postTaxInfo = ""
    Else
        ans = CStr(Hdata.dblTaxRate * 100)
        postTaxInfo = "上記に関わる消費税(" & ans & "％)"
    End If
End Function
Private Function postTaxFx(Hdata As HyoudaiData) As String
    If SinseiType_is_zeikomi(Hdata.strFormat) = True Then
        postTaxFx = ""
    Else
        postTaxFx = "=ROUNDDOWN(J35*" & CStr(Hdata.dblTaxRate) & ",0)"
    End If
End Function
Private Function postTaxType(strTaxType As String) As String
    If SinseiType_is_zeikomi(strTaxType) = True Then
        postTaxType = "(消費税内税)"
    Else
        postTaxType = "(税込)"
    End If
End Function
Private Function getBumonSignature(strBumon As String) As Range
'(見積・請求用)担当部門の署名を返す
Dim rngMy As Range
Dim strMy(3) As String
Dim i As Long
    If strBumon Like "" Then
        Exit Function
    End If
    Set rngMy = findBumonName(strBumon)
    If rngMy Is Nothing Then
        Exit Function
    End If
    Set getBumonSignature = Range(rngMy.Cells(1, 2), rngMy.Cells(4, 2))
End Function
Private Function getFurikomiSaki(SeikyuuType As String) As String()
'請求振込先の配列を返す
Dim rngFurikomiSaki As Range
Dim strType As String
Dim rngMy As Range
Dim i As Long
Dim strMy() As String
    Set rngFurikomiSaki = Sheets("担当部門").Range("g:g")
    Select Case SeikyuuType
    Case "", "官庁"
        strType = "振込"
    Case Else
        strType = SeikyuuType
    End Select
    Set rngMy = rngFurikomiSaki.Find(strType, , , xlWhole, xlByColumns, xlNext, False, False)
    If rngMy Is Nothing Then
        ReDim getFurikomiSaki(0): Exit Function
    End If
    ReDim strMy(3)
    For i = 0 To UBound(strMy)
        strMy(i) = rngMy.Cells(i + 1, 2)
    Next
    getFurikomiSaki = strMy
End Function
Private Sub writeSignature(DFormat As String, Dbumon As String, sht As Worksheet)
Dim Bs As Range
Dim i As Long
    If is_match("官庁", DFormat) = True Then
        sht.Range("j3:m11").ClearContents
    End If
    If is_match("日付なし", DFormat) = True Then
        sht.Range("k2") = ""
    End If
    If is_match("日付空白", DFormat) = True Then
        sht.Range("k2") = "平成　　年　　月　　日"
    End If
    If is_match("官庁", DFormat) = False Then
        Set Bs = getBumonSignature(Dbumon)
        If Not Bs Is Nothing Then
            sht.Cells(7, 10) = "(担当部門)"
            Bs.Cells.Copy Destination:=sht.Cells(8, 10)
        End If
    End If
End Sub
Private Sub writeFurikomisaki(Hdata As HyoudaiData, dteFurikomi As Date, shtDst As Worksheet)
 '請求振込先の書き込み
Dim strFurikomi() As String
Dim i As Long
    strFurikomi = getFurikomiSaki(Hdata.strSeikyuuType)
    For i = 0 To UBound(strFurikomi)
        shtDst.Cells(11 + i, 2) = strFurikomi(i)
    Next
    Select Case Hdata.strSeikyuuType
    Case "", "振込"
        If is_match("官庁", Hdata.strFormat) = False Then
            Sheets("担当部門").Range("k3:l4").Copy
            shtDst.Range("f13").PasteSpecial
        End If
    Case "控除"
        shtDst.Range("b12") = repMonth(shtDst.Range("b12").Value, dteFurikomi)
    End Select
End Sub
Private Function writeHyoudaiBase(Hdata As HyoudaiData, sht As Worksheet)
    sht.Range("k1") = postNO(Hdata.strMitumoriNo)
    sht.Range("b3") = postCustomerName(Hdata.strCustomer) & " 様"
    sht.Range("d4") = postKenmei(Hdata)
    sht.Range("b36") = postTaxInfo(Hdata)
    sht.Range("j36") = postTaxFx(Hdata)
    Call writeSignature(Hdata.strFormat, Hdata.strBumon, sht)
End Function
Private Function writeMitumoriHyoudai(Hdata As HyoudaiData, sht As Worksheet) As Boolean
'見積表題部の書き込み
    If Hdata.strMitumoriNo Like "" Then Exit Function
    sht.Range("k2") = Hdata.dteMitumoriDay
    Call writeHyoudaiBase(Hdata, sht)
    sht.Range("d12") = Hdata.strDeliveryPlace
    sht.Range("d13") = Hdata.strSiharai
    sht.Range("d14") = Hdata.strYuukoukikann
    sht.Range("e9") = postTaxType(Hdata.strFormat)
    writeMitumoriHyoudai = True
End Function
Private Function writeSeikyuuHyoudai(Hdata As HyoudaiData, sht As Worksheet) As Boolean
'請求表題部の書き込み
    If Hdata.strMitumoriNo Like "" Then Exit Function
    sht.Range("k2") = Hdata.dteSeikyuuDay
    sht.Range("e8") = postTaxType(Hdata.strFormat)
    Call writeHyoudaiBase(Hdata, sht)
    Call writeFurikomisaki(Hdata, Hdata.dteSeikyuuDay, sht)
    writeSeikyuuHyoudai = True
End Function
Private Function writeSyousai(Sdata() As SyousaiData, sht As Worksheet) As Boolean
Dim rngS As Range
Dim i As Long
Dim j As Long
Dim Syousai As Variant
Set rngS = sht.Range("b17:m34")
    If Sdata(0).strMitumoriNo Like "" Then Exit Function
    i = 1
        With rngS
            For i = 0 To UBound(Sdata)
                .Cells(i + 1, 1) = Sdata(i).strHeader
                .Cells(i + 1, 2) = Sdata(i).strContents
                .Cells(i + 1, 4) = Sdata(i).strSpec
                .Cells(i + 1, 6) = Sdata(i).strNumber
                .Cells(i + 1, 7) = Sdata(i).strUnit
                .Cells(i + 1, 8) = Sdata(i).strPrice
                .Cells(i + 1, 9) = Sdata(i).strSum
                .Cells(i + 1, 11) = Sdata(i).strNote
            Next
        End With
    writeSyousai = True
End Function
Private Function getMitumoriUtiwakeStartRow(page As Long) As Long
'見積の詳細内訳データ開始行を返す
    getMitumoriUtiwakeStartRow = 40 + ((page - 1) * 43)
End Function
Private Function writeMitumoriUtiwake(Udata() As UtiwakeData, _
                                      shtDst As Worksheet) As Boolean
'見積の内訳詳細データ書き込み
Dim page As String
Dim pages As Long
Dim newPage As Long
Dim FeedRow As Long
Dim i As Long
Dim j As Long
Dim shtSrc As Worksheet
Set shtSrc = Sheets("内訳原紙")
    If Utiwake_is_empty(Udata) = True Then Exit Function
    pages = getUtiwakeDataPage(Udata(UBound(Udata)).strPage)
    Call initUtiwakeSyosiki(shtSrc, shtDst, pages)
    For i = 0 To UBound(Udata)
        If Not Udata(i).strPage Like page Then
            page = Udata(i).strPage
            newPage = getUtiwakeDataPage(Udata(i).strPage)
            FeedRow = getMitumoriUtiwakeStartRow(newPage)
            shtDst.Cells(FeedRow - 2, 13) = postNO(Udata(i).strMitumoriNo)
            shtDst.Cells(FeedRow + 40, 13) = "Page" & newPage & "/" & pages
            j = 0
        End If
        shtDst.Cells(FeedRow + j, 2) = Udata(i).strHeader
        shtDst.Cells(FeedRow + j, 3) = Udata(i).strContents
        shtDst.Cells(FeedRow + j, 5) = Udata(i).strSpec
        shtDst.Cells(FeedRow + j, 7) = Udata(i).strNumber
        shtDst.Cells(FeedRow + j, 8) = Udata(i).strUnit
        shtDst.Cells(FeedRow + j, 9) = Udata(i).strPrice
        shtDst.Cells(FeedRow + j, 10) = Udata(i).strSum
        shtDst.Cells(FeedRow + j, 12) = Udata(i).strNote
        j = j + 1
    Next
    writeMitumoriUtiwake = True
End Function

Function publishMitumori(MitumoriNo As String, _
                         shtSrc As Worksheet, _
                         shtDst As Worksheet) As Boolean
'見積書を発行する
    shtSrc.Parent.Activate
    Call initSyosiki(shtSrc, shtDst)
    If writeMitumoriHyoudai(getHyoudaiData(MitumoriNo), shtDst) = False Then Exit Function
    If writeSyousai(getSyousaiData(MitumoriNo), shtDst) = False Then Exit Function
    Call writeMitumoriUtiwake(getUtiwakeData(MitumoriNo), shtDst)
    publishMitumori = True
End Function
Function publishSeikyuu(MitumoriNo As String, _
                        shtSrc As Worksheet, _
                        shtDst As Worksheet) As Boolean
'請求書を発行する
Dim Hdata As HyoudaiData
    shtSrc.Parent.Activate
    Hdata = getHyoudaiData(MitumoriNo)
    If Hdata.dteSeikyuuDay <= 0 Then
        Hdata.dteSeikyuuDay = postSeikyuuDate(Hdata)
    End If
    Call initSyosiki(shtSrc, shtDst)
    If writeSeikyuuHyoudai(Hdata, shtDst) = False Then Exit Function
    If writeSyousai(getSyousaiData(MitumoriNo), shtDst) = False Then Exit Function
    Call reWriteHyoudaiData(Hdata, shtSrc.Parent)
    publishSeikyuu = True
End Function

