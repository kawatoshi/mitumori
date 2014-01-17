Attribute VB_Name = "MainModule"
Option Explicit

Function rekeyMitumori(MitumoriNo As String) As String
'mitumorinoの見積データを入力シートへ転記する
'問題が発生した場合は返値にエラーを記載する
'正常に終了した場合は空文字を返す
Dim Mno As MitumoriNumber
Dim Hdata As HyoudaiData
Dim Sdata() As SyousaiData
Dim Gdata() As GyousyaData
Dim Udata() As UtiwakeData
Set Mno = New MitumoriNumber
    If Mno.Push(MitumoriNo) = False Then
        rekeyMitumori = "見積Noが正しくありません": Exit Function
    End If
    Call initInputSheet
    Hdata = getHyoudaiData(MitumoriNo)
    Sdata = getSyousaiData(MitumoriNo)
    Udata = getUtiwakeData(MitumoriNo)
    Gdata = getGyousyaData(MitumoriNo)
    rekeyMitumori = writeSheetInputData(Hdata, Sdata, Udata, Gdata)
End Function
Function SheetInput(bokMy As Workbook) As Boolean
'入力シートから必要情報を取り込み、見積Noをを割り振った後に
'各シートへ転記する
Dim ans As String
Dim Mno As String
Dim Hdata As HyoudaiData
Dim Sdata() As SyousaiData
Dim Udata() As UtiwakeData
Dim Gdata() As GyousyaData
Dim i As Long
'オートフィルターによる非表示のチェック
    ans = chkAllSheetFilter
    If Not ans Like "" Then
        MsgBox (ans & " 処理を中断します")
        Exit Function
    End If
    Set bokMy = ActiveWorkbook
'表題データ取得
    Hdata = getSheetInputHyoudai(bokMy)
    ans = chkInputSheetMitumoriNumber(Hdata.strMitumoriNo, getSheetInputMitumoriType)
    If Not ans Like "" Then
        MsgBox (ans)
        Exit Function
    End If
'見積Noの作成
    Mno = postMitumoriNo(Hdata.strMitumoriNo, getSheetInputMitumoriType)
    Hdata.strMitumoriNo = Mno
    ans = varidateHyoudaiData(Hdata)
    If Not ans Like "" Then
        MsgBox (ans)
        Exit Function
    End If
    If Hdata.strMitumoriNo Like "" Then
        MsgBox ("見積Noが作成できませんでした" & Chr(13) & "管理者に相談してください"): Exit Function
    End If
    If getSheetInputMitumoriType Like "新規" Then Hdata.strSerial = CStr(Range("serial") + 1)
    If Hdata.dteMitumoriDay = 0 Then Hdata.dteMitumoriDay = Now()
    Hdata.strFormat = postHdataFormat(Hdata)
    Hdata.strSeikyuuType = postHdataSeikyuuType(Hdata)
    Hdata.dteSeikyuuDay = postSeikyuuDate(Hdata)
'詳細データ取得
    Sdata = getSheetInputSyousaiData(bokMy, Mno)
'業者データ取得
    Gdata = getSheetInputGyousyaData(bokMy, Mno)
'表題データ書き込み
    Call writeNewHyoudaiData(Hdata, bokMy)
'詳細データ書き込み
    Call writeNewSyousaiData(Sdata, bokMy)
'業者データ書き込み
    Call writeNewGyousyaData(Gdata, bokMy)
'内訳詳細データ取り込み及び書き込み
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
'申請を確認する
Dim shtSrc As Worksheet
Dim shtDst As Worksheet
Dim Hdata As HyoudaiData
Dim shtMy As Worksheet
Set shtMy = ActiveWorkbook.ActiveSheet
    If MitumoriNo Like "" Then Exit Function
    If SinseiType_is_Mitumori(MitumoriNo) = True Then
        Set shtSrc = Sheets("見積原紙")
        Set shtDst = Sheets("見積書")
        If publishMitumori(MitumoriNo, shtSrc, shtDst) = False Then Exit Function
    End If
    If SinseiType_is_Seikyuu(MitumoriNo) = True Then
        Set shtSrc = Sheets("請求原紙")
        Set shtDst = Sheets("請求書")
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
'捺印依頼を作成する
'strPrinterで印刷するプリンターの切り替え
'strBokNameがある場合はワークブックへの出力
'signatureは印鑑を印刷するかどうか
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
    Set shtCommission = ActiveWorkbook.Sheets("捺印依頼書")
    DefaultPrinter = getPrinterName
    OutPutPrinter = findPrinter(strPrinter)
'申請の発行
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
'引数見積Noの見積書または請求書を作成し
'本部指定フォルダの指定ファイルにコピー、保存する
'引数として処理内容と処理結果をStringで返す
Dim bokMy As Workbook
Dim bokDst As Workbook
Dim Hdata As HyoudaiData
Dim strReqFolder As String, strReqFile As String '申請ディレクトリ、申請ファイル名
Dim ans As String
    Set bokMy = ActiveWorkbook
    Hdata = getHyoudaiData(Mno)
    If Hdata.strMitumoriNo Like "" Then
        Call MsgBox("有効な見積ではありません")
        Exit Sub
    End If
'見積申請
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
'請求申請
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
'他ブックの入力シートからの取り込み
'単独取り込み用で、取り込み結果を返す
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
    If Not has_same_sheet("入力", bokMy) = True Then
        TorikomiSheetInput = "入力シートで実行してください"
        Exit Function
    End If
    If getTorikomiWorkbook(bokMy.Name) Is Nothing Then
        TorikomiSheetInput = "取り込むブックがありません"
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
'指定月の定期データを表題、詳細、業者へ転記する
Dim strMno() As String
Dim Hdata() As HyoudaiData
Dim i As Long
Dim Hd As HyoudaiData
Dim Sdata() As SyousaiData
Dim Gdata() As GyousyaData
    strMno = findTeikiMitumoriNumbers(lngMonth, bokSrc)
    If UBound(strMno) < 0 Then
        makeTeiki = CStr(lngMonth) & "月に該当する定期データはありません"
        Exit Function
    End If
    Hdata = getTeikiHyoudaiDatas(strMno, bokSrc)
    For i = 0 To UBound(Hdata)
        Hd = postTeikiToHyoudaiData(Hdata(i), CStr(lngMonth))
        If same_MitumoriNo(Hd.strMitumoriNo) = True Then
            makeTeiki = makeTeiki & "未作成 : " & Hd.strMitumoriNo & Chr(13)
        Else
            Sdata = getSyousaiData(Hdata(i).strMitumoriNo, bokSrc.Sheets("定期詳細"))
            Gdata = getGyousyaData(Hdata(i).strMitumoriNo, bokSrc.Sheets("定期業者"))
            Call replaceTeikiHyoudaiFormat(Hd)
            Call replaceTeikiSyousaiFormat(Sdata, Hd.dteSeikyuuDay)
            Call writeNewHyoudaiData(Hd, bokSrc)
            Call writeNewSyousaiData(postTeikiToSyousaiData(Sdata, CStr(lngMonth)), bokSrc)
            Call writeNewGyousyaData(postTeikiToGyousyaData(Gdata, CStr(lngMonth)), bokSrc)
            makeTeiki = makeTeiki & "作成 : " & Hd.strMitumoriNo & Chr(13)
        End If
    Next
    makeTeiki = Trim(makeTeiki)
End Function
