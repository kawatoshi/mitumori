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
'引数に与えられたシートにオートフィルターによる非表示行が
'あるかを返す
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
    SheetNames = Split("表題,詳細,内訳,業者", ",")
    For Each sName In SheetNames
        If chkFilter(CStr(sName)) = True Then
            chkAllSheetFilter = CStr(sName) & "にフィルターがかかっています。"
            Exit Function
        End If
    Next
    chkAllSheetFilter = ""
End Function
Function Utiwake_is_empty(Udata() As UtiwakeData) As Boolean
'与えられた内訳にデータが含まれていないかを返す
    If Udata(0).strMitumoriNo Like "" Then
        Utiwake_is_empty = True
    Else
        Utiwake_is_empty = False
    End If
End Function
Function same_MitumoriNo(MitumoriNo As String) As Boolean
'与えられた見積Noが表題に存在するかを返す
Dim shtMy As Worksheet
Set shtMy = Sheets("表題")
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
    Case "新規"
        If MitumoriNo Like "" Then Exit Function
        ans = "見積Noが入力されています。新規の場合は見積Noに何も記入しないでください"
    Case "再見積"
        If Mno.Push(MitumoriNo) = False Then ans = "有効な見積Noではありません"
        If MitumoriNo Like "" Then ans = "見積Noが入力されていません。有効な見積Noを記入してください"
        If same_MitumoriNo(Mno.MainNo & "*") = False Then ans = "再見積に必要なデータが見つかりません" & _
                                                            Chr(13) & "表題シートにデータのある見積番号が必要です"
    Case "定期"
        If Mno.Push(MitumoriNo) = False Then ans = "有効な見積Noではありません"
        If MitumoriNo Like "" Then ans = "見積Noが入力されていません。有効な見積Noを記入してください"
        If same_MitumoriNo(Mno.MainNo & "*") = False Then ans = "定期見積に必要なデータが見つかりません" & _
                                                            Chr(13) & "表題シートにデータのある見積番号が必要です"
    Case "転記"
        If Mno.Push(MitumoriNo) = False Then ans = "有効な見積Noではありません"
        If MitumoriNo Like "" Then ans = "見積Noが入力されていません。有効な見積Noを記入してください"
        If same_MitumoriNo(Mno.Publish) = True Then ans = "すでにその見積Noは使用されています"
    Case Else
        ans = "見積タイプが不明です"
    End Select
    chkInputSheetMitumoriNumber = ans
End Function
Function is_match_Mitumori(strTxt As String) As Boolean
    is_match_Mitumori = is_match("見積", strTxt)
End Function
Function SinseiType_is_Mitumori(MitumoriNo As String) As Boolean
Dim Hdata As HyoudaiData
    Hdata = getHyoudaiData(MitumoriNo)
    If Hdata.strMitumoriNo Like "" Then Exit Function
    SinseiType_is_Mitumori = is_match_Mitumori(Hdata.strPublishRequestType)
End Function
Function is_match_Seikyuu(strTxt As String) As Boolean
    is_match_Seikyuu = is_match("請求", strTxt)
End Function
Function SinseiType_is_Seikyuu(MitumoriNo As String) As Boolean
Dim Hdata As HyoudaiData
    Hdata = getHyoudaiData(MitumoriNo)
    If Hdata.strMitumoriNo Like "" Then Exit Function
    SinseiType_is_Seikyuu = is_match_Seikyuu(Hdata.strPublishRequestType)
End Function
Function SinseiType_is_zeikomi(strFormatOrPublishRequestType As String) As Boolean
'見積または請求方法が「税込」を含んでいるかを返す
    SinseiType_is_zeikomi = is_match("税込", strFormatOrPublishRequestType)
End Function
Function Book_is_opend(BokName As String) As Workbook
'同名のブックが開かれていれば、そのブックを返す
Dim bokMy As Variant
    For Each bokMy In Application.Workbooks
        If bokMy.Name Like BokName Then
            Set Book_is_opend = bokMy
            Exit Function
        End If
    Next
End Function
Function has_same_sheet(SheetName As String, bokCheck As Workbook) As Boolean
'sheetnameと同じシートが存在するかを返す
Dim shtMy As Worksheet
    For Each shtMy In bokCheck.Worksheets
        If shtMy.Name Like SheetName Then
            has_same_sheet = True
            Exit Function
        End If
    Next
End Function
