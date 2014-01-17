Attribute VB_Name = "ReplaceModule"
Option Explicit

Function replace(strSrc As String, strPattern As String, strReplace As String) As String
Dim RegEx As Variant
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.pattern = strPattern
    RegEx.ignorecase = False
    RegEx.Global = True
    replace = RegEx.replace(strSrc, strReplace)
End Function
Function repZM(strSrc As String, dteSrc As Date) As String
Dim strMonth As String
Dim pattern As String
    pattern = "%zm"
    If is_match(pattern, strSrc) = True Then
        strMonth = StrConv(CStr(Month(dteSrc)), vbWide)
        repZM = replace(strSrc, pattern, strMonth)
        Exit Function
    End If
    repZM = strSrc
End Function
Function repM(strSrc As String, dteSrc As Date) As String
Dim strMonth As String
Dim pattern As String
    pattern = "%m"
    If is_match(pattern, strSrc) = True Then
        strMonth = StrConv(CStr(Month(dteSrc)), vbNarrow)
        repM = replace(strSrc, pattern, strMonth)
        Exit Function
    End If
    repM = strSrc
End Function
Function repMonth(strSrc As String, dteSrc As Date) As String
Dim strText As String
    strText = repZM(strSrc, dteSrc)
    repMonth = repM(strText, dteSrc)
End Function
Sub replaceTeikiHyoudaiFormat(Hdata As HyoudaiData)
'定期表題データのフォーマット変換を行う
    Hdata.strLocation = repMonth(Hdata.strLocation, Hdata.dteSeikyuuDay)
    Hdata.strName = repMonth(Hdata.strName, Hdata.dteSeikyuuDay)
    Hdata.strContents = repMonth(Hdata.strContents, Hdata.dteSeikyuuDay)
End Sub
Sub replaceTeikiSyousaiFormat(Sdata() As SyousaiData, dteMy As Date)
'定期詳細データのフォーマット変換を行う
Dim i As Long
    For i = 0 To UBound(Sdata)
        Sdata(i).strContents = repMonth(Sdata(i).strContents, dteMy)
        Sdata(i).strSpec = repMonth(Sdata(i).strSpec, dteMy)
        Sdata(i).strNote = repMonth(Sdata(i).strNote, dteMy)
    Next
End Sub
