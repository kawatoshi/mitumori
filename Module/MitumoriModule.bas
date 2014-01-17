Attribute VB_Name = "MitumoriModule"
Option Explicit

Private Function getMitumoriSection() As String
'見積Noの営業所を取得する
    getMitumoriSection = Range("mitumori_head")
End Function
Private Function getMitumoriYear() As String
'見積Noの年度を取得する
    getMitumoriYear = CStr(year(Now()) - 1988)
End Function
Private Function getMitumoriBasho() As String
'見積Noの場所コードを取得する
    getMitumoriBasho = Range("basho")
End Function
Private Function postMitumoriSerial() As String
'見積Noのシリアルを作成する
Dim lngSerial As Long
Dim strSerial As String

    lngSerial = Range("serial")
    strSerial = CStr(lngSerial + 1)
    Select Case Len(strSerial)
    Case 1
        strSerial = "000" & strSerial
    Case 2
        strSerial = "00" & strSerial
    Case 3
        strSerial = "0" & strSerial
    Case 4
    Case Else
        strSerial = "a" & Right(strSerial, 3)
    End Select
    postMitumoriSerial = strSerial
End Function

Function postNewMitumoriNo() As MitumoriNumber
'見積Noを新規で自動作成する
Set postNewMitumoriNo = New MitumoriNumber
    Call postNewMitumoriNo.Push(getMitumoriYear & _
                                getMitumoriSection & "-" & _
                                postMitumoriSerial)
End Function
Function postMaxSaiMitumoriNo(MitumoriNo As String) As MitumoriNumber
'引数に与えられた見積Noの再見積Noを作成する
Dim Mno As MitumoriNumber
Dim rngNo() As Range
Dim rngMy As Variant
Dim maxMno As String
Dim shtH As Worksheet
Set Mno = New MitumoriNumber
Set postMaxSaiMitumoriNo = New MitumoriNumber
Set shtH = ActiveWorkbook.Sheets("表題")
    If Mno.Push(MitumoriNo) = False Then Exit Function
    If findMitumoriNo(MitumoriNo, shtH) Is Nothing Then Exit Function
    rngNo() = findMitumoriNumRanges(Mno.MainNo & "*", shtH)
    maxMno = ""
    For Each rngMy In rngNo
        If maxMno < rngMy.Value Then
            maxMno = rngMy.Value
        End If
    Next
    postMaxSaiMitumoriNo.Push (maxMno)
    postMaxSaiMitumoriNo.NextNumber
End Function
Function postTeikiMitumoriNo(MitumoriNo As String) As MitumoriNumber
'引数に与えられた見積Noの定期見積No作成する
Set postTeikiMitumoriNo = New MitumoriNumber
    Call postTeikiMitumoriNo.Push(MitumoriNo)
    Call postTeikiMitumoriNo.to_teiki
End Function
Function postMitumoriNo(MitumoriNo As String, MitumoriType As String) As String
'見積Noと見積typeから新規見積Noを作成する
    Select Case MitumoriType
    Case "新規"
        postMitumoriNo = postNewMitumoriNo.Publish
    Case "再見積"
        postMitumoriNo = postMaxSaiMitumoriNo(MitumoriNo).Publish
    Case "定期"
        postMitumoriNo = postTeikiMitumoriNo(MitumoriNo).Publish
    Case Else
        postMitumoriNo = postNewMitumoriNo.Publish
    End Select
End Function

