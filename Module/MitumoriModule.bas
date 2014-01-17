Attribute VB_Name = "MitumoriModule"
Option Explicit

Private Function getMitumoriSection() As String
'����No�̉c�Ə����擾����
    getMitumoriSection = Range("mitumori_head")
End Function
Private Function getMitumoriYear() As String
'����No�̔N�x���擾����
    getMitumoriYear = CStr(year(Now()) - 1988)
End Function
Private Function getMitumoriBasho() As String
'����No�̏ꏊ�R�[�h���擾����
    getMitumoriBasho = Range("basho")
End Function
Private Function postMitumoriSerial() As String
'����No�̃V���A�����쐬����
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
'����No��V�K�Ŏ����쐬����
Set postNewMitumoriNo = New MitumoriNumber
    Call postNewMitumoriNo.Push(getMitumoriYear & _
                                getMitumoriSection & "-" & _
                                postMitumoriSerial)
End Function
Function postMaxSaiMitumoriNo(MitumoriNo As String) As MitumoriNumber
'�����ɗ^����ꂽ����No�̍Č���No���쐬����
Dim Mno As MitumoriNumber
Dim rngNo() As Range
Dim rngMy As Variant
Dim maxMno As String
Dim shtH As Worksheet
Set Mno = New MitumoriNumber
Set postMaxSaiMitumoriNo = New MitumoriNumber
Set shtH = ActiveWorkbook.Sheets("�\��")
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
'�����ɗ^����ꂽ����No�̒������No�쐬����
Set postTeikiMitumoriNo = New MitumoriNumber
    Call postTeikiMitumoriNo.Push(MitumoriNo)
    Call postTeikiMitumoriNo.to_teiki
End Function
Function postMitumoriNo(MitumoriNo As String, MitumoriType As String) As String
'����No�ƌ���type����V�K����No���쐬����
    Select Case MitumoriType
    Case "�V�K"
        postMitumoriNo = postNewMitumoriNo.Publish
    Case "�Č���"
        postMitumoriNo = postMaxSaiMitumoriNo(MitumoriNo).Publish
    Case "���"
        postMitumoriNo = postTeikiMitumoriNo(MitumoriNo).Publish
    Case Else
        postMitumoriNo = postNewMitumoriNo.Publish
    End Select
End Function

