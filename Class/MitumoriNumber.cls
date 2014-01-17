VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MitumoriNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'見積・請求書の番号ルールについて
'
'23K-10001-b-2312
'
'23・・・・・年度(nendo)
'K-01 ・・・営業所名と営業所固有番号(eigyousyo_code, eigyousyo_num)
'0001・・・・見積書番号(serial_num)
'
'再見積の場合には　[a-z］の記号を付加する(sai_num)
'
'定期作業請求については､
'-1105
'2桁の和暦と2桁の月の数列をハイフンの後に付加する(teiki_year, teiki_month)
'→元号にするか西暦にするかを要検討
'
'※清掃部門の見積はどうするか？

Private m_Nendo As String
Private m_Eigyousyo_code As String
Private m_eigyousyo_num As String
Private m_Serial_num As String
Private m_Sai_num As String
Private m_Teiki_year As String
Private m_Teiki_month As String
Private m_source As String
Private Message As String

Private Function pat_old_Mitumori() As String
    pat_old_Mitumori = "^(\d{2})(K|H|M|I|U|SG|AG|SK|A|F|)-(\d{2})(\d{4})(-|)([1-9]|)(-\d{2}-\d{2}|)$"
End Function
Private Function is_old_type(MitumoriNo As String) As Boolean
'見積Noが旧タイプか判定する
Dim RegEx As Variant
Dim pat As String
Dim Matches As Variant
    pat = pat_old_Mitumori
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .pattern = pat
        .ignorecase = False
        Set Matches = .Execute(MitumoriNo)
        If Matches.Count = 0 Then
            is_old_type = False
            Exit Function
        End If
    End With
    is_old_type = True
End Function
Private Function old_to_modern_no(MitumoriNo As String) As String
'旧タイプの見積Noを現行方式のNoへ変換する
'変換できない場合は引数をそのまま帰す
Dim RegEx As Variant
Dim Matches As Variant
Dim ans As String
Dim sai As String
Dim teiki As String
    If is_old_type(MitumoriNo) = True Then
        Set RegEx = CreateObject("VBScript.RegExp")
        With RegEx
            .pattern = pat_old_Mitumori
            .ignorecase = False
            Set Matches = .Execute(MitumoriNo)
            ans = Matches(0).submatches(0)
            ans = ans & Matches(0).submatches(1)
            ans = ans & "-" & Matches(0).submatches(3)
            teiki = Matches(0).submatches(6)
            If Not Matches(0).submatches(5) Like "" Then
                sai = convert_sai_num(CLng(Matches(0).submatches(5)))
                If sai Like "" Then
                    old_to_modern_no = MitumoriNo & convert_teiki_num(teiki)
                    Exit Function
                End If
                ans = ans & sai
            End If
        End With
    old_to_modern_no = ans & convert_teiki_num(teiki)
    Exit Function
    End If
    old_to_modern_no = MitumoriNo
End Function
Private Function convert_sai_num(sai_num As Long) As String
'旧再見積ニモニックを現行ニモニックにして返す
Dim table(9) As String
Dim i As Long
table(1) = "a"
table(2) = "b"
table(3) = "c"
table(4) = "d"
table(5) = "e"
table(6) = "f"
table(7) = "g"
table(8) = "h"
table(9) = "i"
    For i = 1 To 9
        If sai_num = i Then
            convert_sai_num = table(i)
            Exit Function
        End If
    Next
    convert_sai_num = ""
End Function
Private Function convert_teiki_num(teiki_num As String) As String
'旧見積定期番号を現行にして返す
    If Not teiki_num Like "" Then
        convert_teiki_num = Left(teiki_num, 3) & Right(teiki_num, 2)
        Exit Function
    End If
    convert_teiki_num = ""
End Function
Private Function getMainNo() As String
    getMainNo = m_Nendo & m_Eigyousyo_code & "-" & m_Serial_num
End Function
Private Function getTeikiNo() As String
    getTeikiNo = m_Teiki_year & m_Teiki_month
End Function
Private Sub ClearData()
    m_Nendo = ""
    m_Eigyousyo_code = ""
    m_eigyousyo_num = ""
    m_Serial_num = ""
    m_Sai_num = ""
    m_Teiki_year = ""
    m_Teiki_month = ""
    m_source = ""
    Message = ""
End Sub

Public Property Get MitumoriType() As String
'インスタンスの内部状態を返す
    If Not m_Teiki_month Like "" Then
        MitumoriType = "teiki"
        Exit Function
    End If
    If Not m_Sai_num Like "" Then
        MitumoriType = "sai"
        Exit Function
    End If
    If Not m_Serial_num = "" Then
        MitumoriType = "nomal"
        Exit Function
    End If
    MitumoriType = "empty"
End Property
Public Property Get source() As String
'pushで与えられた文字列を返す
    source = m_source
End Property
Public Property Get nendo() As String
'見積年を返す
    If m_Nendo = "" Then Message = "no year data"
    nendo = m_Nendo
End Property
Public Property Get MainNo() As String
'見積標準番号を返す
    Select Case MitumoriType
    Case "empty"
    Case Else
        MainNo = getMainNo
    End Select
End Property
Public Property Get saiNo() As String
'再見積ニモニックを返す
    saiNo = m_Sai_num
End Property
Public Function to_teiki(Optional m As String) As Boolean
'見積Noを定期No化する
Dim y As String
    y = CStr(year(Now) - 1988)
    Select Case m
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
    Case Else
        m = CStr(Month(Now))
    End Select
    If Len(m) = 1 Then m = "0" & m
    to_teiki = Push(MainNo & saiNo & "-" & y & m)
End Function
Public Function Push(MitumoriNo As String) As Boolean
'正規の見積番号か判定し、問題ない場合は格納し結果を返す
Dim RegEx As Variant
Dim pat As String
Dim Matches As Variant
Dim Mno As String
    ClearData
    m_source = MitumoriNo
    Mno = old_to_modern_no(MitumoriNo)
    pat = "^(\d{2})(F|K|KT|H|HD|M|I|AG|U|SK|A|KK|KH|KM|KI|KAG|KU|)-(\d{4})([a-z]|)(-\d{4}|)$"
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .pattern = pat
        .ignorecase = False
        Set Matches = .Execute(Mno)
        If Matches.Count = 0 Then
            Message = "not match"
            Push = False
            Exit Function
        End If
        m_Nendo = Matches(0).submatches(0)
        m_Eigyousyo_code = Matches(0).submatches(1)
        m_Serial_num = Matches(0).submatches(2)
        m_Sai_num = Matches(0).submatches(3)
        m_Teiki_month = Right(Matches(0).submatches(4), 2)
        m_Teiki_year = Mid(Matches(0).submatches(4), 2, 2)
    End With
    Push = True
End Function
Public Function Publish() As String
    Select Case MitumoriType
    Case "nomal"
        Publish = getMainNo
    Case "sai"
        Publish = getMainNo & m_Sai_num
    Case "teiki"
        Publish = getMainNo & m_Sai_num & "-" & getTeikiNo
    Case Else
        Publish = ""
    End Select
End Function
Public Function NextNumber() As Boolean
'nomalの場合は最初の再見積No
'saiの場合は次の再見積No
'teikiの場合は作成月の定期見積Noを返す
Dim mType As String
    Select Case MitumoriType
    Case "nomal"
        Push (Publish & "a")
    Case "sai"
        Push (MainNo & Chr(Asc(saiNo) + 1))
    Case "teiki"
        to_teiki
    Case Else
        ClearData
        Exit Function
    End Select
    NextNumber = True
End Function
