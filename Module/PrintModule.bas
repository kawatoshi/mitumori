Attribute VB_Name = "PrintModule"
Option Explicit
Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long) As Long
     
     
Public Function pb_fncGetPrinter(ByRef arg_varPrinter() As Variant, _
                        ByRef arg_varPort() As Variant, ByRef arg_strErr As String) As Long
  
    Const STR_APPNAME As String = "Devices"               '�ړI�̃L�[���������Ă���Z�N�V�����̖��O�ilpAppName�j
    Const STR_DEFAULT As String = "������܂���ł���"    '�K��̕�����(lpDefault)
    Const LNG_SIZE As Long = 1024                         '�����i�[����o�b�t�@�̃T�C�Y(nSize)
    Const STR_KEYNAME As String = vbNullString            '�Z�N�V�������̑S�ẴL�[���擾(NULL���w��)
  
    Dim lngRet As Long      'GetProfileString�֐��̖߂�l
    Dim strReturnedString As String * 1024
    Dim strTmp As String
    Dim lngNull As Long
    Dim i As Long
    Dim lngStart As Long
    Dim strErr As String
  
    On Error GoTo ErrHandler
  
    '-Start---------------------------------------------------------
    '�v�����^�ꗗ���擾
  
    '�w�肵���Z�N�V�������������A�Z�N�V�����̑S�L�[���擾�A�Y���f�[�^�̃o�C�g����Ԃ�
    '�o�b�t�@�istrReturnedString�j�Ɋi�[���ꂽ���������Ԃ�
    lngRet = GetProfileString(STR_APPNAME, STR_KEYNAME, STR_DEFAULT, strReturnedString, LNG_SIZE)
  
    '�Ō��NULL������
    strTmp = Left(strReturnedString, InStr(1, strReturnedString, Chr(0) & Chr(0)) - 1)
  
    '�߂�l�`�F�b�N
    If strTmp = STR_DEFAULT Then
        strErr = "�v�����^�[�����擾�ł��܂���ł���"
        GoTo ErrHandler
    End If
  
    lngNull = 0
    i = 0
    lngStart = 0
  
    Do
        i = i + 1
        lngNull = InStr(lngNull + 1, strTmp, Chr(0))
        If lngNull = 0 Then lngNull = Len(strTmp)
        ReDim Preserve arg_varPrinter(1 To i)
        arg_varPrinter(i) = Mid(strTmp, lngStart + 1, lngNull - lngStart)
        If Right(arg_varPrinter(i), 1) = Chr(0) Then    '������NULL���폜
            arg_varPrinter(i) = Left(arg_varPrinter(i), Len(arg_varPrinter(i)) - 1)
        End If
        lngStart = lngNull
    Loop Until lngNull = Len(strTmp)
    '-End-----------------------------------------------------------
  
    pb_fncGetPrinter = i
    ReDim arg_varPort(1 To i)
  
    '-Start---------------------------------------------------------
    '�|�[�g�ꗗ���擾
    For i = 1 To pb_fncGetPrinter
      
        lngRet = GetProfileString(STR_APPNAME, arg_varPrinter(i), STR_DEFAULT, strReturnedString, LNG_SIZE)
    
        '�Ō��NULL������
        strTmp = Left(strReturnedString, InStr(1, strReturnedString, Chr(0)) - 1)
        strTmp = Mid(strTmp, InStr(1, strTmp, ",") + 1)
    
        '�߂�l�`�F�b�N
        If strTmp = STR_DEFAULT Then
            strErr = "�|�[�g�����擾�ł��܂���ł���"
            GoTo ErrHandler
        Else
            arg_varPort(i) = strTmp
        End If
  
    Next i
    '-End-----------------------------------------------------------
  
    Exit Function

ErrHandler:
    arg_strErr = strErr & vbCrLf & _
                "�t�H�[������ďI�������Ă��������B" & _
                vbCrLf & vbCrLf & Err.Number & " : " & Err.Description
    pb_fncGetPrinter = 0

End Function

Function getPrinterName()
    getPrinterName = Application.ActivePrinter
End Function
Function findPrinter(strPrinter) As String
'�����̃v�����^�[�����݂���ꍇ�̓v�����^�[����Ԃ�
'���݂��Ȃ��ꍇ�A�f�t�H���g�v�����^�[��Ԃ�
Dim varPrinter() As Variant
Dim varPort() As Variant
Dim strErrMsg As String
Dim varName As Variant
Dim i As Long
    Call pb_fncGetPrinter(varPrinter(), varPort(), strErrMsg)
    i = 1
    For Each varName In varPrinter
        If varName Like strPrinter Then
            findPrinter = varName & " on " & varPort(i)
            Exit Function
        End If
        i = i + 1
    Next
    findPrinter = getPrinterName
End Function
