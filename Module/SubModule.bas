Attribute VB_Name = "SubModule"
Option Explicit
Private Function chkFolder(strPath As String) As String
'�����œn���ꂽ�t�H���_�܂��̓t�@�C�������݂��邩��
'�m�F���A���ʂ�Ԃ�:folder:file:NG
Dim lngERROR As Long
'�t�H���_�̃`�F�b�N
    On Error GoTo ER
    If Len(Dir(strPath, vbDirectory)) > 0 Then
        If (GetAttr(strPath) And vbDirectory) = vbDirectory Then
            chkFolder = "folder"
        Else
            chkFolder = "file"
        End If
        On Error GoTo 0
    Else
        chkFolder = "NG"
    End If
    Exit Function
ER:
    chkFolder = "ERROR"
End Function

Function BookSave(wbkMy As Workbook, strPath As String, strFileName As String) As String
Dim lngErr As Long
    On Error Resume Next
    Application.DisplayAlerts = False
    wbkMy.SaveAs (getBaseName(strPath & strFileName) & ".xls")
    Application.DisplayAlerts = True
    lngErr = Err.Number
    On Error GoTo 0
    If lngErr <> 0 Then
        BookSave = "�ۑ��ł��܂���ł���"
    Else
        BookSave = "ok"
    End If
End Function
Function getBaseName(strName As String) As String
Dim i As Long
    i = InStrRev(strName, ".")
    If i > 0 Then
        getBaseName = Left(strName, i - 1)
    Else
        getBaseName = strName
    End If
End Function
Function OpenBook(strReqFolder As String, strReqFile As String, bokDst As Workbook) As String
'�w��t�H���_���w��u�b�N�����݂��Ă��邩���m�F���ʃt�H���_�����t�@�C�������݂��Ȃ����Ƃ�
'�m�F���ĊJ��
    '�l�b�g���[�N�f�B���N�g���ƃt�@�C���̊m�F
    If chkFolder(strReqFolder) <> "folder" Then
        OpenBook = "�\������f�B���N�g���[�����݂��܂���" & Chr(13) & _
                   "folder: " & strReqFolder
        Exit Function
    End If
    If chkFolder(strReqFolder & "\" & strReqFile) <> "file" Then
        OpenBook = "�\������t�@�C�������݂��܂���" & Chr(13) & _
                   "file: " & strReqFile
        Exit Function
    End If
    '�����u�b�N���J���Ă��邩�̊m�F
    Set bokDst = Book_is_opend(strReqFile)
    If bokDst Is Nothing Then
         Set bokDst = Workbooks.Open(Filename:=strReqFolder & "\" & strReqFile, UpdateLinks:=0, ReadOnly:=False)
    Else
         '�ʃt�H���_�̓����t�@�C�����J���Ă��Ȃ����̊m�F
        If Not LCase(bokDst.Path) Like LCase(strReqFolder) Then
            OpenBook = "�����̃t�@�C�������łɊJ���Ă��܂��B " & Chr(13) & _
                       "file: " & strReqFile & "����Ă�������"
            Set bokDst = Nothing
            Exit Function
        End If
        '�ǂݎ���p�ŊJ���Ă��Ȃ����̊m�F
        If bokDst.ReadOnly = True Then
            OpenBook = "�ǂݎ���p�ŊJ����Ă���̂ŏ������ݏo���܂���ERROR"
            Set bokDst = Nothing
        End If
    End If
End Function
Function MakeSinsei(strFolder As String, _
                    strFile As String, _
                    Hdata As HyoudaiData, _
                    bokMy As Workbook, _
                    strSinseiType As String, _
                    lngZoom As Long) As String
'�\�������J�������NO�Ɠ����̃V�[�g���쐬��]�L����
'strSinseiType: "m"���ρ@"s"����
Dim bokDst As Workbook
Dim ans As String
Dim shtDst As Worksheet
Dim shtSrc As Worksheet
'�w�胏�[�N�u�b�N���J��
    ans = OpenBook(strFolder, strFile, bokDst)
    If bokDst Is Nothing Then
        MakeSinsei = ans
        Exit Function
    End If
'�����̃V�[�g�����݂����ꍇ�̊m�F
    If has_same_sheet(Hdata.strMitumoriNo, bokDst) = True Then
        If Not MsgBox(Hdata.strMitumoriNo & "�͂��łɐ\������Ă��܂��B���������܂���?", vbYesNo) = vbYes Then
            MakeSinsei = "�����������~"
            Exit Function
        End If
        Set shtDst = bokDst.Sheets(Hdata.strMitumoriNo)
    End If
'�����^�C�v�ʂ̓]�L
    If shtDst Is Nothing Then
        Set shtDst = bokDst.Sheets.Add(after:=bokDst.Worksheets(bokDst.Sheets(bokDst.Sheets.Count).Name))
        shtDst.Name = Hdata.strMitumoriNo
        Call printerSetUp(shtDst)
    End If
    Select Case strSinseiType
    Case "m"
        Set shtSrc = bokMy.Sheets("���ό���")
        If publishMitumori(Hdata.strMitumoriNo, shtSrc, shtDst) = False Then
            ans = "�\���]�L�ɃG���[���������܂����B�Ǘ��҂ɑ��k���Ă�������" & Chr(13) & _
                  "����No : " & Hdata.strMitumoriNo & "   " & shtSrc.Name
        End If
    Case "s"
        Set shtSrc = bokMy.Sheets("��������")
        If publishSeikyuu(Hdata.strMitumoriNo, shtSrc, shtDst) = False Then
            ans = "�\���]�L�ɃG���[���������܂����B�Ǘ��҂ɑ��k���Ă�������" & Chr(13) & _
                  "����No : " & Hdata.strMitumoriNo & "   " & shtSrc.Name
        End If
        If ans Like "" Then Call SheetZoomSetUp(shtDst, lngZoom)
   End Select
    MakeSinsei = ans
End Function
Function printerSetUp(shtPrint As Worksheet) As Boolean
Dim appVer As Double
    appVer = Application.version
    If appVer >= 14 Then
        Application.PrintCommunication = False
    End If
    shtPrint.PageSetup.PrintArea = ""
    With shtPrint.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        If appVer >= 14 Then
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End If
    End With
    If appVer >= 14 Then
        Application.PrintCommunication = True
    End If
    printerSetUp = True
End Function
Function SheetZoomSetUp(shtZoom As Worksheet, ZoomRate As Long) As Boolean
        shtZoom.Parent.Windows(1).Zoom = ZoomRate
End Function
