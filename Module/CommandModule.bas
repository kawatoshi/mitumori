Attribute VB_Name = "CommandModule"
Option Explicit
'�V�[�g�ɔz�u����Ă���{�^���͂��ׂĂ��̃��W���[����
'�L�ڂ���Ă���sub�v���V�[�W���[����N�������

Public Sub version()
    Call MsgBox("version: 2.7.4" & Chr(13) & "code by kawakita  2011.03.17", vbOKOnly, "version")
End Sub

Sub �V�K����()
    Call initInputSheet
    ActiveWorkbook.Sheets("����").Select
    Cells(2, 2).Select
End Sub

Sub �ē���()
'�\��V�[�g�̓��e���Č��ς��肷��
Dim ans As String
    ans = rekeyMitumori(getMitumoriNo)
    If Not ans Like "" Then
        MsgBox (ans)
    Else
        Sheets("����").Activate
        Range("f2").Select
    End If
End Sub

Sub �]�L()
Dim bokMy As Workbook
Set bokMy = ActiveWorkbook
    If SheetInput(bokMy) = True Then
        Call initInputSheet
        bokMy.Sheets("�\��").Activate
        getHyoudaiEndCell(bokMy).Select
    End If
End Sub

Sub �\���m�F()
'�\�������쐬����
Dim Mno As String
    Mno = getMitumoriNo
    If SinseiKakunin(Mno) = False Then
        Call MsgBox("�\���f�[�^�ɖ�肪���邽�߁A�쐬�ł��܂���ł���")
        Exit Sub
    End If
    If SinseiType_is_Seikyuu(Mno) = True Then
        Sheets("������").Activate
        Exit Sub
    End If
    If SinseiType_is_Mitumori(Mno) = True Then
        Sheets("���Ϗ�").Activate
        Exit Sub
    End If
    Call MsgBox("�\�����쐬�ł��܂���ł���", vbOKOnly)
End Sub

Sub ���˗��쐬()
'�ʏ�o��
Dim Mnos() As String
Dim strOutPut As String
    Mnos = getMitumoriNoOnRangeAreas
    strOutPut = Range("printer_name").Value
    Select Case strOutPut
    Case ""
        Call MakeCommissions(Mnos, "", "", False)
    Case Else
        Call MakeCommissions(Mnos, strOutPut, "", False)
    End Select
End Sub
Sub �˗�����()
'�v�����^�w����
Dim Mnos() As String
Dim strOutPut As String
    Mnos = getMitumoriNoOnRangeAreas
    strOutPut = Range("toku_printer").Value
    Call MakeCommissions(Mnos, strOutPut, "", False)
End Sub
Sub �˗�Book()
'Book�o��
Dim Mnos() As String
Dim strBook As String
    Mnos = getMitumoriNoOnRangeAreas
    strBook = Range("commission_file")
    Call MakeCommissions(Mnos, "", strBook, True)
End Sub
Sub �\�����ޕύX()
        If getMitumoriNo Like "" Then
            MsgBox ("�I���s�ł͐\����ύX�ł��܂���")
        Else
            frmSinsei.Show
        End If
End Sub
Public Sub �{���\��()
'�\��V�[�g�̔��s�\���^�C�v����K�؂Ȑ\���������ōs��
Dim Mnos() As String
Dim i As Long
    Mnos = getMitumoriNoOnRangeAreas
    For i = 0 To UBound(Mnos)
        Call MakeSinseiToHQ(Mnos(i))
    Next
End Sub
Sub ��荞��()
'���̓V�[�g�ւ̎�荞��
Dim ans As String
    ans = TorikomiSheetInput
    If Not ans Like "" Then
        Call MsgBox(ans)
    End If
End Sub
Sub ����쐬()
'������ς�]�L����
Dim lngMonth As Long
Dim strMonth As String
Dim bokMy As Workbook
Dim ans As String
    lngMonth = Month(Now())
    strMonth = InputBox("�쐬������͂��Ă�������", "����쐬", CStr(lngMonth))
    Set bokMy = ActiveWorkbook
    ans = makeTeiki(CLng(strMonth), bokMy)
    Call MsgBox(ans, vbOKOnly, "����쐬����")
End Sub
