Attribute VB_Name = "InitModule"
Option Explicit
Private Function initShapes(shtMy As Worksheet) As Boolean
'shtMy�ɗ^����ꂽ�V�[�g��"Button"Shapes��S�ď�������
    Dim i As Long, n As Long
    Dim lngShapes As Long
    Dim strShapeName As String
    
    initShapes = False
    lngShapes = shtMy.Shapes.Count
    If lngShapes <= 1 Then Exit Function
    With shtMy
    n = 1
        For i = lngShapes To 1 Step -1
            strShapeName = Left$(.Shapes(n).Name, 5)
            Select Case strShapeName
                Case "Drop "
                    n = n + 1
                Case Else
                    .Shapes(n).Delete
            End Select
        Next
    End With
    initShapes = True
End Function
Private Function ClearSheet(shtMy As Worksheet)
'shtMy�V�[�g�̃f�[�^���N���A����
    Call initShapes(shtMy)
    shtMy.Cells.Delete Shift:=xlUp
End Function
Private Function MakeButton(shtName As Worksheet, botton As BottonCoordinates) As String
'varCoordinates�ɍ��W�AstrOnAction�Ɏ��s����A�N�V����
'strChaText�Ƀ{�^���ɕ\������e�L�X�g��^��
'shtName�V�[�g�ɋL�q����
Dim strMacroBook As String
Dim strButtonName As String

    strMacroBook = ThisWorkbook.Name
    With shtName
        strButtonName = .Buttons.Add(botton.dblX, botton.dblY, botton.dblW, botton.dblH).Name
        .Buttons(strButtonName).OnAction = strMacroBook & "!" & botton.strOnAction
        .Buttons(strButtonName).Characters.Text = botton.strChaText
    End With
    
    MakeButton = "OK"
End Function
Private Function MakeHyoudaiButtons()
'�\��V�[�g�Ƀ{�^���ƃ}�N����z�u����
    Dim shtinput As Worksheet
    Dim botton() As BottonCoordinates
    Dim i As Long

    '�z�u�{�^���̃f�[�^�ݒ�
    'X���W�AY���W�A�{�^�����A�{�^�������A�N���}�N���A�\���e�L�X�g
    ReDim botton(6)
    botton(0).dblX = 4.25: botton(0).dblY = 4.25: _
        botton(0).dblW = 70.25: botton(0).dblH = 20: _
        botton(0).strOnAction = "�\�����ޕύX": botton(0).strChaText = "�\�����ޕύX"
        
    botton(1).dblX = 4.25: botton(1).dblY = 28.5: _
        botton(1).dblW = 70.25: botton(1).dblH = 20: _
        botton(1).strOnAction = "���˗��쐬": botton(1).strChaText = "���˗��쐬"
        
    botton(2).dblX = 78.75: botton(2).dblY = 4.25: _
        botton(2).dblW = 70.25: botton(2).dblH = 20: _
        botton(2).strOnAction = "�{���\��": botton(2).strChaText = "�{���\��"
        
    botton(3).dblX = 78.75: botton(3).dblY = 28.5: _
        botton(3).dblW = 70.25: botton(3).dblH = 20: _
        botton(3).strOnAction = "�ē���": botton(3).strChaText = "�ē���"
        
    botton(4).dblX = 153.25: botton(4).dblY = 4.25: _
        botton(4).dblW = 70.25: botton(4).dblH = 20: _
        botton(4).strOnAction = "�\���m�F": botton(4).strChaText = "�\�����ފm�F"
    
    botton(5).dblX = 153.25: botton(5).dblY = 28.5: _
        botton(5).dblW = 70.25: botton(5).dblH = 20: _
        botton(5).strOnAction = "�V�K����": botton(5).strChaText = "�V�K�쐬"
        
    botton(6).dblX = 227.75: botton(6).dblY = 4.25: _
        botton(6).dblW = 70.25: botton(6).dblH = 20: _
        botton(6).strOnAction = "����쐬": botton(6).strChaText = "����쐬"
        
    '�{�^���̔z�u
    Set shtinput = Sheets("�\��")
    For i = 0 To UBound(botton())
        Call MakeButton(shtinput, botton(i))
    Next
    Set shtinput = Nothing
End Function
Private Function MakeInputButtons()
    Dim shtinput As Worksheet
    Dim botton() As BottonCoordinates
    Dim i As Long

    '�z�u�{�^���̃f�[�^�ݒ�
    'X���W�AY���W�A�{�^�����A�{�^�������A�N���}�N���A�\���e�L�X�g
    ReDim botton(2)
    botton(0).dblX = 746.75: botton(0).dblY = 4.25: _
        botton(0).dblW = 62.25: botton(0).dblH = 20: _
        botton(0).strOnAction = "�V�K����": botton(0).strChaText = "�V�K����"
        
    botton(1).dblX = 746.75: botton(1).dblY = 28.5: _
        botton(1).dblW = 62.25: botton(1).dblH = 20: _
        botton(1).strOnAction = "��荞��": botton(1).strChaText = "��荞��"
        
    botton(2).dblX = 746.75: botton(2).dblY = 52.75: _
        botton(2).dblW = 62.25: botton(2).dblH = 20: _
        botton(2).strOnAction = "�]�L": botton(2).strChaText = "�]�L"
    '�{�^���̔z�u
    Set shtinput = ActiveWorkbook.Sheets("����")
    For i = 0 To UBound(botton())
        Call MakeButton(shtinput, botton(i))
    Next
    Set shtinput = Nothing
End Function

Function initSyosiki(shtSource As Worksheet, shtDestination As Worksheet) As Boolean
'�]�L�V�[�g���������A�������V�K�ɃR�s�[����
Dim i As Long
Dim lngShapesCounter As Long

On Error GoTo Error
'�]�L�V�[�g�̏���
    Call ClearSheet(shtDestination)
'��������̃R�s�[
    shtSource.Cells.Copy Destination:=shtDestination.Cells
    shtDestination.Range("a1").Copy Destination:=shtDestination.Range("a1")
    shtDestination.ResetAllPageBreaks
    initSyosiki = True
    Exit Function
Error:
initSyosiki = False
End Function
Function initUtiwakeSyosiki(shtSrc As Worksheet, shtDst As Worksheet, page As Long) As Boolean
'���ϓ���̏����z�u���s��
Dim UtiwakeRows As Long
Dim strRows As String
Dim DstRow As Long
Dim i As Long
    UtiwakeRows = getUtiwakePageRows + 4
    strRows = "1:" & CStr(UtiwakeRows)
    For i = 0 To page - 1
    DstRow = 38 + (i * UtiwakeRows)
        shtSrc.Rows(strRows).Copy Destination:=shtDst.Rows(DstRow)
        shtDst.HPageBreaks.Add before:=shtDst.Cells(DstRow, 1)
    Next
End Function
Function initInputSheet()
'�V�K���͂̏�����
Dim shtMy As Worksheet
Dim shtCopy As Worksheet
    Set shtMy = Sheets("�\��")
    Call initShapes(shtMy)
    MakeHyoudaiButtons
    Set shtCopy = ActiveWorkbook.Sheets("���͌���")
    Set shtMy = ActiveWorkbook.Sheets("����")
    Call initSyosiki(shtCopy, shtMy)
    MakeInputButtons
End Function
