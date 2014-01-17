Attribute VB_Name = "TypeModule"
Type HyoudaiData
    strSerial As String         '���͔ԍ�
    strMitumoriNo As String     '����No
    strCustomer As String       '����
    dteMitumoriDay As Date      '���ϓ�
    strFormat As String         '�����^�C�v
    strBumon As String          '�S������
    strSite As String           '���ݒn
    strLocation As String       '�ʒu
    strKiHyouki As String       '�M�\�L
    strName As String           '�ꏊ�E�e�i���g��
    strContents As String       '�H������
    strDeliveryPlace As String  '�[���ꏊ
    strSiharai As String        '�x�������@
    strYuukoukikann As String   '�L������
    dblProceeds As Double       '���z�i�ō��j
    dblSum As Double            '���z(�ŕʁj
    dblCost As Double           '�����i�ŕʁj
    strNotes As String          '����
    strMaker As String          '�쐬��
    dteSeikyuuDay As Date       '������
    strSeikyuuType As String    '�������@
    strMsinsei As String        '�\������
    dblTaxRate As Double            '�����
    strPublishRequestType As String     '���s�\��
    strMitumoriPresentDay As String     '���ϒ�o��
    strAccountsDate As String       '���ϓ�
    strCheckOfAccounts As String     '�󒍊m�F��
    strCheckOfFinishing As String   '�����m�F��
    strWorkReport As String         '��ƕ񍐏�
    strUriageTuki As String         '���㌎
End Type

Public Type MitumoriHyoudaiData
'���s�����̕\��p
    strMno As String            '���Ϗ��pNo
    strDate As String           '���t
    strCustomer As String       '����
    strFujibil As String
    strEigyousyo As String
    strWork1 As String
    strWork2 As String
    strKoujiKikan As String
    strYuukoukikann As String   '���ϗL������
    strSiharai As String        '��x�����@
    strFjibilFace As Long       '�s��r���\�L
    strEigyousyoFace As Long    '�c�Ə��\�L
    dblTaxRate As Double            '�����
End Type

Public Type SyousaiData
    strMitumoriNo As String     '����NO
    strHeader As String         'No
    strContents As String       '�i��
    strSpec As String           '�d�l
    strNumber As String         '����
    strUnit As String           '�P��
    strPrice As String          '�P��
    strSum As String            '���z
    strNote As String           '���l
End Type

Public Type MitumoriSyousaidata
'���s�����̏ڍחp
    strStar As String
    strWork1 As String
    strWork2 As String
End Type

Public Type UtiwakeData
    strMitumoriNo As String     '����NO
    strHeader As String         'No
    strContents As String       '�i��
    strSpec As String           '�d�l
    strNumber As String         '����
    strUnit As String           '�P��
    strPrice As String          '�P��
    strSum As String            '���z
    strNote As String           '���l
    strPage As String           '�y�[�WNo
End Type

Public Type GyousyaData
    strMitumoriNo As String     '����No
    strGyousya As String        '���ϋƎ�
    strCost As String           '�x�����z�i�Ŕ��j
    strCostWithTax As String    '�x�����z(�ō�)
    strBillMonth As String      '�x������
End Type

Public Type ReadMailData
    strTo As String             '����
    strFrom As String           '���M��
    strDate As String           '������
    strSubject As String
    strBody As String           '�{��
    strFile As String           '�Y�t�t�@�C�����i�t���p�X)
End Type

Public Type BottonCoordinates
    dblX As Double
    dblY As Double
    dblW As Double
    dblH As Double
    strOnAction As String
    strChaText As String
End Type

Public Type CustomaryHyoudaiData
    strHidukeType As String     '���t�\�L���@
    strEigyousyo As String      '�c�Ə�
    strSite As String           '���ݒn
    strKiHyouki As String       '�M�\�L
    strSiharai As String        '��x�����@
End Type

Public Type MakeParameter
'putMake�ɓn���f�[�^
    strKey As String
    lngShiftColumn As Long
    lngColumns As Long
    lngRows As Long
    rngPaste As Range
End Type

Public Type CommissionData
'���˗����̃f�[�^
    strSinseiType As String     '�\�����@
    strMno As String            '���ρE�����ԍ�
    strCustomer As String       '����
    strWork As String           '�H������
    dblProceeds As Double       '�������z�i�ō��݁j
    strPublishRequestType As String   '�x�������@
    strDates As String          '��Ɠ�
    strGyousya As String        '�ƎҖ�
    strCosts As String          '�x�����z
End Type

Public Type CommissionRange
'���˗�����Range
    Mno As Range
    Customor As Range
    Work As Range
    Proceeds As Range
    Dates As Range
    Gyousya As Range
    Costs As Range
End Type

