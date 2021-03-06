Attribute VB_Name = "TypeModule"
Type HyoudaiData
    strSerial As String         'üÍÔ
    strMitumoriNo As String     '©ÏNo
    strCustomer As String       '¶¼
    dteMitumoriDay As Date      '©Ïú
    strFormat As String         '®^Cv
    strBumon As String          'Så
    strSite As String           'Ýn
    strLocation As String       'Êu
    strKiHyouki As String       'M\L
    strName As String           'êEeig¼
    strContents As String       'H¼Ì
    strDeliveryPlace As String  '[üê
    strSiharai As String        'x¥¢û@
    strYuukoukikann As String   'LøúÀ
    dblProceeds As Double       'àziÅj
    dblSum As Double            'àz(ÅÊj
    dblCost As Double           '´¿iÅÊj
    strNotes As String          '
    strMaker As String          'ì¬Ò
    dteSeikyuuDay As Date       '¿ú
    strSeikyuuType As String    '¿û@
    strMsinsei As String        '\¿ú
    dblTaxRate As Double            'ÁïÅ
    strPublishRequestType As String     '­s\¿
    strMitumoriPresentDay As String     '©Ïñoú
    strAccountsDate As String       'Ïú
    strCheckOfAccounts As String     'ómFÒ
    strCheckOfFinishing As String   '®¹mFú
    strWorkReport As String         'ìÆñ
    strUriageTuki As String         'ã
End Type

Public Type MitumoriHyoudaiData
'­s¶Ì\èp
    strMno As String            '©ÏpNo
    strDate As String           'út
    strCustomer As String       '¶¼
    strFujibil As String
    strEigyousyo As String
    strWork1 As String
    strWork2 As String
    strKoujiKikan As String
    strYuukoukikann As String   '©ÏLøúÔ
    strSiharai As String        'äx¥û@
    strFjibilFace As Long       'sñr\L
    strEigyousyoFace As Long    'cÆ\L
    dblTaxRate As Double            'ÁïÅ
End Type

Public Type SyousaiData
    strMitumoriNo As String     '©ÏNO
    strHeader As String         'No
    strContents As String       'i¼
    strSpec As String           'dl
    strNumber As String         'Ê
    strUnit As String           'PÊ
    strPrice As String          'P¿
    strSum As String            'àz
    strNote As String           'õl
End Type

Public Type MitumoriSyousaidata
'­s¶ÌÚ×p
    strStar As String
    strWork1 As String
    strWork2 As String
End Type

Public Type UtiwakeData
    strMitumoriNo As String     '©ÏNO
    strHeader As String         'No
    strContents As String       'i¼
    strSpec As String           'dl
    strNumber As String         'Ê
    strUnit As String           'PÊ
    strPrice As String          'P¿
    strSum As String            'àz
    strNote As String           'õl
    strPage As String           'y[WNo
End Type

Public Type GyousyaData
    strMitumoriNo As String     '©ÏNo
    strGyousya As String        '©ÏÆÒ
    strCost As String           'x¥àziÅ²j
    strCostWithTax As String    'x¥àz(Å)
    strBillMonth As String      'x¥¢
End Type

Public Type ReadMailData
    strTo As String             '¶æ
    strFrom As String           'MÒ
    strDate As String           'ú
    strSubject As String
    strBody As String           '{¶
    strFile As String           'Ytt@C¼itpX)
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
    strHidukeType As String     'út\Lû@
    strEigyousyo As String      'cÆ
    strSite As String           'Ýn
    strKiHyouki As String       'M\L
    strSiharai As String        'äx¥û@
End Type

Public Type MakeParameter
'putMakeÉn·f[^
    strKey As String
    lngShiftColumn As Long
    lngColumns As Long
    lngRows As Long
    rngPaste As Range
End Type

Public Type CommissionData
'æóËÌf[^
    strSinseiType As String     '\¿û@
    strMno As String            '©ÏE¿Ô
    strCustomer As String       '¶¼
    strWork As String           'H¼Ì
    dblProceeds As Double       '¿àziÅÝj
    strPublishRequestType As String   'x¥¢û@
    strDates As String          'ìÆú
    strGyousya As String        'ÆÒ¼
    strCosts As String          'x¥àz
End Type

Public Type CommissionRange
'æóËÌRange
    Mno As Range
    Customor As Range
    Work As Range
    Proceeds As Range
    Dates As Range
    Gyousya As Range
    Costs As Range
End Type

