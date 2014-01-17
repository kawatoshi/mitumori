Attribute VB_Name = "TypeModule"
Type HyoudaiData
    strSerial As String         '入力番号
    strMitumoriNo As String     '見積No
    strCustomer As String       '宛名
    dteMitumoriDay As Date      '見積日
    strFormat As String         '書式タイプ
    strBumon As String          '担当部門
    strSite As String           '所在地
    strLocation As String       '位置
    strKiHyouki As String       '貴表記
    strName As String           '場所・テナント名
    strContents As String       '工事名称
    strDeliveryPlace As String  '納入場所
    strSiharai As String        '支払い方法
    strYuukoukikann As String   '有効期限
    dblProceeds As Double       '金額（税込）
    dblSum As Double            '金額(税別）
    dblCost As Double           '原価（税別）
    strNotes As String          'メモ
    strMaker As String          '作成者
    dteSeikyuuDay As Date       '請求日
    strSeikyuuType As String    '請求方法
    strMsinsei As String        '申請日時
    dblTaxRate As Double            '消費税
    strPublishRequestType As String     '発行申請
    strMitumoriPresentDay As String     '見積提出日
    strAccountsDate As String       '決済日
    strCheckOfAccounts As String     '受注確認者
    strCheckOfFinishing As String   '完了確認日
    strWorkReport As String         '作業報告書
    strUriageTuki As String         '売上月
End Type

Public Type MitumoriHyoudaiData
'発行文書の表題用
    strMno As String            '見積書用No
    strDate As String           '日付
    strCustomer As String       '宛名
    strFujibil As String
    strEigyousyo As String
    strWork1 As String
    strWork2 As String
    strKoujiKikan As String
    strYuukoukikann As String   '見積有効期間
    strSiharai As String        '御支払方法
    strFjibilFace As Long       '不二ビル表記
    strEigyousyoFace As Long    '営業所表記
    dblTaxRate As Double            '消費税
End Type

Public Type SyousaiData
    strMitumoriNo As String     '見積NO
    strHeader As String         'No
    strContents As String       '品名
    strSpec As String           '仕様
    strNumber As String         '数量
    strUnit As String           '単位
    strPrice As String          '単価
    strSum As String            '金額
    strNote As String           '備考
End Type

Public Type MitumoriSyousaidata
'発行文書の詳細用
    strStar As String
    strWork1 As String
    strWork2 As String
End Type

Public Type UtiwakeData
    strMitumoriNo As String     '見積NO
    strHeader As String         'No
    strContents As String       '品名
    strSpec As String           '仕様
    strNumber As String         '数量
    strUnit As String           '単位
    strPrice As String          '単価
    strSum As String            '金額
    strNote As String           '備考
    strPage As String           'ページNo
End Type

Public Type GyousyaData
    strMitumoriNo As String     '見積No
    strGyousya As String        '見積業者
    strCost As String           '支払金額（税抜）
    strCostWithTax As String    '支払金額(税込)
    strBillMonth As String      '支払い月
End Type

Public Type ReadMailData
    strTo As String             '宛先
    strFrom As String           '送信者
    strDate As String           '処理日
    strSubject As String
    strBody As String           '本文
    strFile As String           '添付ファイル名（フルパス)
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
    strHidukeType As String     '日付表記方法
    strEigyousyo As String      '営業所
    strSite As String           '所在地
    strKiHyouki As String       '貴表記
    strSiharai As String        '御支払方法
End Type

Public Type MakeParameter
'putMakeに渡すデータ
    strKey As String
    lngShiftColumn As Long
    lngColumns As Long
    lngRows As Long
    rngPaste As Range
End Type

Public Type CommissionData
'捺印依頼書のデータ
    strSinseiType As String     '申請方法
    strMno As String            '見積・請求番号
    strCustomer As String       '宛名
    strWork As String           '工事名称
    dblProceeds As Double       '請求金額（税込み）
    strPublishRequestType As String   '支払い方法
    strDates As String          '作業日
    strGyousya As String        '業者名
    strCosts As String          '支払金額
End Type

Public Type CommissionRange
'捺印依頼書のRange
    Mno As Range
    Customor As Range
    Work As Range
    Proceeds As Range
    Dates As Range
    Gyousya As Range
    Costs As Range
End Type

