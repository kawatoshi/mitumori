Attribute VB_Name = "InitModule"
Option Explicit
Private Function initShapes(shtMy As Worksheet) As Boolean
'shtMyに与えられたシートの"Button"Shapesを全て消去する
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
'shtMyシートのデータをクリアする
    Call initShapes(shtMy)
    shtMy.Cells.Delete Shift:=xlUp
End Function
Private Function MakeButton(shtName As Worksheet, botton As BottonCoordinates) As String
'varCoordinatesに座標、strOnActionに実行するアクション
'strChaTextにボタンに表示するテキストを与え
'shtNameシートに記述する
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
'表題シートにボタンとマクロを配置する
    Dim shtinput As Worksheet
    Dim botton() As BottonCoordinates
    Dim i As Long

    '配置ボタンのデータ設定
    'X座標、Y座標、ボタン幅、ボタン高さ、起動マクロ、表示テキスト
    ReDim botton(6)
    botton(0).dblX = 4.25: botton(0).dblY = 4.25: _
        botton(0).dblW = 70.25: botton(0).dblH = 20: _
        botton(0).strOnAction = "申請書類変更": botton(0).strChaText = "申請書類変更"
        
    botton(1).dblX = 4.25: botton(1).dblY = 28.5: _
        botton(1).dblW = 70.25: botton(1).dblH = 20: _
        botton(1).strOnAction = "捺印依頼作成": botton(1).strChaText = "捺印依頼作成"
        
    botton(2).dblX = 78.75: botton(2).dblY = 4.25: _
        botton(2).dblW = 70.25: botton(2).dblH = 20: _
        botton(2).strOnAction = "本部申請": botton(2).strChaText = "本部申請"
        
    botton(3).dblX = 78.75: botton(3).dblY = 28.5: _
        botton(3).dblW = 70.25: botton(3).dblH = 20: _
        botton(3).strOnAction = "再入力": botton(3).strChaText = "再入力"
        
    botton(4).dblX = 153.25: botton(4).dblY = 4.25: _
        botton(4).dblW = 70.25: botton(4).dblH = 20: _
        botton(4).strOnAction = "申請確認": botton(4).strChaText = "申請書類確認"
    
    botton(5).dblX = 153.25: botton(5).dblY = 28.5: _
        botton(5).dblW = 70.25: botton(5).dblH = 20: _
        botton(5).strOnAction = "新規入力": botton(5).strChaText = "新規作成"
        
    botton(6).dblX = 227.75: botton(6).dblY = 4.25: _
        botton(6).dblW = 70.25: botton(6).dblH = 20: _
        botton(6).strOnAction = "定期作成": botton(6).strChaText = "定期作成"
        
    'ボタンの配置
    Set shtinput = Sheets("表題")
    For i = 0 To UBound(botton())
        Call MakeButton(shtinput, botton(i))
    Next
    Set shtinput = Nothing
End Function
Private Function MakeInputButtons()
    Dim shtinput As Worksheet
    Dim botton() As BottonCoordinates
    Dim i As Long

    '配置ボタンのデータ設定
    'X座標、Y座標、ボタン幅、ボタン高さ、起動マクロ、表示テキスト
    ReDim botton(2)
    botton(0).dblX = 746.75: botton(0).dblY = 4.25: _
        botton(0).dblW = 62.25: botton(0).dblH = 20: _
        botton(0).strOnAction = "新規入力": botton(0).strChaText = "新規入力"
        
    botton(1).dblX = 746.75: botton(1).dblY = 28.5: _
        botton(1).dblW = 62.25: botton(1).dblH = 20: _
        botton(1).strOnAction = "取り込み": botton(1).strChaText = "取り込み"
        
    botton(2).dblX = 746.75: botton(2).dblY = 52.75: _
        botton(2).dblW = 62.25: botton(2).dblH = 20: _
        botton(2).strOnAction = "転記": botton(2).strChaText = "転記"
    'ボタンの配置
    Set shtinput = ActiveWorkbook.Sheets("入力")
    For i = 0 To UBound(botton())
        Call MakeButton(shtinput, botton(i))
    Next
    Set shtinput = Nothing
End Function

Function initSyosiki(shtSource As Worksheet, shtDestination As Worksheet) As Boolean
'転記シートを消去し、原紙より新規にコピーする
Dim i As Long
Dim lngShapesCounter As Long

On Error GoTo Error
'転記シートの消去
    Call ClearSheet(shtDestination)
'原紙からのコピー
    shtSource.Cells.Copy Destination:=shtDestination.Cells
    shtDestination.Range("a1").Copy Destination:=shtDestination.Range("a1")
    shtDestination.ResetAllPageBreaks
    initSyosiki = True
    Exit Function
Error:
initSyosiki = False
End Function
Function initUtiwakeSyosiki(shtSrc As Worksheet, shtDst As Worksheet, page As Long) As Boolean
'見積内訳の初期配置を行う
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
'新規入力の初期化
Dim shtMy As Worksheet
Dim shtCopy As Worksheet
    Set shtMy = Sheets("表題")
    Call initShapes(shtMy)
    MakeHyoudaiButtons
    Set shtCopy = ActiveWorkbook.Sheets("入力原紙")
    Set shtMy = ActiveWorkbook.Sheets("入力")
    Call initSyosiki(shtCopy, shtMy)
    MakeInputButtons
End Function
