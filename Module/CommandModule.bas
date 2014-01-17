Attribute VB_Name = "CommandModule"
Option Explicit
'シートに配置されているボタンはすべてこのモジュールに
'記載されているsubプロシージャーから起動される

Public Sub version()
    Call MsgBox("version: 2.7.4" & Chr(13) & "code by kawakita  2011.03.17", vbOKOnly, "version")
End Sub

Sub 新規入力()
    Call initInputSheet
    ActiveWorkbook.Sheets("入力").Select
    Cells(2, 2).Select
End Sub

Sub 再入力()
'表題シートの内容を再見積もりする
Dim ans As String
    ans = rekeyMitumori(getMitumoriNo)
    If Not ans Like "" Then
        MsgBox (ans)
    Else
        Sheets("入力").Activate
        Range("f2").Select
    End If
End Sub

Sub 転記()
Dim bokMy As Workbook
Set bokMy = ActiveWorkbook
    If SheetInput(bokMy) = True Then
        Call initInputSheet
        bokMy.Sheets("表題").Activate
        getHyoudaiEndCell(bokMy).Select
    End If
End Sub

Sub 申請確認()
'申請書を作成する
Dim Mno As String
    Mno = getMitumoriNo
    If SinseiKakunin(Mno) = False Then
        Call MsgBox("申請データに問題があるため、作成できませんでした")
        Exit Sub
    End If
    If SinseiType_is_Seikyuu(Mno) = True Then
        Sheets("請求書").Activate
        Exit Sub
    End If
    If SinseiType_is_Mitumori(Mno) = True Then
        Sheets("見積書").Activate
        Exit Sub
    End If
    Call MsgBox("申請を作成できませんでした", vbOKOnly)
End Sub

Sub 捺印依頼作成()
'通常出力
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
Sub 依頼特殊()
'プリンタ指定印刷
Dim Mnos() As String
Dim strOutPut As String
    Mnos = getMitumoriNoOnRangeAreas
    strOutPut = Range("toku_printer").Value
    Call MakeCommissions(Mnos, strOutPut, "", False)
End Sub
Sub 依頼Book()
'Book出力
Dim Mnos() As String
Dim strBook As String
    Mnos = getMitumoriNoOnRangeAreas
    strBook = Range("commission_file")
    Call MakeCommissions(Mnos, "", strBook, True)
End Sub
Sub 申請書類変更()
        If getMitumoriNo Like "" Then
            MsgBox ("選択行では申請を変更できません")
        Else
            frmSinsei.Show
        End If
End Sub
Public Sub 本部申請()
'表題シートの発行申請タイプから適切な申請を自動で行う
Dim Mnos() As String
Dim i As Long
    Mnos = getMitumoriNoOnRangeAreas
    For i = 0 To UBound(Mnos)
        Call MakeSinseiToHQ(Mnos(i))
    Next
End Sub
Sub 取り込み()
'入力シートへの取り込み
Dim ans As String
    ans = TorikomiSheetInput
    If Not ans Like "" Then
        Call MsgBox(ans)
    End If
End Sub
Sub 定期作成()
'定期見積を転記する
Dim lngMonth As Long
Dim strMonth As String
Dim bokMy As Workbook
Dim ans As String
    lngMonth = Month(Now())
    strMonth = InputBox("作成月を入力してください", "定期作成", CStr(lngMonth))
    Set bokMy = ActiveWorkbook
    ans = makeTeiki(CLng(strMonth), bokMy)
    Call MsgBox(ans, vbOKOnly, "定期作成結果")
End Sub
