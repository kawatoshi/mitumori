Attribute VB_Name = "CommisionModule"
Option Explicit
Private Function getCommissionbox(shtDst As Worksheet, box As Long) As Range
'申請記入欄の先頭rangeを返す
Dim startRow As Long
Dim startCol As Long
    If box <= 0 Then Exit Function
    If box > 4 Then Exit Function
    If box Mod 2 = 1 Then
        startCol = 3
    Else
        startCol = 7
    End If
    If box / 3 < 1 Then
        startRow = 14
    Else
        startRow = 25
    End If
    With shtDst
        Set getCommissionbox = .Range(.Cells(startRow, startCol), .Cells(startRow + 7, startCol))
    End With
End Function
Private Function writeCommissionData(shtDst As Worksheet, bokMy As Workbook, _
                                     Hdata As HyoudaiData, box As Long) As Boolean
Dim rngBox As Range
    Set rngBox = getCommissionbox(shtDst, box)
    If Hdata.strMitumoriNo Like "" Then Exit Function
    rngBox.Cells(1) = Hdata.strMitumoriNo
    rngBox.Cells(2) = Hdata.strCustomer
    rngBox.Cells(3) = Hdata.strContents
    rngBox.Cells(4) = postWithTax(Hdata.dblSum, Hdata.dblTaxRate)
    rngBox.Cells(5) = postSiharaiTypeOnCommision(Hdata)
    rngBox.Cells(6) = postSagyouDate(Hdata)
    rngBox.Cells(7) = postGyousyaNames(Hdata)
    rngBox.Cells(8) = postGyousyaSumWithTax(Hdata)
    Call VisibleCommissionShapes(shtDst, Hdata.strMitumoriNo, box)
    Call postSumAndCost(bokMy, Hdata)
    Hdata.strMsinsei = CStr(Now())
writeCommissionData = reWriteHyoudaiData(Hdata, bokMy)
End Function
Private Function VisibleCommissionShapes(shtDst As Worksheet, MitumoriNo As String, box As Long) As Boolean
    If SinseiType_is_Mitumori(MitumoriNo) = True Then
        Call VisibleCommissionShape(shtDst, box, 1)
    End If
    If SinseiType_is_Seikyuu(MitumoriNo) = True Then
        Call VisibleCommissionShape(shtDst, box, 2)
    End If
End Function
Private Function VisibleCommissionShape(shtDst As Worksheet, box As Long, typeNum As Long) As Boolean
Dim strName As String
    strName = "maru" & box & "_" & typeNum
    shtDst.Shapes(strName).Visible = msoTrue
    VisibleCommissionShape = True
End Function
Private Function initCommissionShapes(shtDst As Worksheet) As Boolean
'捺印依頼書のShapesを初期化する
    Dim i As Long
    Dim lngCount As Long
    lngCount = shtDst.Shapes.Count
    For i = 1 To lngCount
        shtDst.Shapes(i).Visible = msoFalse
    Next
    initCommissionShapes = True
End Function
Private Function printCommission(shtCommission As Worksheet) As Boolean
    Application.Wait (Now + TimeValue("0:00:03"))
    shtCommission.PrintOut
    printCommission = True
End Function
Private Function makeCommissionBook(shtSrc As Worksheet, strFolder As String, strFile As String) As Workbook
'依頼書を別Bookで作成する
'usage call makeCommissionBook(shtSrc, "c:\mitumori", "test.xls")
Dim shtMy As Worksheet
Dim bokDst As Workbook
Dim ans As String
Dim shtDst As Worksheet
Dim MarginPoint As Double
    Set shtMy = ActiveSheet
    ans = OpenBook(strFolder, strFile, bokDst)
    If bokDst Is Nothing Then
        If is_match("ファイルが存在しません", ans) = True Then
            Set bokDst = Workbooks.Add
            bokDst.SaveAs (strFolder & "\" & strFile)
            shtMy.Activate
        Else
            Call MsgBox(ans, vbOKOnly, "makeCommissionBook")
            Exit Function
        End If
    End If
    Set shtDst = bokDst.Sheets.Add(after:=Sheets(bokDst.Sheets.Count))
    If initSyosiki(shtSrc, shtDst) = False Then
        Call MsgBox("何故か転記できません、管理者に相談してください", vbOKOnly, "makeCommissionBook")
        Exit Function
    End If
    With shtDst.PageSetup
        .LeftMargin = Application.InchesToPoints(0.590551181102362)
        .RightMargin = Application.InchesToPoints(0.31496062992126)
        .TopMargin = Application.InchesToPoints(0.511811023622047)
        .BottomMargin = Application.InchesToPoints(0.393700787401575)
        .HeaderMargin = Application.InchesToPoints(0.354330708661417)
        .FooterMargin = Application.InchesToPoints(0.196850393700787)

    End With
    Set makeCommissionBook = bokDst
End Function

Sub visibleCommisionSignature(shtDst As Worksheet, bolVisible As Boolean)
'申請印を表示、非表示する
Dim signature As Variant
    If bolVisible = True Then
        signature = msoTrue
    Else
        signature = msoFalse
    End If
    shtDst.Shapes(9).Visible = signature
    shtDst.Shapes(10).Visible = signature
End Sub
Function clearCommissionData(shtDst As Worksheet) As Boolean
'捺印依頼書を初期化する
Dim rngBox As Range
Dim cell As Variant
Dim i As Long
    For i = 1 To 4
        Set rngBox = getCommissionbox(shtDst, i)
        For Each cell In rngBox
            cell.Value = ""
        Next
    Next
    Call initCommissionShapes(shtDst)
    clearCommissionData = True
End Function
Function CommissionShapesVisible(shtDst As Worksheet) As Boolean
'捺印依頼書のShapesを全て表示する
    Dim i As Long
    Dim lngCount As Long
    lngCount = shtDst.Shapes.Count
    For i = 1 To lngCount
        shtDst.Shapes(i).Visible = msoTrue
    Next
    CommissionShapesVisible = True
End Function
Function WriteCommission(shtDst As Worksheet, Mitumorinos() As String, signature As Boolean) As Boolean
'捺印依頼書のboxに依頼内容を記載する
Dim Hdata() As HyoudaiData
Dim i As Long
Dim counter As Long
Dim bokMy As Workbook
    Set bokMy = ActiveWorkbook
    ReDim Hdata(UBound(Mitumorinos))
    counter = 1
    Call clearCommissionData(shtDst)
    Call visibleCommisionSignature(shtDst, signature)
    For i = 0 To UBound(Mitumorinos)
        If counter > 4 Then Exit For
        Hdata(i) = getHyoudaiData(Mitumorinos(i))
        If writeCommissionData(shtDst, bokMy, Hdata(i), counter) = False Then
            If Not Hdata(i).strMitumoriNo Like "" Then
                Call MsgBox(Hdata(i).strMitumoriNo & "の書き込みに異常がありました")
            End If
        End If
        counter = counter + 1
    Next
    WriteCommission = True
End Function
Function publishCommission(shtSrc As Worksheet, _
                           strBokName As String) As Boolean
'捺印依頼書を発行する
'strBokNameが有効ならば指定bookを作成してそこに発行
'それ以外はOutPutPrinterで印刷する
    Select Case strBokName
    Case ""
        Call printCommission(shtSrc)
    Case Else
        Call makeCommissionBook(shtSrc, getDeskTopPath, strBokName)
    End Select
End Function
