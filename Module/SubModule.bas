Attribute VB_Name = "SubModule"
Option Explicit
Private Function chkFolder(strPath As String) As String
'引数で渡されたフォルダまたはファイルが存在するかを
'確認し、結果を返す:folder:file:NG
Dim lngERROR As Long
'フォルダのチェック
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
        BookSave = "保存できませんでした"
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
'指定フォルダ内指定ブックが存在しているかを確認し別フォルダ同名ファイルが存在しないことを
'確認して開く
    'ネットワークディレクトリとファイルの確認
    If chkFolder(strReqFolder) <> "folder" Then
        OpenBook = "申請するディレクトリーが存在しません" & Chr(13) & _
                   "folder: " & strReqFolder
        Exit Function
    End If
    If chkFolder(strReqFolder & "\" & strReqFile) <> "file" Then
        OpenBook = "申請するファイルが存在しません" & Chr(13) & _
                   "file: " & strReqFile
        Exit Function
    End If
    '同名ブックが開いているかの確認
    Set bokDst = Book_is_opend(strReqFile)
    If bokDst Is Nothing Then
         Set bokDst = Workbooks.Open(Filename:=strReqFolder & "\" & strReqFile, UpdateLinks:=0, ReadOnly:=False)
    Else
         '別フォルダの同名ファイルが開いていないかの確認
        If Not LCase(bokDst.Path) Like LCase(strReqFolder) Then
            OpenBook = "同名のファイルがすでに開いています。 " & Chr(13) & _
                       "file: " & strReqFile & "を閉じてください"
            Set bokDst = Nothing
            Exit Function
        End If
        '読み取り専用で開いていないかの確認
        If bokDst.ReadOnly = True Then
            OpenBook = "読み取り専用で開かれているので書き込み出来ませんERROR"
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
'申請書を開き､見積NOと同名のシートを作成､転記する｡
'strSinseiType: "m"見積　"s"請求
Dim bokDst As Workbook
Dim ans As String
Dim shtDst As Worksheet
Dim shtSrc As Worksheet
'指定ワークブックを開く
    ans = OpenBook(strFolder, strFile, bokDst)
    If bokDst Is Nothing Then
        MakeSinsei = ans
        Exit Function
    End If
'同名のシートが存在した場合の確認
    If has_same_sheet(Hdata.strMitumoriNo, bokDst) = True Then
        If Not MsgBox(Hdata.strMitumoriNo & "はすでに申請されています。書き換えますか?", vbYesNo) = vbYes Then
            MakeSinsei = "書き換え中止"
            Exit Function
        End If
        Set shtDst = bokDst.Sheets(Hdata.strMitumoriNo)
    End If
'請求タイプ別の転記
    If shtDst Is Nothing Then
        Set shtDst = bokDst.Sheets.Add(after:=bokDst.Worksheets(bokDst.Sheets(bokDst.Sheets.Count).Name))
        shtDst.Name = Hdata.strMitumoriNo
        Call printerSetUp(shtDst)
    End If
    Select Case strSinseiType
    Case "m"
        Set shtSrc = bokMy.Sheets("見積原紙")
        If publishMitumori(Hdata.strMitumoriNo, shtSrc, shtDst) = False Then
            ans = "申請転記にエラーが発生しました。管理者に相談してください" & Chr(13) & _
                  "見積No : " & Hdata.strMitumoriNo & "   " & shtSrc.Name
        End If
    Case "s"
        Set shtSrc = bokMy.Sheets("請求原紙")
        If publishSeikyuu(Hdata.strMitumoriNo, shtSrc, shtDst) = False Then
            ans = "申請転記にエラーが発生しました。管理者に相談してください" & Chr(13) & _
                  "見積No : " & Hdata.strMitumoriNo & "   " & shtSrc.Name
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
