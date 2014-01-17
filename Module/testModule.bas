Attribute VB_Name = "testModule"
Option Explicit
Private Sub testGetHyoudaiData()
Dim MitumoriNo As String
Dim Hdata As HyoudaiData
MitumoriNo = "22A-0346"
    Hdata = getHyoudaiData(MitumoriNo)
End Sub

Private Sub testGetSyousaiData()
Dim MitumoriNo As String
Dim Sdata() As SyousaiData
MitumoriNo = "23A-0233"
    Sdata = getSyousaiData(MitumoriNo)
End Sub

Private Sub testGetGyousyaData()
Dim MitumoriNo As String
Dim Gdata() As GyousyaData
MitumoriNo = "22A-0346"
    Gdata = getGyousyaData(MitumoriNo)
End Sub

Private Sub testGetUtiwakeData()
Dim MitumoriNo As String
Dim Udata() As UtiwakeData
MitumoriNo = "23A-0090"
    Udata = getUtiwakeData(MitumoriNo)
End Sub

Private Sub testMitumoriPush()
Dim Mno As MitumoriNumber
Dim mitumoris(10) As String
Dim mitumori As Variant
mitumoris(0) = "23A-0004"
mitumoris(1) = "23KK-0255b"
mitumoris(2) = "23A-0430-1112"
mitumoris(3) = ""
mitumoris(4) = "99-010001"
mitumoris(5) = "24-0334f-1503"
mitumoris(6) = "23A-0334A-1112"
mitumoris(7) = "23A-0004-8"
mitumoris(8) = "23A-110039-2"
mitumoris(9) = "23A-110039-2-23-01"
mitumoris(10) = "23K-110139-23-01"

Set Mno = New MitumoriNumber
    For Each mitumori In mitumoris
        Debug.Print Mno.Push(CStr(mitumori)) & ": " & _
                    "source:" & Mno.source & _
                    " return:" & Mno.Publish & _
                    " type:" & Mno.MitumoriType
        If Not Mno.MitumoriType Like "empty" Then
'            Debug.Print "year is: " & mno.year
            Debug.Print "main No is: " & Mno.MainNo
'            Debug.Print "include old number? " & mno
        End If
'        Debug.Print ""
    Next
End Sub

Private Sub testFindReMitumoriNumbers()
Dim MitumoriNo(5) As String
Dim No As Variant
MitumoriNo(0) = "23A-0129"
MitumoriNo(1) = "23A-0119"
MitumoriNo(2) = "23A-0129z"
MitumoriNo(3) = "23A-0129-1903"
MitumoriNo(4) = "23A-01290"
MitumoriNo(5) = "23A-9828"
    For Each No In MitumoriNo
        Debug.Print "input: " & No
        Debug.Print "anser: " & postMaxSaiMitumoriNo(CStr(No)).Publish
    Next
End Sub

Private Sub testNewInput()
    新規入力
End Sub

Private Sub testSheetInputHyoudai()
Dim bokMy As Workbook
Dim Hdata As HyoudaiData
    Set bokMy = ActiveWorkbook
    Hdata = getSheetInputHyoudai(bokMy)
End Sub

Private Sub testSheetInputSyousaiData()
Dim bokMy As Workbook
Dim Sdata() As SyousaiData
    Set bokMy = ActiveWorkbook
    Sdata = getSheetInputSyousaiData(bokMy, "testNO")
End Sub

Private Sub testSheetInputUtikakePage()
Dim bokMy As Workbook
Dim Udata() As UtiwakeData
    Set bokMy = ActiveWorkbook
    Udata = getSheetInputUtiwakePage(bokMy, 1, "testNO")
End Sub

Private Sub testSheetInputGyousyaData()
Dim bokMy As Workbook
Dim Gdata() As GyousyaData
    Set bokMy = ActiveWorkbook
    Gdata = getSheetInputGyousyaData(bokMy, "testNO")
End Sub

Private Sub testPostMitumoriNo()
Dim Mno As String
Dim mTypes(3) As String
Dim varType As Variant
mTypes(0) = ""
mTypes(1) = "新規"
mTypes(2) = "再見積"
mTypes(3) = "定期"
    For Each varType In mTypes
        Debug.Print postMitumoriNo("23A-0210a", CStr(varType))
    Next
End Sub

Private Sub testReKeyMitumori()
Dim ans As String
    Call initInputSheet
    ans = rekeyMitumori("23A-0129b")
End Sub
Private Sub testInitUtiwakeSyosiki()
Dim shtSrc As Worksheet
Dim shtDst As Worksheet
Dim page As Long
    Set shtSrc = Sheets("内訳原紙")
    Set shtDst = Sheets("見積書")
    page = 3
    Call initUtiwakeSyosiki(shtSrc, shtDst, page)
End Sub
Private Sub testPublishMitumori()
Dim shtSrc As Worksheet
Dim shtDst As Worksheet
Dim Mno As String
    Set shtSrc = Sheets("見積原紙")
    Set shtDst = Sheets("見積書")
    Mno = "23A-0129b"
'    mno = ""
   Mno = "dammy"
    Call publishMitumori(Mno, shtSrc, shtDst)
    shtDst.Activate
End Sub
Private Sub testPublishSeikyuu()
Dim shtSrc As Worksheet
Dim shtDst As Worksheet
Dim Mno As String
    Set shtSrc = Sheets("請求原紙")
    Set shtDst = Sheets("請求書")
    Mno = "23A-0129b"
'    mno = ""
'    mno = "dammy"
    Call publishSeikyuu(Mno, shtSrc, shtDst)
    shtDst.Activate
End Sub
Private Sub testHyoudaiEndActive()
Dim bokMy As Workbook
Set bokMy = ActiveWorkbook
    bokMy.Sheets("表題").Activate
    getHyoudaiEndCell(bokMy).Select
End Sub
Private Sub testClearCommission()
Dim shtDst As Worksheet
    Set shtDst = ActiveWorkbook.Sheets("捺印依頼書")
    Call clearCommissionData(shtDst)
End Sub
Private Sub testWriteCommission()
Dim shtDst As Worksheet
Dim Mno(3) As String
    Call testClearCommission
    Set shtDst = ActiveWorkbook.Sheets("捺印依頼書")
    Mno(0) = "23A-0129b"
    Mno(1) = "23A-0151"
    Mno(2) = "23A-0240"
    Mno(3) = "23A-0229"
    Call WriteCommission(shtDst, Mno, True)
End Sub
Private Sub testMakeCommission()
Dim shtDst As Worksheet
Dim Mno(1) As String
    Mno(0) = "23A-0129b"
    Mno(1) = "23A-0151"
    Call MakeCommissions(Mno, "Bullzip PDF Printer", "", True)
End Sub
Private Sub testMakeCommission2()
Dim shtDst As Worksheet
Dim Mno(7) As String
    Mno(0) = "23A-0129b"
    Mno(1) = "23A-0151"
    Mno(2) = "23A-0240"
    Mno(3) = "23A-0229"
    Mno(4) = "23A-0230"
    Mno(5) = "23A-0231"
    Mno(6) = "23A-0232"
    Mno(7) = "23A-0233"
    Call MakeCommissions(Mno, "Bullzip PDF Printer", "test", True)
End Sub
Private Sub testRepostHyoudai()
Dim bokMy As Workbook
Dim Hdata As HyoudaiData
    Set bokMy = ActiveWorkbook
    Hdata = getHyoudaiData("23A-0164")
    Hdata.strNotes = CStr(Now())
    Debug.Print reWriteHyoudaiData(Hdata, bokMy)
End Sub
Private Sub testOpenBook()
Dim bokMy As Workbook
Dim Folder(3) As String
Dim strFile(3) As String
Dim ans(3) As String
Dim i As Long
    Folder(0) = "x:\"
    Folder(1) = "d:\mitumoritest"
    Folder(2) = "d:\mitumoritest"
    Folder(3) = "d:\mitumoritest"
    strFile(0) = "test"
    strFile(1) = "book.xls"
    strFile(2) = "bookopentest.xls"
    strFile(3) = "test"
    ans(0) = "申請するディレクトリーが存在しないかネットワークに接続されていませんERROR"
    For i = 0 To UBound(Folder)
        Debug.Print OpenBook(Folder(i), strFile(i), bokMy)
    Next
End Sub
Private Sub testGetTeikiHyoudaiData()
Dim bokMy As Workbook
Dim shtMy As Worksheet
Dim strMno() As String
Dim Hdata() As HyoudaiData
Dim ans As String
    Set bokMy = ActiveWorkbook
    Set shtMy = bokMy.Sheets("定期表題")
    strMno = findTeikiMitumoriNumbers(1, bokMy)
    If UBound(strMno) < 0 Then
        Call MsgBox("nodata")
        Exit Sub
    End If
    Hdata() = getTeikiHyoudaiDatas(strMno, bokMy)
    ans = makeTeiki(1, bokMy)
    Call MsgBox(ans)
End Sub
