Attribute VB_Name = "tools"
Option Explicit

Function getRanges2()
Attribute getRanges2.VB_ProcData.VB_Invoke_Func = "b\n14"
'�I�����ꂽ�͈͂̓��e����荞��ŃC�~�f�B�G�C�g�ɏo�͂���p
On Error GoTo HandleErr
Dim rngSelection As Range
Dim rngMy As Range
Dim strValue As String
Dim strAdd As String
Dim lngcolumn As Long
Dim str() As String
    Set rngSelection = Selection
    ReDim str(1)
    str(0) = "set comrange = shtmy.range("""
    str(1) = " = .cells(lnggetrow, "
    For Each rngMy In rngSelection
        strValue = rngMy.Value
        lngcolumn = rngMy.Column
        Debug.Print str(0) & rngMy.Address & """)"
    Next
    Exit Function
HandleErr:
    Debug.Print Err.Number & ": " & Err.Description & " tools.getRanges2"
    Resume Next
End Function

Function getRanges()
'�I�����ꂽ�͈͂̓��e����荞��ŃC�~�f�B�G�C�g�ɏo�͂���
'public const�p
Dim rngSelection As Range
Dim rngMy As Range
Dim strValue As String
Dim strAdd As String
Dim lngcolumn As Long
Dim str() As String
    Set rngSelection = Selection
    ReDim str(1)
    str(0) = "public const "
    str(1) = " as long = "
    For Each rngMy In rngSelection
        strValue = rngMy.Value
        lngcolumn = rngMy.Column
        Debug.Print str(0) & strValue & str(1) & CStr(lngcolumn)
    Next
End Function
Function trimRanges()
'�I�����ꂽ�͈͂̓��e����荞��ŋ󔒂��������ď����߂�
Dim rngSelection As Range
Dim rngMy As Range
    Set rngSelection = Selection
    For Each rngMy In rngSelection
        rngMy.Value = Trim(rngMy.Value)
    Next
End Function
Function chkMitumoriNo()
'�I��͈͂̌���No���`�F�b�N���ĕϊ�����
Dim Mno As MitumoriNumber
Dim rngSelection As Range
Dim rngMy As Range
    Set rngSelection = Selection
    Set Mno = New MitumoriNumber
    For Each rngMy In rngSelection
        If Mno.Push(rngMy.Value) = True Then
            rngMy.Value = Mno.Publish
        Else
            rngMy.Value = "error: " & rngMy.Value
        End If
    Next
End Function
