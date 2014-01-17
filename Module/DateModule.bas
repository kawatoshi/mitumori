Attribute VB_Name = "DateModule"
Option Explicit


Function firstDay(dte As Date) As Date
    firstDay = DateSerial(year(dte), Month(dte), 1)
End Function
Function endOfMonth(dte As Date) As Date
Dim dteNext As Date
Dim dteFirst As Date
    dteNext = DateAdd("m", 1, dte)
    endOfMonth = firstDay(dteNext) - 1
End Function
