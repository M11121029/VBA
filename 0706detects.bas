Attribute VB_Name = "Module1"
Sub detect()
Dim dtRange As Range
Set dtRange = ActiveSheet.UsedRange
Dim rowNum, columnNum As Long
rowNum = dtRange.Rows.Count
columnNum = dtRange.Columns.Count
MsgBox ("Rows: " & rowNum)
MsgBox ("Columns: " & columnNum)
End Sub
Sub detect2()
Dim dtRange As Range
Set dtRange = ActiveSheet.UsedRange
Dim n, i As Long
n = Cells(Rows.Count, 1).End(xlUp).Row
i = Cells(n, Columns.Count).End(xlToLeft).column
MsgBox ("Rows: " & n)
MsgBox ("Columns: " & i)
End Sub
