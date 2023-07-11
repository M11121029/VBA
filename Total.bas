Attribute VB_Name = "Module1"
Sub preProc()
Dim n As Long
n = Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long
i = Cells(1, Columns.Count).End(xlToLeft).Column
MsgBox ("There are " & n & " rows and " & i & " columns.")
End Sub
Sub acc()
Dim sum As Variant
Dim rIdx As Long
Dim cIdx As Long

Dim n As Long
n = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Long
i = Cells(1, Columns.Count).End(xlToLeft).Column

For cIdx = 1 To i
    sum = 0
    For rIdx = 2 To n
    If IsNumeric(Cells(rIdx, cIdx).Value) = True Then
        sum = sum + Cells(rIdx, cIdx).Value
    End If
    Next
    MsgBox ("The total of " & Cells(1, cIdx).Value & " is: " & sum)
Next

End Sub
