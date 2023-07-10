Attribute VB_Name = "Module1"
Sub demo1()
Dim targetvalue As String
targetvalue = InputBox("Please insert the target value: ")
If Cells(3, 4).Value = targetvalue Then
    Cells(3, 4).Font.ColorIndex = 3
    Cells(3, 4).Interior.ColorIndex = 6
Else
    Cells(3, 4).Font.ColorIndex = 1
    Cells(3, 4).Interior.ColorIndex = 0
End If
End Sub
Sub demo2()
Dim targetvalue As Variant
targetvalue = InputBox("Please insert your target: ")
Dim rIdx, n As Integer
n = Cells(Rows.Count, 4).End(xlUp).Row
For rIdx = 3 To n
If Cells(rIdx, 4).Value = targetvalue Then
    Cells(rIdx, 4).Font.ColorIndex = 3
    Cells(rIdx, 4).Interior.ColorIndex = 43
Else
    Cells(rIdx, 4).Font.ColorIndex = 1
    Cells(rIdx, 4).Interior.ColorIndex = 0
End If
Next
End Sub
