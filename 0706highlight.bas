Attribute VB_Name = "Module2"
Sub highlight()
Dim targetCol As Integer
targetCol = CInt(InputBox("Please insert your target column: "))
Dim targetValue As Double
targetValue = CDbl(InputBox("Please insert the highlight value: "))

Dim n, i As Long
n = Cells(Rows.Count, 1).End(xlUp).Row
i = Cells(n, column.Count).End(xlLeft).column

Dim rIdx As Long
For rIdx = 2 To n
    If Cells(rIdx, targetCol).Value > targetValue Then
        Cells(rIdx, targetCol).Font.ColorIndex = 3
        Cells(rIdx, targetCol).Interior.ColorIndex = 6
    Else
        Cells(rIdx, targetCol).Font.ColorIndex = 1
        Cells(rIdx, targetCol).Interior.ColorIndex = 0
    End If
Next
End Sub
