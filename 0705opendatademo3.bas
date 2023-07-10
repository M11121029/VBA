Attribute VB_Name = "Module2"
Sub demo3()
Cells.Select
    Cells.Font.ColorIndex = 1
    Cells.Interior.ColorIndex = 0
    Cells.Font.Bold = False
Range("G13").Select

Dim targetvalue As String
targetvalue = InputBox("Please insert your target: ", "Target input")

Dim targetCId As Integer
targetCId = InputBox("Please insert your target row: ", "Column input")

Dim rIdx, n As Integer
n = Cells(Rows.Count, targetCId).End(xlUp).Row
For rIdx = 3 To n
If Cells(rIdx, targetCId).Value = targetvalue Then
    Cells(rIdx, targetCId).Font.ColorIndex = 3
    Cells(rIdx, targetCId).Font.Bold = True
    Cells(rIdx, targetCId).Interior.ColorIndex = 43
Else
    'Cells(rIdx, targetCId).Font.ColorIndex = 1
    'Cells(rIdx, targetCid).font.bold = False
    'Cells(rIdx, targetCId).Interior.ColorIndex = 0
End If
Next

End Sub
