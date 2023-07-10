Attribute VB_Name = "Module1"
Sub foreachdemo()
Dim wSHT As Worksheet
For Each wSHT In Worksheets

wSHT.Cells(15, 7).Value = "Nice job!"
wSHT.Cells(15, 7).Interior.ColorIndex = 4

'MsgBox ("Found the sheet: " & wSHT.Name)
Next
End Sub
Sub reset()
Dim wSHT As Worksheet
For Each wSHT In Worksheets

wSHT.Cells(15, 7).Value = ""
wSHT.Cells(15, 7).Interior.ColorIndex = 0

Next
End Sub
