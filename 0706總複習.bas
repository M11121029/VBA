Attribute VB_Name = "Module1"
'上次應急做法
Sub DemoStep4New()
Attribute DemoStep4New.VB_ProcData.VB_Invoke_Func = "p\n14"
'如果D3-D24儲存格的值為新北市 時
'today-改成式萬用
Dim targetValue As Variant
targetValue = InputBox("請輸入要篩選值")



'end
'----------step4
Dim targetCId As Integer
targetCId = InputBox("請輸入欲篩選的藍索引")

 

'未來會教 typename
'If (targetCId = 1 Or targetCId = 3 Or targetCId = 5) Then
If IsNumeric(targetValue) Then
targetValue = CDbl(targetValue)
End If

 

Dim row As Integer
For row = 3 To 24

 

If Cells(row, targetCId).Value = targetValue Then
'字體為紅色
Cells(row, targetCId).Font.ColorIndex = 3
'背景為黃色
Cells(row, targetCId).Interior.ColorIndex = 6

 

'否則
Else
'字體為黑色
Cells(row, targetCId).Font.ColorIndex = 1
'背景為透明
Cells(row, targetCId).Interior.ColorIndex = 0

End If
Next
End Sub




'今天優化解法
Sub Demo0705Advance()
'如果D3-D24儲存格的值為新北市 時
'today-改成式萬用
Dim targetValue As Variant
targetValue = InputBox("請輸入要篩選值")
'end
'----------step4
Dim targetCId As Integer
targetCId = InputBox("請輸入欲篩選的藍索引")

Dim dtRange As Range
Set dtRange = ActiveSheet.UsedRange
Dim rowNum As Long

rowNum = dtRange.Rows.Count
 

'未來會教 typename (未來就是今天!!)上次程式請改寫如下
If (IsNumeric(targetValue)) Then
targetValue = CDbl(targetValue)
End If

 

Dim row As Integer
For row = 3 To rowNum

 

If Cells(row, targetCId).Value = targetValue Then
'字體為紅色
Cells(row, targetCId).Font.ColorIndex = 3
'背景為黃色
Cells(row, targetCId).Interior.ColorIndex = 6

 

'否則
Else
'字體為黑色
Cells(row, targetCId).Font.ColorIndex = 1
'背景為透明
Cells(row, targetCId).Interior.ColorIndex = 0

End If
Next
End Sub

