Attribute VB_Name = "Module1"
'�W�����氵�k
Sub DemoStep4New()
Attribute DemoStep4New.VB_ProcData.VB_Invoke_Func = "p\n14"
'�p�GD3-D24�x�s�檺�Ȭ��s�_�� ��
'today-�令���U��
Dim targetValue As Variant
targetValue = InputBox("�п�J�n�z���")



'end
'----------step4
Dim targetCId As Integer
targetCId = InputBox("�п�J���z�諸�ů���")

 

'���ӷ|�� typename
'If (targetCId = 1 Or targetCId = 3 Or targetCId = 5) Then
If IsNumeric(targetValue) Then
targetValue = CDbl(targetValue)
End If

 

Dim row As Integer
For row = 3 To 24

 

If Cells(row, targetCId).Value = targetValue Then
'�r�鬰����
Cells(row, targetCId).Font.ColorIndex = 3
'�I��������
Cells(row, targetCId).Interior.ColorIndex = 6

 

'�_�h
Else
'�r�鬰�¦�
Cells(row, targetCId).Font.ColorIndex = 1
'�I�����z��
Cells(row, targetCId).Interior.ColorIndex = 0

End If
Next
End Sub




'�����u�ƸѪk
Sub Demo0705Advance()
'�p�GD3-D24�x�s�檺�Ȭ��s�_�� ��
'today-�令���U��
Dim targetValue As Variant
targetValue = InputBox("�п�J�n�z���")
'end
'----------step4
Dim targetCId As Integer
targetCId = InputBox("�п�J���z�諸�ů���")

Dim dtRange As Range
Set dtRange = ActiveSheet.UsedRange
Dim rowNum As Long

rowNum = dtRange.Rows.Count
 

'���ӷ|�� typename (���ӴN�O����!!)�W���{���Ч�g�p�U
If (IsNumeric(targetValue)) Then
targetValue = CDbl(targetValue)
End If

 

Dim row As Integer
For row = 3 To rowNum

 

If Cells(row, targetCId).Value = targetValue Then
'�r�鬰����
Cells(row, targetCId).Font.ColorIndex = 3
'�I��������
Cells(row, targetCId).Interior.ColorIndex = 6

 

'�_�h
Else
'�r�鬰�¦�
Cells(row, targetCId).Font.ColorIndex = 1
'�I�����z��
Cells(row, targetCId).Interior.ColorIndex = 0

End If
Next
End Sub

