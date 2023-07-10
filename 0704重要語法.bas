Attribute VB_Name = "Module1"
Sub selectdemo()
Select Case Range("A2").Value
    Case "Guava"
        Range("B2").Value = "Ian Tsau"
    Case "O-King Kong Black Peanut"
        Range("B2").Value = "Tuku"
    Case "Milkfish"
        Range("B2").Value = "Xuejia"
    Case "Talents of VBA"
        Range("B2").Value = "Douliu"
End Select
End Sub
Sub ifdemo()
If Range("B1").Value > 38 Then
    Range("B2").Value = "Positive"
Else
    Range("B2").Value = "Negative"
End If
End Sub
Sub elseifdemo()
Dim i, n As Integer
i = 2
n = Cells(Rows.Count, 4).End(xlUp).Row
Do While i <= n
If Cells(i, 5).Value = "Setosa" Then
    Cells(i, 6).Value = "1"
ElseIf Cells(i, 5).Value = "Versicolor" Then
    Cells(i, 6).Value = "2"
Else
    Cells(i, 6).Value = "3"
End If
i = i + 1
Loop

End Sub
Sub Resetdemo()
Attribute Resetdemo.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("F2:F151").Select
    Selection.ClearContents
    'Selection.Value = ""
    Range("F2").Select
End Sub
Sub fordemo()
Dim i, s As Integer ' We can dim 2 things at the same time
s = 0
For i = 1 To 100 ' step = 1
s = s + i
    If i >= 10 Then
        Exit For
    End If
Next i
MsgBox ("s = " & s)

End Sub
Sub countdemo() ' For counting rows in the sheet.
No_of_rows = Cells(Rows.Count, 3).End(xlUp).Row
MsgBox (No_of_rows)
End Sub
Sub forapp()
n = Cells(Rows.Count, 3).End(xlUp).Row
Dim rIdx As Integer
For rIdx = 2 To n
Cells(rIdx, 4).Value = Cells(rIdx, 2).Value * Cells(rIdx, 3).Value
Next
End Sub
Sub Reset2demo()
    Range("D2:D100").Select
    Selection.ClearContents
    Range("E2").Select
End Sub
