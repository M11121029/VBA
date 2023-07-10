Attribute VB_Name = "Module3"
Sub dowhileloop()
Dim i, s, x As Integer
s = 0
i = 1
x = InputBox("Please insert the X")
Do While i <= x
    s = s + i
    i = i + 1
Loop
MsgBox ("s = " & s)
End Sub
