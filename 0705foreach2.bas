Attribute VB_Name = "Module2"
Sub foreachdemo2()
Dim usrStr As String
usrStr = InputBox("Please insert the name of sheet that you're going to active: ", "Hint: ")
Dim wSHT As Worksheet
For Each wSHT In Worksheets
    If wSHT.Name = usrStr Then
        wSHT.Activate
        MsgBox ("Done")
    Else
        MsgBox ("Sheet: " & ustStr & "is not found.")
    End If
    Next
End Sub
