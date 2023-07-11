Attribute VB_Name = "Module11"
Sub dataInsertValue()
    Dim rIdx As Long
    Dim cIdx As Long
    Dim customizedChar As String
    Dim customizedCInt As Long

    Dim strartRow As Integer
    startRow = CInt(InputBox("Please insert the starting row: "))
    Dim dtRange As Range
    Set dtRange = ActiveSheet.UsedRange
    rowNum = dtRange.Rows.Count
    colNum = dtRange.Columns.Count

    customizedCnt = 0
    customizedChar = InputBox("Please insert what the missing values look like: ")
    For cIdx = 1 To colNum
        For rIdx = 1 To rowNum
            If Trim(Cells(rIdx, cIdx).Value) Like Trim(customizedChar) Then
                customizedCnt = customizedCnt + 1
            End If
    Next
        Dim curColName As String
        If IsNumeric(Cells(startRow, cIdx).Value) Then
            curColName = Split(Cells(1, cIdx).Address(True, False), "$")(0)
            MsgBox (curColName)
            Dim maxStr, minStr As String
            maxStr = "=Max('" & Sheets(1).Name & "'!" & curColName & ":" & curColName & ")"
            minStr = "=Min('" & Sheets(1).Name & "'!" & curColName & ":" & curColName & ")"
            Dim maxValue As Double
            Dim minValue As Single
            Sheets(2).Cells(1, 1).Value = "MAX"
            Sheets(2).Cells(2, 1).Value = "Min"
            Sheets(2).Cells(1, cIdx).Formula = maxStr
            Sheets(2).Cells(2, cIdx).Formula = minStr
        End If
    Next
    If customizedCnt > 0 Then
        MsgBox "There are " & customizedCnt & " empty values."
    Else
        MsgBox ("There is no missing value in this dataset.")
    End If
   
End Sub

