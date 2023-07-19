Attribute VB_Name = "Module2"
Sub multifilesmix()
    Dim Path As String
    Dim fileName As String
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    'Dim wsSource As Worksheet
    Dim n As Long
    Dim headerRange As Range
    Dim copyRange As Range

    Set wbTarget = Workbooks.Add
    Set wsTarget = wbTarget.Sheets(1)

    Path = "D:\dealrecords\"

    fileName = Dir(Path & "Date*.xlsx")
    Do While fileName <> ""

        Set wbSource = Workbooks.Open(Path & fileName)
        Set wsSource = wbSource.Sheets(1)

        Set headerRange = wsSource.Rows(1)
        n = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        headerRange.Copy wsTarget.Cells(1, 1)

        n = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
        Set copyRange = wsSource.Range("A2").Resize(n - 1, wsSource.UsedRange.Columns.Count)
        n = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        copyRange.Copy wsTarget.Cells(n + 1, 1)

        wbSource.Close SaveChanges:=False
        fileName = Dir
    Loop
    wbTarget.SaveAs "D:\dealrecords\NewRecords.xlsx"
    wbTarget.Close

    'Set wsTarget = Nothing
    'Set wbTarget = Nothing
    'Set wsSource = Nothing
    'Set wbSource = Nothing
    
    MsgBox "Done. "
End Sub

