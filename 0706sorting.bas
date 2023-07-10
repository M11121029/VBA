Attribute VB_Name = "Module1"
Sub sorting()
Attribute sorting.VB_ProcData.VB_Invoke_Func = "b\n14"
' Ctrl + b
    Range("B1").Select
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    n = Cells(Rows.Count, 2).End(xlUp).Row
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:=Range(Cells(2, 2), Cells(n, 2)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Ascending***
        
    'with with:
    With ActiveWorkbook.Worksheets(1).Sort
        '.SetRange Range("A1:B414")***
        .SetRange Range(Cells(1, 1), Cells(n, 2)) '***
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'without with:
    'ActiveWorkbook.Worksheets(1).Sort.SetRange Range(Cells(1, 1), Cells(n, 2))
    'ActiveWorkbook.Worksheets(1).Sort.Header = xlYes......
    
End Sub
Sub sorting2()
Attribute sorting2.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("B1").Select
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    n = Cells(Rows.Count, 2).End(xlUp).Row
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add2 Key:=Range(Cells(2, 2), Cells(n, 2)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal 'Descending***
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range(Cells(1, 1), Cells(n, 2))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
