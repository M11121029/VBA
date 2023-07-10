Attribute VB_Name = "Module2"
Sub SAdemo1()
Attribute SAdemo1.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G2").Select
End Sub
Sub SAdemo2()
Dim i As Integer
i = InputBox("Please insert the target column: ")
n = Cells(Rows.Count, 2).End(xlUp).row
Range("E1").Value = 0
Dim bbb As Integer
For bbb = 2 To n
    Range("E1").Value = Range("E1").Value + Cells(bbb, i)
Next
Range("G1").Value = Range("E1").Value / (n - 1)
End Sub
Sub reset()
Range("E1").Value = 0
Range("G1").Value = 0
End Sub
