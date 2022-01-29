Sub DeleteNamedRanges()
'   Purpose: Delete all name ranges

    Dim i As Long
    
    Application.Calculation = xlCalculationManual
    For i = ThisWorkbook.Names.Count To 1 Step -1
        ThisWorkbook.Names(i).Delete
    Next
    Application.Calculation = xlCalculationAutomatic

End Sub