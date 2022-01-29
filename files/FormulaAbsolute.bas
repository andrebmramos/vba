Sub FormulaAbsolute()
' Purpose: Convert selected values to absolute

    Dim rng As Range
    Dim myFormula As String
    Dim cellValue As Double
    Set rng = Selection

    For Each c In rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            c.Formula = "=ABS(" & myFormula & ")"
        Else
            cellValue = c.Value
            c.Formula = "=ABS(" & cellValue & ")"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c
    
End Sub