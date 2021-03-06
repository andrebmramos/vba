Sub FormulaRound()
' Purpose: Convert selected values to absolute

    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    For Each c In rng
        If c.HasFormula = True Then
            myFormula = Right(c.Formula, Len(c.Formula) - 1)
            c.Formula = "=ROUND(" & myFormula & ",0)"
        Else
            c.Formula = "=ROUND(" & c.Value & ",0)"
        End If
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Next c

End Sub