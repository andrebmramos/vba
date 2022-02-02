Sub YMW_InsertColumnWidth()
'   Purpose: Insert column width counter

    Dim rng As Range
    Dim myFormula As String
    Set rng = Selection

    For Each c In rng
        c.Formula = "=" & "XCOLUMNWIDTH(" & c.Address & ")"
        c.WrapText = False
        c.HorizontalAlignment = xlRight
        c.NumberFormat = "_(#,##0.0_);_((#,##0.0);_(""-""??_);_(@_)"
    Next c

End Sub