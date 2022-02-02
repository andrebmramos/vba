Sub FormatAccounting()
'   Purpose: Set accounting number format on selected range

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlRight
            c.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
'       End If
    Next c
    
End Sub
