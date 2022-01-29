Sub FormatTextToValue()
'   Purpose: Convert text format to number format on selected range

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlRight
            c.NumberFormat = "General"
            c.Value = c.Value
'       End If
    Next c

End Sub