Sub FormatDateFull()
'   Purpose: Set date format on selected range

    Dim rngSelection As Range
    Set rngSelection = Selection

    For Each c In rngSelection
'       If Not c.Value = vbNullString Then
            c.WrapText = False
            c.HorizontalAlignment = xlCenter
            c.NumberFormat = "DD MMMM YYYY"
'       End If
    Next c
    
End Sub