Sub FormatLineDotted()
' Purpose: Insert dotted line

    Dim rng As Range
    Set rng = Selection
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    
End Sub