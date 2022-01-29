Sub FormulaAbsoluteReference()
' Purpose: Absolute reference selected cells
' Source: http://www.excelforum.com/excel-general/372383-making-multiple-cells-absolute-at-once.html
' Source: http://www.contextures.com/xlvba01.html#videoreg

    Dim Cell As Range
    
    For Each Cell In Selection
        If Cell.HasFormula Then
            Cell.Formula = _
            Application.ConvertFormula(Cell.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next
 
End Sub