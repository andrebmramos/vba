Sub CaseProper()
' Purpose: Set upper case on selection

    Dim rng As Range
    Set rng = Selection
    For Each Cell In rng
        Cell.Value = StrConv(Cell, vbProperCase)
    Next Cell
    
End Sub
