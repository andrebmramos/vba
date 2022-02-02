Sub CaseUpper()
' Purpose: Set upper case on selection

    Dim rng As Range
    Set rng = Selection
    For Each Cell In rng
        Cell.Value = UCase(Cell)
    Next Cell
    
End Sub