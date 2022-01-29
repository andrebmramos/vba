Sub FormatColorRed()
' Purpose: To highlight range for follow-up

    Dim rng As Range
    Set rng = Selection
    
    rng.Interior.Color = RGB(255, 204, 204)
    
End Sub