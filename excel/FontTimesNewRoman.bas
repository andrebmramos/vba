Sub FontTimesNewRoman()
' Purpose: Set selected range to Times New Roman

    Dim rng As Range
    Set rng = Selection
    rng.Font.Name = "Times New Roman"
    rng.Font.Size = 10

End Sub