Sub FontArial()
' Purpose: Set selected range to Arial

    Dim rng As Range
    Set rng = Selection
    rng.Font.Name = "Arial"
    rng.Font.Size = 10

End Sub