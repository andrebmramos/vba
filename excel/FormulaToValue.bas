Sub FormulaToValue()
' Purpose: Convert selected formulas to values

    Dim rng As Range
    Set rng = Selection
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    
End Sub