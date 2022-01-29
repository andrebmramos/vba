Public Function XLOOKUP(text As Variant, targetList As Range, resultList As Range)
' Purpose: Custom XLOOKUP

    XLOOKUP = WorksheetFunction.Index(resultList, WorksheetFunction.Match(text, targetList, 0))

End Function