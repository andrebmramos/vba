Public Function XLOOKUP(text As Variant, targetList As Range, resultList As Range)
' Purpose: Custom XLOOKUP

    XLOOKUP = WorksheetFunction.Index(resultList, WorksheetFunction.Match(text, targetList, 0))

End Function

' ************************
' To-Do
' ************************
' Option to return a user-defined result if matched.
' =XLOOKUP(lookup_value, reference_list, "TRUE")
'
' This helps do away with IF statement
'
