Public Function XLOOKUP(text As Variant, targetList As Range, resultList As Variant, Optional errResult As Variant)
'   Purpose: Custom XLOOKUP
'   Usage 01: =XLOOKUP(A1, A1:A10, B1:B10)
'   Usage 02: =XLOOKUP(A1, A1:A10, "True", "False")
'   Ref: https://stackoverflow.com/questions/44638867/vba-excel-try-catch
'   Ref: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
'   Ref: https://stackoverflow.com/questions/32008841/best-way-to-return-error-in-udf-vba-function

    Application.ScreenUpdating = False
    
    On Error GoTo XLOOKUP_Error
    
    If (VarType(resultList) = 8204) Then
            XLOOKUP = WorksheetFunction.Index(resultList, WorksheetFunction.Match(text, targetList, 0))
    Else
        If (WorksheetFunction.Match(text, targetList, 0)) Then
            XLOOKUP = resultList
        Else: End If
    End If
    
XLOOKUP_Error:
'   Handles error if not match found
    If IsMissing(errResult) Then
        XLOOKUP = CVErr(xlErrValue)
    Else
        XLOOKUP = errResult
    End If
    
    Application.ScreenUpdating = True
    
End Function
