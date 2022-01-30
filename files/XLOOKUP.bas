Public Function XLOOKUP(text As Variant, targetList As Range, resultList As Variant, Optional errResult As Variant)
'   Purpose: Custom XLOOKUP
'   Usage 01: =XLOOKUP(A1, A1:A10, B1:B10)
'   Usage 02: =XLOOKUP(A1, A1:A10, "True", "False")
'   Ref: https://stackoverflow.com/questions/44638867/vba-excel-try-catch
'   Ref: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
'   Ref: https://stackoverflow.com/questions/32008841/best-way-to-return-error-in-udf-vba-function
'   Ref: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement
'   Version: 20220130.2350

    On Error GoTo XLOOKUP_Error
    Application.ScreenUpdating = False
    
    If TypeName(resultList) = "Range" Then
        XLOOKUP = WorksheetFunction.Index(resultList, WorksheetFunction.Match(text, targetList, 0))
    Else
        If IsError(WorksheetFunction.Match(text, targetList, 0)) Then
            GoTo XLOOKUP_Error
        Else
            XLOOKUP = resultList
        End If
    End If
    
    Application.ScreenUpdating = True
    Exit Function
    
XLOOKUP_Error:

    If IsMissing(errResult) Then
        XLOOKUP = CVErr(xlErrValue)
    Else
        XLOOKUP = errResult
    End If
    Resume Next
    
End Function
