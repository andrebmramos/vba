Sub YMW_InsertHeadingTax()
'   Purpose: Insert customised headings for tax workpapers
'   Note: Utilises CCH Engagement functions

    Dim rng As Range
    Dim myClient As String
    Dim myYear As String
    Set rng = Selection
    
    myClient = "=UPPER(PJNAME())"

    myYear = "=" & Chr(34) & "YEAR OF ASSESSMENT " & Chr(34) & "&"
    myYear = myYear & "UPPER(TEXT(" & "CYBDATE()+1" & "," & Chr(34) & "yyyy" & Chr(34)
    myYear = myYear & "))"

    If rng.HasFormula = True Then
        rng.Formula = Replace(rng.Formula, rng.Formula, myClient)
    Else
        rng.Formula = "=1"
        rng.Formula = Replace(rng.Formula, rng.Formula, myClient)
    End If
    
    If rng.Offset(1, 0).HasFormula = True Then
        rng.Offset(1, 0).Formula = Replace(rng.Offset(1, 0).Formula, rng.Offset(1, 0).Formula, myYear)
    Else
        rng.Offset(1, 0).Formula = "=1"
        rng.Offset(1, 0).Formula = Replace(rng.Offset(1, 0).Formula, rng.Offset(1, 0).Formula, myYear)
    End If
    
    rng.Copy
    rng.PasteSpecial Paste:=xlPasteValues
    rng.Offset(1, 0).Copy
    rng.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
    
    rng.Font.Bold = True
    rng.Offset(1, 0).Font.Bold = True
    
End Sub