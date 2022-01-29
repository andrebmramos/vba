Sub WorkbookFormulaToValue()
' Purpose: Convert all workbook formulas to values (most efficient way)
' Alternative but slower solution:
'    Sub AllValues()
'    Dim wSh As Worksheet
'     For Each wSh In ActiveWorkbook.Worksheets
'     With wSh.UsedRange
'     .Copy
'     .PasteSpecial xlPasteValues
'     End With
'     Next wSh
'
'     Application.CutCopyMode = False
'    End Sub

    Dim sh As Worksheet, HidShts As New Collection
    For Each sh In ActiveWorkbook.Worksheets
    If Not sh.Visible Then
    HidShts.Add sh
    sh.Visible = xlSheetVisible
    End If
    Next sh
     
    Worksheets.Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Select
    Application.CutCopyMode = False
     
    For Each sh In HidShts
    ' sh.Delete
    sh.Visible = xlSheetHidden
    Next sh
   
End Sub