Sub RemoveBlankRows()
' Purpose: Remove empty rows
' Source: https://www.ablebits.com/office-addins-blog/2018/12/19/delete-blank-lines-excel/

    Dim SourceRange As Range
    Dim EntireRow As Range
 
    Set SourceRange = Application.Selection
 
    If Not (SourceRange Is Nothing) Then
        Application.ScreenUpdating = False
 
        For i = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
                EntireRow.Delete
            End If
        Next
 
        Application.ScreenUpdating = True
    End If
    
End Sub