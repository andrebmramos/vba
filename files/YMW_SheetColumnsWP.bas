Sub YMW_SheetColumnsWP()
' Purpose: Standardise workbook columns width

    Dim ws As Worksheet
    For Each ws In Worksheets
        Columns.COLUMNWIDTH = 14
        Columns("A").COLUMNWIDTH = 1
        Columns("B").COLUMNWIDTH = 3
        Columns("C").COLUMNWIDTH = 5
    Next ws
    
End Sub