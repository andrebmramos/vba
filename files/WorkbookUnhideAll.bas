Sub WorkbookUnhideAll()
'   Purpose: Unhide all rows and columns
'   Source: https://www.extendoffice.com/documents/excel/1173-excel-break-all-links.html

    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Columns.EntireColumn.Hidden = False
        ws.Rows.EntireRow.Hidden = False
    Next ws
    
End Sub