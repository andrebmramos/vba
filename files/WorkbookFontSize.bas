Sub WorkbookFontSize()
' Purpose: Standardise workbook font size

    Dim ws As Worksheet
    For Each ws In Worksheets
         With ws
            .Cells.Font.Size = 10
         End With
    Next ws
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Activate
        ActiveWindow.Zoom = 100
    Next
    Application.ScreenUpdating = True
    
End Sub