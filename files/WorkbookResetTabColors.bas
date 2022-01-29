Sub WorkbookResetTabColors()
'   Purpose: Reset all tab colors
'   Source: https://www.extendoffice.com/documents/excel/5179-excel-remove-tab-color.html
    Dim xSheet As Worksheet
    
    For Each xSheet In ActiveWorkbook.Worksheets
        xSheet.Tab.ColorIndex = xlColorIndexNone
    Next xSheet
    
End Sub