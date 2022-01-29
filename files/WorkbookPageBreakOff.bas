Sub WorkbookPageBreakOff()
'   Purpose: This removes all page breaks for all worksheets in the workbook
'   Source: www.DedicatedExcel.com
 
    Dim ws As Worksheet
 
    For Each ws In Sheets
        ws.DisplayPageBreaks = False
    Next ws
 
     For Each ws In Sheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    Next ws
          
End Sub