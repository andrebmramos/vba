Sub RevertFile()
'   Purpose: Revert macro changes
'   Reference: https://www.excelforum.com/excel-programming-vba-macros/491103-undoing-a-macro.html

    wkname = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    ActiveWorkbook.Close Savechanges:=False
    
    Workbooks.Open FileName:=wkname

End Sub