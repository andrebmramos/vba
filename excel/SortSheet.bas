Sub SortSheet()
'   Purpose: Sort worksheets alphabetically
    
    Dim i As Integer
    Dim j As Integer
   
    For i = 1 To Sheets.Count
        For j = 1 To Sheets.Count - 1
            If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
                Sheets(j).Move After:=Sheets(j + 1)
            End If
        Next j
    Next i
    
End Sub
