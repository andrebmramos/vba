Sub SortRight()
'   Purpose: Sort a series of numbers from left to right

    Dim row As Range
    
    For Each row In Selection.Rows
        row.Sort Key1:=row, Order1:=xlAscending, Orientation:=xlSortRows
    Next row

'   row.Sort Key1:=row, Order1:=xlAscending, Orientation:=xlSortRows
'   row.Sort Key1:=row, Order1:=xlAscending, Orientation:=xlLeftToRight

End Sub