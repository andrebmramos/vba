Sub HyperAdd()
'   Purpose: Converts each text hyperlink selected into a working hyperlink
'   Source: https://superuser.com/questions/580387/how-to-turn-plain-text-links-into-hyperlinks-in-excel

    Dim xCell As Range
        
    For Each xCell In Selection
        ActiveSheet.Hyperlinks.Add Anchor:=xCell, Address:=xCell.Formula
    Next xCell
    
End Sub