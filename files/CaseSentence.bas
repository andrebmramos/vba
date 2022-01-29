Sub CaseSentence()
' Purpose: Set sentence case on selection

    Dim rng As Range
    Dim WorkRng As Range
    On Error Resume Next
'    xTitleId = "KutoolsforExcel"
    Set WorkRng = Application.Selection
'    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    For Each rng In WorkRng
        xValue = rng.Value
        xStart = True
        For i = 1 To VBA.Len(xValue)
            ch = Mid(xValue, i, 1)
            Select Case ch
                Case "."
                xStart = True
                Case "?"
                xStart = True
                Case "a" To "z"
                If xStart Then
                    ch = UCase(ch)
                    xStart = False
                End If
                Case "A" To "Z"
                If xStart Then
                    xStart = False
                Else
                    ch = LCase(ch)
                End If
            End Select
            Mid(xValue, i, 1) = ch
        Next
        rng.Value = xValue
    Next
    
End Sub