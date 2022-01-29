Sub WorkbookResetStyles()
'   The purpose of this macro is to remove all styles in the active
'   workbook and rebuild the default styles.
'   It rebuilds the default styles by merging them from a new workbook.

'   Dimension variables.
    Dim MyBook As Workbook
    Dim tempBook As Workbook
    Dim CurStyle As Style

'   Set MyBook to the active workbook.
    Set MyBook = ActiveWorkbook
    On Error Resume Next
   
'   Delete all the styles in the workbook.
    For Each CurStyle In MyBook.Styles
   
'   If CurStyle.Name <> "Normal" Then CurStyle.Delete
    Select Case CurStyle.Name
        Case "20% - Accent1", "20% - Accent2", _
            "20% - Accent3", "20% - Accent4", "20% - Accent5", "20% - Accent6", _
            "40% - Accent1", "40% - Accent2", "40% - Accent3", "40% - Accent4", _
            "40% - Accent5", "40% - Accent6", "60% - Accent1", "60% - Accent2", _
            "60% - Accent3", "60% - Accent4", "60% - Accent5", "60% - Accent6", _
            "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", _
            "Bad", "Calculation", "Check Cell", "Comma", "Comma [0]", "Currency", _
            "Currency [0]", "Explanatory Text", "Good", "Heading 1", "Heading 2", _
            "Heading 3", "Heading 4", "Input", "Linked Cell", "Neutral", "Normal", _
            "Note", "Output", "Percent", "Title", "Total", "Warning Text"
               
'   Do nothing, these are the default styles
        Case Else
            CurStyle.Delete
    End Select

   Next CurStyle

'   ==============================================
'   Alternative approach
'   Purpose: Delete unwanted styles
'   Source: https://excel.tips.net/T002135_Deleting_Unwanted_Styles.html

    'Dim styT As Style
    'Dim intRet As Integer
    'On Error Resume Next
    'For Each styT In ActiveWorkbook.Styles
    '    If Not styT.BuiltIn Then
    '        If styT.Name <> "1" Then styT.Delete
    '    End If
    'Next styT
     
'   ==============================================

End Sub