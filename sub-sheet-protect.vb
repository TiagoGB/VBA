Sub sProtect(Byval pWorkbookName As String, ByVal pSheetName As String)

    'gPass - definite global constant

    With Workbooks(pWorkbookName).Worksheets(pSheetName)
        .Protect Password:=gPass, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True
    End With

End Sub