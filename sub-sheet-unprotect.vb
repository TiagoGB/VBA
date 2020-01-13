Sub sUnprotect(ByVal pWorkbookName As String, ByVal pSheet As String)

    'gPass - definite global constant

    Workbooks(pWorkbookName).Worksheets(pSheet).Unprotect gPass

End Sub