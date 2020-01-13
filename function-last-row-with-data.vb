Function fLastRowData( _
    ByVal pWorkbookName As String, _
    ByVal pSheetName As String, _
    ByVal pCol As String _
    )

    Dim vR As Long
    Dim vLr As Long
    vR = 0
    vLr = 1
    Do
        vR = vLr
            vLr = Workbooks(pWorkbookName).Worksheets(pSheetName).Range(pCol & vR).End(xlDown).Row
    Loop While vLr <> vR
    vLr = Workbooks(pWorkbookName).Worksheets(pSheetName).Range(pCol & vR).End(xlUp).Row
    
    fLastRowData = vLr
    
End Function
