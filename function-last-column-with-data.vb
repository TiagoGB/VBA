Function fLastColData( _
    ByVal pWorkbookName As String, _
    ByVal pSheetName As String, _
    ByVal pRow As Long, _
    Optional ByVal pReturnType As String _
    )

    Dim vC As Long
    Dim vLc As Long
    Dim vCChr As String
    Dim vAdd As String
    vC = 1
    vLc = 1
        
    With Workbooks(pWorkbookName).Worksheets(pSheetName)
        Do
            vC = vLc
            vCChr = fAddressToColumnLetter(.Cells(pRow, vC).Address)
            vLc = .Range(vCChr & pRow).End(xlToRight).Column
        Loop While vLc <> vC
        
            vCChr = fAddressToColumnLetter(.Cells(pRow, vC).Address)
            vLc = .Range(vCChr & pRow).End(xlToLeft).Column
        
    If pReturnType = "LETRA" Then
        fLastColData = fAddressToColumnLetter(.Cells(pRow, vLc).Address)
        Exit Function
    End If
    
    End With
    
    fLastColData = vLc
    
End Function	