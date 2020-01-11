Function fAddressToColumnLetter(ByVal pAddress As String)

    Dim vPos1 As Integer
    Dim vPos2 As Integer
    
    vPos1 = 2
    vPos2 = InStr(2, pAddress, "$", vbTextCompare) - vPos1
    
    fAddressToColumnLetter = Mid(pAddress, vPos1, vPos2)

End Function