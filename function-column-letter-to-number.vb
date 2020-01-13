Function fColN2L(ByVal pCol As Long) As String'Converte Númeo da coluna para Letra

    Dim vPos1 As Integer
    Dim vPos2 As Integer
    Dim vAddress As String
    
    vAddress = ThisWorkbook.Worksheets(1).Cells(1, pCol).Address
    
    vPos1 = 2
    vPos2 = InStr(2, vAddress, "$", vbTextCompare) - vPos1
    
    fColN2L = Mid(vAddress, vPos1, vPos2)

End Function