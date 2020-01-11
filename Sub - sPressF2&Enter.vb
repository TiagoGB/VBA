Sub sPressF2Enter()

    Call fDisableOptimization

    Dim L As Long

    L = 1
    Cells(L, 8).Select
    
    For L = 1 To 412
    
    SendKeys "{F2}"
    SendKeys "{ENTER}"

    Next L
    
    Call fEnableOptimization
    
End Sub