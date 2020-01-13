Sub sDisableOptimization()
    
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        
    End With

End Sub