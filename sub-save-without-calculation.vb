Sub sSaveWithoutCalculation(ByVal pWorkbookName As String)

    Dim vCalculation As Double
    
    With Application
    
        vCalculation = .Calculation
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        .ScreenUpdating = False
        .CalculateBeforeSave = False
        
        Workbooks(pWorkbookName).Save
        
        .CalculateBeforeSave = True
        .DisplayAlerts = True
        .Calculation = vCalculation
        .ScreenUpdating = True
        
    End With

End Sub