Sub sPressF2Enter()

    '----------------------------------------------------------------------------------------------------'
    '-- START - DECLARATION OF VARIABLES
    '----------------------------------------------------------------------------------------------------'    

        Dim vWB As String       ' Workbook Name
        Dim vSheet As String    ' Sheet Name
        Dim vWS As Worksheet    ' Active Sheet
        Dim vRIni As Long       ' Initial Row
        Dim vREnd As Long       ' Final Row
        Dim vR As Long          ' Row
        Dim vCIni As Long       ' Initial Column
        Dim vCnd As Long        ' Final Column
        Dim vC As Long          ' Column

    '----------------------------------------------------------------------------------------------------'
    '-- END - DECLARATION OF VARIABLES
    '----------------------------------------------------------------------------------------------------'    
    
    '----------------------------------------------------------------------------------------------------'
    '-- START - INITIAL DEFINITIONS
    '----------------------------------------------------------------------------------------------------'    

        vWB = "Excel"
        vSheet = "Plan1"

        Set vWS = Workbooks(vWB).Worksheets(vSheet)

        vRIni = 1       ' Initial Row
        vREnd = 200     ' Final Row

        vRIni = 1       ' Initial Column
        vREnd = 200     ' Final Column

    '----------------------------------------------------------------------------------------------------'
    '-- END - INITIAL DEFINITIONS
    '----------------------------------------------------------------------------------------------------'    

    '----------------------------------------------------------------------------------------------------'
    '-- START - CODE
    '----------------------------------------------------------------------------------------------------'    

        vR = vRIni
        vC = vCIni
        vWS.Cells(vR, vC).Select
        
        'For Rows

        For vR = vRIni To vREnd
        
            SendKeys "{F2}"
            SendKeys "{ENTER}"

        Next vR

        'For Columns

        For vC = vCIni To vCEnd
        
            SendKeys "{F2}"
            SendKeys "{ENTER}"

        Next vC

    '----------------------------------------------------------------------------------------------------'
    '-- END - CODE
    '----------------------------------------------------------------------------------------------------'    
    
End Sub