Dim vWB As String
Dim vPlan1 As String
Dim vPlan2 As String

Dim wb As Workbook 'Workbook
Dim ws1 As Worksheet 'Worksheet 1
Dim ws2 As Worksheet 'Worksheet 2

vWB = ThisWorkbook.Name
vPlan1 = "Plan1"
vPlan2 = "Plan2"

Set wb = Workbooks(vWB)
Set ws1 = wb.Worksheets(vPlan1)
Set ws2 = wb.Worksheets(vPlan2)