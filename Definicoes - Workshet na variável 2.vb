Dim vWB As String
Dim vAba1 As String
Dim vAba2 As String

Dim wb As Workbook 'Workbook novo arquivo
Dim ws1 As Worksheet 'Worksheet geral
Dim ws2 As Worksheet 'Worksheet rating

vWB = ThisWorkbook.Name
vAba1 = "GERAL"
vAba2 = "Planilha1"

Set wb = Workbooks(vWB)
Set ws1 = wb.Worksheets(vAba1)
Set ws2 = wb.Worksheets(vAba2)
