Private Sub txtData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii.Value)) Or Len(txtData.Text) >= 10 Then
        KeyAscii.Value = 0
    Else
        If Len(txtData.Text) = 2 Or Len(txtData.Text) = 5 Then
            txtData.Text = txtData.Text & "/"
        End If
    End If
End Sub