Private Sub txtCEP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii.Value)) Or Len(txtCEP.Text) >= 10 Then
        KeyAscii.Value = 0
    Else
        If Len(txtCEP.Text) = 2 Then
            txtCEP.Text = txtCEP.Text & "."
        End If
        If Len(txtCEP.Text) = 6 Then
            txtCEP.Text = txtCEP.Text & "-"
        End If
    End If
End Sub