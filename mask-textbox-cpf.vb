Private Sub txtCPF_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii.Value)) Or Len(txtCPF.Text) >= 14 Then
        KeyAscii.Value = 0
    Else
        If Len(txtCPF.Text) = 3 Or Len(txtCPF.Text) = 7 Then
            txtCPF.Text = txtCPF.Text & "."
        End If
        If Len(txtCPF.Text) = 11 Then
            txtCPF.Text = txtCPF.Text & "-"
        End If
    End If
End Sub