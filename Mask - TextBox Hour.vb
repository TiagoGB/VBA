Private Sub txtHora_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii.Value)) Or Len(txtHora.Text) >= 5 Then
        KeyAscii.Value = 0
    Else
        If Len(txtHora.Text) = 2 Then
            txtHora.Text = txtHora.Text & ":"
        End If
    End If
End Sub