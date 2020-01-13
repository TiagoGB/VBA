Private Sub TB_CNPJ_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii.Value)) Or Len(TB_CNPJ.Text) >= 18 Then
        KeyAscii.Value = 0
    Else
        If Len(TB_CNPJ.Text) = 2 Or Len(TB_CNPJ.Text) = 6 Then
            TB_CNPJ.Text = TB_CNPJ.Text & "."
        End If
        If Len(TB_CNPJ.Text) = 10 Then
            TB_CNPJ.Text = TB_CNPJ.Text & "/"
        End If
        If Len(TB_CNPJ.Text) = 15 Then
            TB_CNPJ.Text = TB_CNPJ.Text & "-"
        End If
    End If
End Sub