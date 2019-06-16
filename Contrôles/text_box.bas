Sub Lancement_Boite()
' Affiche la boite de dialogue
    UserForm1.Show
    
' Masque la boite et efface les données de la mémoire
    Unload UserForm1
End Sub

' Userform
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.TextBox1.Text <> "" Then
        If Not IsNumeric(Me.TextBox1.Text) Then
            MsgBox "Vous devez saisir une valeur numérique"
        Else
            If Me.TextBox1.Text <= 0 Or Me.TextBox1.Text > 10 Then
                MsgBox "Valeur erronée"
            End If
        End If
    End If
End Sub
