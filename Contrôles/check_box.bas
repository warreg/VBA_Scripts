Sub Lancement_Boite()
' Affiche la boite de dialogue
    UserForm1.Show
    
' Masque la boite et efface les données de la mémoire
    Unload UserForm1
End Sub


Private Sub CheckBox1_Click()
    If Me.CheckBox1.Value = True Then
        Me.CommandButton1.Enabled = True
    Else
        Me.CommandButton1.Enabled = False
    End If

End Sub

Private Sub UserForm_Initialize()
    Me.CommandButton1.Enabled = False
    Me.CheckBox1.Value = False
End Sub
