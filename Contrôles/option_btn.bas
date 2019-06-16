Sub Lancement_Boite()
' Affiche la boite de dialogue
    UserForm1.Show
    
' Masque la boite et efface les données de la mémoire
    Unload UserForm1
End Sub


Private Sub CommandButton1_Click()
Select Case True
    Case Me.OptionButton1.Value
        MsgBox Me.OptionButton1.Caption
    Case Me.OptionButton2.Value
        MsgBox Me.OptionButton2.Caption
    Case Me.OptionButton3.Value
        MsgBox Me.OptionButton3.Caption
    Case Me.OptionButton4.Value
        MsgBox Me.OptionButton4.Caption
    Case Me.OptionButton5.Value
        MsgBox Me.OptionButton5.Caption
    End Select
End Sub
