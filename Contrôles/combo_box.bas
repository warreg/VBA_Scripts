Sub Lancement_Boite()
' Affiche la boite de dialogue
    UserForm1.Show
    
' Masque la boite et efface les données de la mémoire
    Unload UserForm1
End Sub


Private Sub CommandButton1_Click()
Dim Derniere_Ligne As Long
    Derniere_Ligne = Cells(1, 1).CurrentRegion.Rows.Count + 1
    Cells(Derniere_Ligne, 1) = Me.ComboBox1.Text
    
    Call UserForm_Initialize
    
End Sub

Private Sub UserForm_Initialize()
Dim i_Ligne As Long
    Me.ComboBox1.RowSource = ""
    i_Ligne = 1
    While Cells(i_Ligne, 1) <> ""
        Me.ComboBox1.AddItem (Cells(i_Ligne, 1))
        i_Ligne = i_Ligne + 1
    Wend
End Sub
