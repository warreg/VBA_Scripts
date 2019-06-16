Sub Lancement_Boite()
' Affiche la boite de dialogue
    UserForm1.Show
    
' Masque la boite et efface les données de la mémoire
    Unload UserForm1
End Sub


Private Sub ListBox1_Change()
Dim Val_Indice As Byte
    Val_Indice = Me.ListBox1.ListIndex
    Me.Label1.Caption = "Valeur sélectionnée : " & _
       Me.ListBox1.List(Val_Indice)
End Sub

' Cas d'une multiselection:
' on fait une boucle sur tous les éléments de la liste

Public Val_Selection As String

Private Sub CommandButton1_Click()
    MsgBox Val_Selection
End Sub

Private Sub ListBox1_Change()
Dim i As Byte
    Val_Selection = ""
    With ListBox1
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Val_Selection = Val_Selection & Me.ListBox1.List(i) & Chr(13)
            End If
        Next
    End With
End Sub


' ALIMENTER ET SUPPRIMER DES ELEMENTS DE LA LISTE BOX 

Public Val_Selection As String

Private Sub CommandButton1_Click()
    Me.ListBox1.RemoveItem (Me.ListBox1.ListIndex)
End Sub

Private Sub ListBox1_Change()
Dim i As Byte
    Val_Selection = ""
    With ListBox1
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Val_Selection = Val_Selection & Me.ListBox1.List(i) & Chr(13)
            End If
        Next
    End With
End Sub


Private Sub UserForm_Initialize()
Dim i As Byte
    Me.ListBox1.RowSource = ""
    For i = 0 To 5
        Me.ListBox1.AddItem (i * 2)
    Next
End Sub

