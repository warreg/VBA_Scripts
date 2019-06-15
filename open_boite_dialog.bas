Sub Ouvrir_Boite_Dialogue()
   Application.FileDialog(msoFileDialogFolderPicker).Show
End Sub

' Sélectionner un répertoire et afficher son nom dans un message box
Sub Rep_Selection()
Dim Rep As String
    Application.FileDialog(msoFileDialogFolderPicker).Show
    Rep = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    MsgBox Rep
End Sub


' Sélectionner un seul fichier
Sub Selection_Un_Fichier()
    Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogFilePicker).Show
    MsgBox Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
End Sub

' Sélectionner plusieurs fichiers
Sub Selection_X_Fichiers()
Dim Tab_Nom_Fichier() As String
Dim Index_Fichier As Long
    Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = True
    Application.FileDialog(msoFileDialogFilePicker).Show
    For Index_Fichier = 1 To Application.FileDialog(msoFileDialogFilePicker).SelectedItems.Count
        ReDim Preserve Tab_Nom_Fichier(Index_Fichier - 1)
        Tab_Nom_Fichier(Index_Fichier - 1) = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(Index_Fichier)
    Next
End Sub


