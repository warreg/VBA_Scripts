Sub Programme_Principal()
Dim Chemin As String
Dim Fichier As String
Dim i_Lig As Long

Dim CP As String * 5
Dim Ville As String * 30
Dim Ref As String * 9
Dim Nom As String * 25
Dim Prenom As String * 15
    
    Chemin = "C:\Votre_Chemin"
    Open Chemin & "\Fichier_Sortie.txt" For Output As #1
    i_Lig = 2
    
    While Cells(i_Lig, 1) <> ""
        CP = Cells(i_Lig, 1)
        Ville = Cells(i_Lig, 2)
        Ref = Cells(i_Lig, 3)
        Nom = Cells(i_Lig, 4)
        Prenom = Cells(i_Lig, 5)
        Fichier = CP & Ville & Ref & Nom & Prenom
        Print #1, Fichier
        i_Lig = i_Lig + 1
    Wend
    
    Close #1
End Sub
