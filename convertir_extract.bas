Public t_col() As Long

Sub Lire_Fichier_Virgule()
Dim Extract As String
Dim Tableau_Extraction As Variant
Dim Chemin As String

    'Changez le chemin ci-dessous pour correspondre à votre ordinateur
    Chemin = "C:\MesDocuments\Fichier Excel\"

    ' Ouverture du fichier
    Open Chemin & "Exemple 11-Billard.txt" For Input As #1
    Line Input #1, Extract    ' Lecture de la première ligne du fichier

    While Not EOF(1)              ' Boucle sur tout le fichier texte jusqu'à la fin du fichier (End Of File)
        If InStr(Extract, "[Billard]") > 0 Then
            ' Séparation de la ligne lue sur le caractère virgule
            ' Le résultat est transmis dans un tableau en mémoire
            Tableau_Extraction = Split(Extract, ",")
            
            Call Ecriture_Excel(Tableau_Extraction)     'Appel de la procédure pour écrire dans la feuille Excel
        End If
        Line Input #1, Extract    ' Lecture du fichier
    Wend

    Close #1                      ' Fermeture du fichier

    Call Traitement_Colonne_C
    
End Sub

'Procédure d'écriture dans la feuille Excel avec comme argument le tableau contenant les tronçons de texte de notre chaîne initiale
Sub Ecriture_Excel(Tableau_Extraction As Variant)
Dim i_tab As Long
Dim Derniere_Ligne  As Long
    
    Derniere_Ligne = Cells(Rows.Count, 1).End(xlUp).Row   ' Astuce pour trouver la dernière ligne non-vide
    Derniere_Ligne = IIf(Cells(1, 1) = "", 1, Derniere_Ligne + 1)
    i_tab = 0
    Do
        ActiveSheet.Cells(Derniere_Ligne, i_tab + 1) = Trim(Tableau_Extraction(i_tab))
        i_tab = i_tab + 1
    Loop While i_tab <= UBound(Tableau_Extraction)    ' Boucle tant que le dernier indice du tableau
                                                      ' n'a pas été atteint
End Sub

Sub Traitement_Colonne_C()
Dim Num_Ligne As Long
Dim i_tab As Long
Dim Tableau_Extraction As Variant
Dim Nom_Ville As String
Dim Max_Dim_Tableau As Long
Dim Test_Fin_Boucle As Boolean  ' Variable qui va nous servir à arrêter de parcourir notre tableau en mémoire.

    Num_Ligne = 1
    While Cells(Num_Ligne, 3) <> ""
        Tableau_Extraction = Split(Cells(Num_Ligne, 3), " ")    ' Séparation du contenu de la cellule sur le caractère "espace"
        
        ' Initialisation des paramètres
        i_tab = 1
        Nom_Ville = ""
        Test_Fin_Boucle = False
        Max_Dim_Tableau = UBound(Tableau_Extraction) 'Détermine la taille maximale de notre tableau
        
        ' Nous arrêtons de boucler quand nous avons récupérer le nom de la ville.
        ' C'est à dire quand la variable Test_Fin_Boucle est égale à Vrai
        While Test_Fin_Boucle = False
            
            ' Traitement du département
            If i_tab = 1 Then
                ActiveSheet.Cells(Num_Ligne, 4) = "'" & Trim(Tableau_Extraction(i_tab))
            Else
            
                ' Test nécessaire pour éviter que nous dépassions la taille du tableau (plantage du programme)
                If Max_Dim_Tableau = 2 Then
                    'Si notre tableau à une taille = 2, nécessairement, la seconde valeur est le nom de la ville
                    Cells(Num_Ligne, 5) = UCase(Trim(Tableau_Extraction(i_tab)))
                    Test_Fin_Boucle = True
                Else
                    'Si notre tableau à une taille supérieure à 2 c'est que le nom de la ville est réparti sur plusieurs cases de notre tableau
                    Do
                        Nom_Ville = Nom_Ville & " " & CStr(Trim(Tableau_Extraction(i_tab))) ' Concaténation de la ville
                        i_tab = i_tab + 1
                    
                    ' Les tests de fin de boucles sont :
                    ' - caractère = (
                    ' - caratère = >
                    ' - caractère = "
                    ' - fin de tableau
                    Loop While Left(Trim(Tableau_Extraction(i_tab)), 1) <> "(" _
                        And Left(Trim(Tableau_Extraction(i_tab)), 1) <> ">" _
                        And Right(Trim(Tableau_Extraction(i_tab)), 1) <> Chr(34) _
                        And i_tab < Max_Dim_Tableau
                    
                    Cells(Num_Ligne, 5) = UCase(Trim(Nom_Ville))
                    Test_Fin_Boucle = True
                End If
            End If
            i_tab = i_tab + 1 'Incrémentation du tableau
        Wend
        Num_Ligne = Num_Ligne + 1 'Incrémentation de la ligne d'écriture
    Wend
End Sub


