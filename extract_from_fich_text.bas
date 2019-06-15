Public t_col() As Long

Sub Lire_Fichier()
Dim Chaine_Lue As String
Dim Est_Entete As Boolean 'Variable utilisée pour ne récupérer la structure du tableau qu'une seule fois

    Open "C:\MesDocuments\Exemple 11-B5.txt" For Input As #1
    Est_Entete = True
    While Not EOF(1)
        Line Input #1, Chaine_Lue
        If Est_Entete Then
            Call Structure_Fichier(Chaine_Lue)
            Est_Entete = False
        End If
        Call Separation_Colonne(Chaine_Lue)
    Wend
    
    Close #1
    
End Sub

' ------------------------------------------------------------------------
' Procédure de détection de la position des colonnes
' ------------------------------------------------------------------------
Sub Structure_Fichier(Texte As String)
Dim i_car As Long           'Indice pour les caractères
Dim i_tab As Long           'Indice pour le tableau
    
    ReDim t_col(0)          'Redimensionnement du tableau
    t_col(0) = 1            'Premiere valeur du tableau obligatoirement 1
    i_tab = 1
    
    For i_car = 1 To Len(Texte) 'Boucle jusqu'à la fin de la ligne
        If Mid(Texte, i_car, 1) = " " Then  'Un espace a été atteint
            
            ' Tant que nous avons des espaces, nous lisons le caractère suivant
            While Mid(Texte, i_car, 1) = " "
                i_car = i_car + 1
            Wend
            
            ' En sortie de boucle, nous sommes positionné sur la colonne suivante
            ReDim Preserve t_col(i_tab) ' Redimensionnement du tableau en conservant les données
            t_col(i_tab) = i_car        ' Enregistrement de la valeur du colonnage dans le tableau
            i_tab = i_tab + 1
        End If
    Next
    
    ReDim Preserve t_col(i_tab)
    t_col(i_tab) = i_car
End Sub

' -----------------------------------------------------------------------------------
' Procédure de remplissage de la feuille Excel en utilisant la structure des colonnes
' -----------------------------------------------------------------------------------
Sub Separation_Colonne(Texte As String)
Dim i_tab As Long
Dim Derniere_Ligne  As Long
    
    ' L'instruction Rows.Count permet de retourner le nombre maximum de lignes de votre feuille de calcul
    Derniere_Ligne = Cells(Rows.Count, 1).End(xlUp).Row
    'Test pour gérer le premier fichier (pas de ligne vide)
    Derniere_Ligne = IIf(Cells(1, 1) = "", 1, Derniere_Ligne + 1)
    i_tab = 0
    Do
        'Transfert des valeurs du fichier texte dans chaque colonne
        ActiveSheet.Cells(Derniere_Ligne, i_tab + 1) = _
         Trim(Mid(Texte, t_col(i_tab), t_col(i_tab + 1) - t_col(i_tab)))
        
        i_tab = i_tab + 1
    Loop While i_tab < UBound(t_col)    ' Boucle tant que le dernier indice du tableau
                                        ' n'a pas été atteint
End Sub




