Sub Mettre_meme_lign()
'ORGANISER SON TABLEAU EN PRESENTANT SUR UNE MEME LIGNE, UNE DATE ET TOUS LES PRENOMS CORRESPONDANTS A CETTE DATE
  
  'li: lign; li_w: ligne write (ligne d'écriture ); col_w: col write
  Dim li&, li_w&, col_w&
  li = 1
  li_w = 1
  While Cells(li, 1) <> ""
    col_w = 5
    Do
      Cells(li_w, col_w) = Cells(li, 2)
      col_w = col_w + 1
      li = li + 1
    Loop While Cells(li - 1, 1) = Cells(li, 1)
    Cells(li_w, 4) = Cells(li - 1, 1)
    li_w = li_w + 1
  Wend
  
End Sub
    

' AMELIORATION: Pour eviter de reécrire la date à chaque fois que le programme traite une ligne.
' On peut alors écrire la date une seule fois par ligne d'écriture en faisant un test 
' pour éviter de réécrire cette donnée si elle est déja présente
' on peut alors soit : - tester que la cellule de la colone 5 est vide 
'                      - tester que la variable Col_Ecriture = 5

Sub Amelioration_1()
Dim Num_Ligne As Long
Dim Col_Ecriture As Long
Dim Lig_Ecriture As Long
    
    Lig_Ecriture = 1
    Num_Ligne = 1
    While Cells(Num_Ligne, 1) <> ""
        Col_Ecriture = 5
        Do
            If Col_Ecriture = 5 Then
                Cells(Lig_Ecriture, 4) = Cells(Num_Ligne, 1)
            End If
            Cells(Lig_Ecriture, Col_Ecriture) = Cells(Num_Ligne, 2)
            Col_Ecriture = Col_Ecriture + 1
            Num_Ligne = Num_Ligne + 1
        Loop While Cells(Num_Ligne - 1, 1) = Cells(Num_Ligne, 1)
        Lig_Ecriture = Lig_Ecriture + 1
    Wend
End Sub

    
    
' On peut aussi écrire la date en sortie de seconde boucle 
' on peut éviter ainsi de réaliser un test. il faut retourner la valeur de la date de la ligne précedente 
' Date écrite à la fin
Sub Amelioration_2()
Dim Num_Ligne As Long
Dim Col_Ecriture As Long
Dim Lig_Ecriture As Long
    
    Lig_Ecriture = 1
    Num_Ligne = 1
    While Cells(Num_Ligne, 1) <> ""
        Col_Ecriture = 5
        Do
            Cells(Lig_Ecriture, Col_Ecriture) = Cells(Num_Ligne, 2)
            Col_Ecriture = Col_Ecriture + 1
            Num_Ligne = Num_Ligne + 1
        Loop While Cells(Num_Ligne - 1, 1) = Cells(Num_Ligne, 1)
        Cells(Lig_Ecriture, 4) = Cells(Num_Ligne - 1, 1)
        Lig_Ecriture = Lig_Ecriture + 1
    Wend
End Sub
        

        
        
' Pour travailler sur deux feuilles 
' Pour toujours lire les données de la feuille 1 mais le résultat doit être copié dans la feuille 2 en colone A
' Ecriture sur deux feuilles de calcul de calcul différent
Sub Programme_Principal()
Dim Num_Ligne As Long
Dim Col_Ecriture As Long
Dim Lig_Ecriture As Long
    
    Lig_Ecriture = 1
    Num_Ligne = 1
    While Sheets("Feuil1").Cells(Num_Ligne, 1) <> ""
        Col_Ecriture = 2
        Do
            Sheets("Feuil2").Cells(Lig_Ecriture, 1) = Sheets("Feuil1").Cells(Num_Ligne, 1)
            Sheets("Feuil2").Cells(Lig_Ecriture, Col_Ecriture) = Sheets("Feuil1").Cells(Num_Ligne, 2)
            Col_Ecriture = Col_Ecriture + 1
            Num_Ligne = Num_Ligne + 1
        Loop While Sheets("Feuil1").Cells(Num_Ligne - 1, 1) = Sheets("Feuil1").Cells(Num_Ligne, 1)
        Lig_Ecriture = Lig_Ecriture + 1
    Wend
End Sub

  

