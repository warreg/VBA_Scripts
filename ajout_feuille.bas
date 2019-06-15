Sub Ajout_Feuille()
Dim Num_Ligne As Long
    Num_Ligne = 2
    While Sheets("Feuil1").Cells(Num_Ligne, 1) <> ""
        Sheets.Add After:=Sheets(Sheets.Count)
        
        ' En utilisant l'opérateur &, nous allons consevoir une variable qui va contenir l'association du nom et du prénom
        ' De plus, entre ces 2 cellules, nous ajoutons aussi un espace pour qu'au résultat final, les valeurs des 2 cellules ne soient
        ' pas collées l'une à l'autre.
        ActiveSheet.Name = Sheets("Feuil1").Cells(Num_Ligne, 1) & " " & Sheets("Feuil1").Cells(Num_Ligne, 2)
        Num_Ligne = Num_Ligne + 1
    Wend
End Sub
