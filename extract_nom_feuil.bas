Sub Navigation()
Dim Ma_Feuille As Object
Dim Num_Ligne As Long

    Num_Ligne = 1
    For Each Ma_Feuille In Worksheets
        If Ma_Feuille.Name <> "Sommaire" Then
        ' Index dans le sommaire
            Sheets("Sommaire").Cells(Num_Ligne, 1) = Ma_Feuille.Name
            Sheets("Sommaire").Hyperlinks.Add Anchor:=Sheets("Sommaire").Cells(Num_Ligne, 1), Address:="", SubAddress:=Ma_Feuille.Name & "!A1", TextToDisplay:=Ma_Feuille.Name
            
        ' Retour vers le sommaire dans chaque feuille
            Sheets(Ma_Feuille.Name).Hyperlinks.Add Anchor:=Sheets(Ma_Feuille.Name).Cells(1, 4), Address:="", SubAddress:="Sommaire!A1", TextToDisplay:="Retour"
            
            Num_Ligne = Num_Ligne + 1
        End If
    Next
End Sub
