Sub Procedure_Principale()
Dim Rep_Travail As String
    Rep_Travail = ActiveWorkbook.Path
    Call Ouverture_Fichier(Rep_Travail)
    Call Tri_Données
End Sub

Sub Ouverture_Fichier(Chemin_Acces As String)
    Workbooks.Open Filename:=Chemin_Acces & "Client.xlsx"
End Sub

Sub Tri_Données()
    Range("H1").Select
    Selection.CurrentRegion.Select
    ActiveWorkbook.Worksheets("Feuil1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Feuil1").Sort.SortFields.Add Key:=Range("H2:H201") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuil1").Sort
        .SetRange Range("A1:J201")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


' AMÉLIORATION: on passe en paramètre la colonne à trier lors de l'appel de la procédure Tri_Donnees
' on oublie pas de creer une liste déroulante en cellule A1 et on remplace partout 
' dans le code initial la référence de la colonne de tri

Sub Procedure_Principale()
Dim Rep_Travail As String
Dim Colonne_Lettre As String
    
' Variables de travail
    Rep_Travail = ActiveWorkbook.Path & "/"
    Colonne_Lettre = Cells(1, 1)
    
    Call Ouverture_Fichier(Rep_Travail)
    
    Call Tri_Données(Colonne_Lettre)
End Sub

Sub Ouverture_Fichier(Chemin_Acces As String)
    Workbooks.Open Filename:=Chemin_Acces & "Client.xlsx"
End Sub

Sub Tri_Données(Lettre_Tri As String)
    Range(Lettre_Tri & "1").Select
    Selection.CurrentRegion.Select
    ActiveWorkbook.Worksheets("Feuil1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Feuil1").Sort.SortFields.Add Key:=Range(Lettre_Tri & "2:" & Lettre_Tri & "201") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuil1").Sort
        .SetRange Range("A1:J201")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


