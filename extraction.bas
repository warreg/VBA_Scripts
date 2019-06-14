Public Val_Colonne As String
Public Val_Sign As String
Public Val_Filter As Variant

Public Num_Colonne As Long
Public Num_Ligne As Long

Sub Extraction()
Dim Ecrit_Ligne As Long

    'Nettoyer le résultat précédent
    Columns("N:S").ClearContents
    
    'Chargement des paramètres de la feuille
    Val_Colonne = Cells(2, 8)
    Val_Sign = Cells(2, 9)
    Val_Filter = Cells(2, 10)
    
    'Conversion de la lettre de la colonne en chiffre
    Num_Colonne = Columns(Val_Colonne & ":" & Val_Colonne).Column
    
    'Initialisation des parametres
    Num_Ligne = 2
    Ecrit_Ligne = 2
    
    'Debut de la boucle pour toutes les lignes de la colonne A
    While Cells(Num_Ligne, 1) <> ""
        If Test1 Then
            Range(Cells(Num_Ligne, 1), Cells(Num_Ligne, 5)).Copy
            Range("O" & Ecrit_Ligne).Select
            ActiveSheet.Paste
            Ecrit_Ligne = Ecrit_Ligne + 1
        End If
        Num_Ligne = Num_Ligne + 1
    Wend
    
End Sub

Function Test1()
    Select Case Val_Sign
        Case "<"
            Test1 = Cells(Num_Ligne, Num_Colonne) < Val_Filter
        Case "<="
            Test1 = Cells(Num_Ligne, Num_Colonne) <= Val_Filter
        Case "="
            Test1 = Cells(Num_Ligne, Num_Colonne) = Val_Filter
        Case ">="
            Test1 = Cells(Num_Ligne, Num_Colonne) >= Val_Filter
        Case ">"
            Test1 = Cells(Num_Ligne, Num_Colonne) > Val_Filter
        Case "<>"
            Test1 = Cells(Num_Ligne, Num_Colonne) <> Val_Filter
        Case "Contient"
            Test1 = Cells(Num_Ligne, Num_Colonne) Like "*" & Val_Filter & "*"
        Case "Ne contient pas"
            Test1 = Not (Cells(Num_Ligne, Num_Colonne) Like "*" & Val_Filter & "*")
    End Select

End Function
