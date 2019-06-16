
Sub Requete_Access_Nouvel_Enregistrement()
Dim Ma_Base As New ADODB.Connection
Dim Ma_Requete As New ADODB.Recordset
Dim Base_Data As String 'Chemin d'accès
Dim SQL_Req As String   'Requête
Dim i_Ligne As Long     'Indice ligne

    'Chemin et Nom de la base de données
    Base_Data = "C:\Users\Frédéric\OneDrive\21_Livres\01_Macro_Langage_VBA\v4\Excel\Base_Access_Exemple.accdb"
    
    ' Ouvre la connexion à la base
    Set Ma_Base = New ADODB.Connection
    Ma_Base.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & Base_Data
    
    ' Ouvre un nouveau recordset
    Set Ma_Requete = New ADODB.Recordset
    Ma_Requete.Open "Table1", Ma_Base, adOpenForwardOnly, adLockOptimistic
    
    i_Ligne = 1
    
    'Boucle sur toutes les cellules de la colonne A
    While Cells(i_Ligne, 1) <> ""
        
        'Insertion des données dans la base Access
        With Ma_Requete
            .AddNew 'Ajout d'un nouvel enregistrement vide
            .Fields(0) = Cells(i_Ligne, 1)  'Ecriture dans la première colonne (indice 0)
            .Fields("Valeur") = Cells(i_Ligne, 2)   ' Ecriture dans la colonne "Valeur"
            .Update 'Mise à jour de l'enregistrement
        End With
        
        i_Ligne = i_Ligne + 1
    Wend
    'Fermeture du Recordset
    Ma_Requete.Close
    
    'Fermeture de la base
    Ma_Base.Close
    
End Sub


Sub Requete_Access_INSERT()
Dim Ma_Base As New ADODB.Connection
Dim Ma_Requete As New ADODB.Recordset
Dim Base_Data As String 'Chemin d'accès
Dim SQL_Req As String   'Requête
Dim i_Ligne As Long     'Indice ligne

    'Chemin et Nom de la base de données
    Base_Data = "C:\Users\Frédéric\OneDrive\21_Livres\01_Macro_Langage_VBA\v4\Excel\Base_Access_Exemple.accdb"
    
    ' Ouvre la connexion à la base
    Set Ma_Base = New ADODB.Connection
    Ma_Base.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & Base_Data
    
    ' Ouvre un nouveau recordset
    Set Ma_Requete = New ADODB.Recordset
    Ma_Requete.Open "Table1", Ma_Base, adOpenForwardOnly, adLockOptimistic
    
    i_Ligne = 1
    
    'Boucle sur toutes les cellules de la colonne A
    While Cells(i_Ligne, 1) <> ""
        
        SQL_Req = "INSERT INTO Table1 VALUES("
        SQL_Req = SQL_Req & Chr(34) & Cells(i_Ligne, 1) & Chr(34) & ","
        SQL_Req = SQL_Req & Cells(i_Ligne, 2) & ");"
        ' le code Chr(34) est utilisé pour indiquer que la variable est à mettre
        ' entre guillemet dans la requête SQL
        
        'Exécution de la requête
        Ma_Base.Execute SQL_Req
        
        i_Ligne = i_Ligne + 1
    Wend
    'Fermeture du Recordset
    Ma_Requete.Close
    
    'Fermeture de la base
    Ma_Base.Close
    
End Sub


