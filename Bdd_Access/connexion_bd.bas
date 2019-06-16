
Sub Requete_Access_SELECT()
Dim Ma_Base As New ADODB.Connection
Dim Ma_Requete As New ADODB.Recordset
Dim Base_Data As String
Dim SQL_Req As String
    'Location et Nom de la base de données
    Base_Data = "C:\Users\Base_Access_Exemple.accdb"
    
    ' Ouvrir la connexion avec la base de données
    Set Ma_Base = New ADODB.Connection
    Ma_Base.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
       "Data Source=" & Base_Data & ";"
    
    ' Création de la requête
    SQL_Req = "SELECT * From Table1"
    
    'Exécution de la requête
    Ma_Requete.Open SQL_Req, Ma_Base, adOpenForwardOnly
    
    'Test pour s'assurer que l'accès à la table est sans erreur
    If Ma_Requete.Status = 0 Then
        'copie de la requête en cellule A1
        Range("A1").CopyFromRecordset Ma_Requete
    End If
    
    'Fermeture du RecordSet et de la connexion
    Ma_Requete.Close
    
    'Fermeture de la connexion à la base pour libérer de la mémoire
    Ma_Base.Close
End Sub


'Même code que précédement mais avec une boucle pour retourner le résultat dans la feuille de calcul
Sub Requete_Access_SELECT_Boucle()
Dim Ma_Base As New ADODB.Connection
Dim Ma_Requete As New ADODB.Recordset
Dim Base_Data As String
Dim SQL_Req As String
Dim i_Ligne As Long     'Indice ligne

    'Location et Nom de la base de données
    Base_Data = "C:\Users\Base_Access_Exemple.accdb"
    
    ' Ouvrir la connexion avec la base de données
    Set Ma_Base = New ADODB.Connection
    Ma_Base.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
       "Data Source=" & Base_Data & ";"
    
    ' Création de la requête
    SQL_Req = "SELECT * From Table1"
    
    'Exécution de la requête
    Ma_Requete.Open SQL_Req, Ma_Base, adOpenForwardOnly
    
    'Test pour s'assurer que l'accès à la table est sans erreur
    If Ma_Requete.Status = 0 Then
        'Boucle sur tous les éléments collectés par la requête
        i_Ligne = 1
        ' Nous nous positionnons sur le premier enregistrement
        Ma_Requete.MoveFirst
        
        While Ma_Requete.EOF = False
            Cells(i_Ligne, 1) = Ma_Requete.Fields("Prénom")
            Cells(i_Ligne, 2) = Ma_Requete.Fields("Valeur")
            
            Ma_Requete.MoveNext     'Passe à l'enregistrement suivant
            i_Ligne = i_Ligne + 1
        Wend
    End If
    
    'Fermeture du RecordSet et de la connexion
    Ma_Requete.Close
    
    'Fermeture de la connexion à la base pour libérer de la mémoire
    Ma_Base.Close
End Sub


