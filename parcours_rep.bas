'Declaration
Public oFSO, oFld, oSubFolder
Public Tab_Rep() As String 'Tableau en mémoire dans lequel sera stocké les noms des fichiers
Public i As Long

'=====================================================================================
'Fonction récursive pour trouver tous les sous-répertoires avec FileSystemObject
'=====================================================================================
 
Sub ParcoursRepT()
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    i = 0
    Call ParcoursRep("C:\Users")
End Sub
 
 
Sub ParcoursRep(ByVal stRep As String)
On Error GoTo Traitement_erreur
    'Debug.Print "Traite : " & stRep
    If oFSO.FolderExists(stRep) Then
    Set oFld = oFSO.GetFolder(stRep)
        If (oFld.Attributes And 1024) = 0 Then 'Test si le répertoire n'est pas un répertoire virtuel (MaMusique, MesVidéos, ...)
            If oFld.SubFolders.Count > 0 Then 'Teste le nombre de sous-répertoire dans stRep
                For Each oSubFolder In oFld.SubFolders
                    ReDim Preserve Tab_Rep(i)
                    Tab_Rep(i) = oSubFolder.Path    'Ecrit dans le tableau en mémoire
                    Cells(i + 1, 1) = oSubFolder.Path 'Ecrit le résultat en colonne A
                    i = i + 1
                    DoEvents
                    ParcoursRep oSubFolder.Path 'appel récursif de la procédure
                Next
            End If
        End If
    End If
Exit Sub
Traitement_erreur:
    Debug.Print Err.Number & " " & Err.Description
    Resume Next
End Sub

