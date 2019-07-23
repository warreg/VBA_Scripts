
Sub Main_test()
'====================
'Procedure principale
'====================

End Sub


Function ConnexionSQL(query$, Log_li&, Log_col&)
   'Query: requete ; Log_li et Log_col: num ligne et num colone dans l'onglet(invisible) Log
   Dim cnx As ADODB.Connection
   Dim rst As New ADODB.Recordset
   ConnexionSQL = False
   'Initialisation de la chaine de connexion
   Set cnx = New ADODB.Connection
   Set rst = New ADODB.Recordset
   cnx.ConnectionString = "UID=" & USFConnect.LoginBox.Value & ";PWD=" & USFConnect.PwdBox.Value & ";" & "DRIVER={SQL Server};Server=" & USFConnect.ServeurBox.Value & ";" 'Database=" & DbName & ";"
   'Verifie que la connexion est bien fermee
   If cnx.State = adStateOpen Then
      cnx.Close
   End If
   On Error GoTo Errconnex
   'Connexion a la base de donnees
   cnx.Open
   'Attendre que la connexion soit etablie
   While (cnx.State = adStateConnecting)
      DoEvents
   Wend
   'Arret verification si erreur de connexion
   On Error GoTo 0
   'Si la requete n'est pas executé on passe a la requete suivante
   On Error Resume Next 'On Error GoTo ErrHandle
   rst.Open query, cnx, adOpenForwardOnly
   ConnexionSQL = True
   'Ramener le curseur au début
   rst.MoveFirst
   Workbooks("DataExtract_V3.xlsm").Worksheets("Log").Cells(Log_li, Log_col).CopyFromRecordset rst
   Exit Function
   cnx.Close
   rst.Close
   
'En cas d'erreur
Errconnex:
   MsgBox "Connexion impossible !" & Chr(10) & "Vérifier parametres de connexion", vbCritical, "Echec connexion"
   Exit Function
   
'ErrHandle:
'   MsgBox Err.Description
'   Resume Next
   
End Function


Sub SSO()

   '***
   Dim Temps_deb As Single, Duree As Single
   Temps_deb = Timer
   '***

   '-------------------------------------------------------
   'Amb1$: 1er ambiguité / Amb2$: 2iem ambiguité / SampIdNa$: l'identifiant saisi dans l'onglet Données Entree / DbSelect$: la derniere base stocké dans l'onglet log
   Dim Amb1$, Amb2$, SampIdNa$, DbSelect$
   'req_sero$: requete serologie simple / req_amb$: requete ambiguité simple / req_list_amb1$: requete1 de la liste des ambiguités correspdt / req_list_amb2$: requete de la liste des ambiguités correspdt / req_last_bd$: requete pour recupérer la derniere base parmi les bd selectionnés
   Dim req_sero$, req_amb$, req_list_amb1$, req_list_amb2$, req_last_bd$
   'req_sero_final$: req_sero complete / req_amb_final$: req_amb complete / req_last_bd_final$: req_last_bd complete
   Dim req_sero_final$, req_amb_final$, req_last_bd_final$
   'li&: num ligne onglet Donnes Entree / col&: num colone onglet Log / i&: compteur de locus / lig&: compteur lignes onglet Log/ li_bd&: compteur bd selectionné onglet Log
   Dim li&, col&, i&, lig&, li_bd&
   'BdTab(): tableau stockage les bases selectionnées  / Locus_tab: tableau stockage difts locus
   Dim BdTab() As String, Locus_tab As Variant
   Locus_tab = Array("A", "B", "C", "DPA1", "DPB1", "DQA1", "DQB1", "DRB1", "DRB345")
   'reg: objet regex pour recherche de pattern
   Dim regex1 As Object, regex2 As Object
   '^[a-zA-Z]{1,2}\*((\d{1,2})[^a-zA-Z]?([a-zA-Z]{2,5}))\s[a-zA-Z]{1,2}\*((\d{1,2})[^a-zA-Z]?([a-zA-Z]{2,5}$))
   li = 2
   col = 2
   lig = 1
   li_bd = 1
   '-------------------------------------------------------
   
   '-------------------------------------------------------
   'Optimisation du temps d'execution
   Application.StatusBar = "Pour votre Santé Manger 5 Fruits et Légumes par Jour...  °L° "    'Message dans barre de statut
   Application.EnableCancelKey = xlDisabled           'Desactivation touche Echap
   Application.ScreenUpdating = False                 'Empeche le rafraichissement de la page
   Application.EnableEvents = False                   'Empeche d'autres programme d'interompre le lancement
   '-------------------------------------------------------
   
   With Workbooks("DataExtract_V3.xlsm")
   
      '----------------------------------------------------
      'Récupération des bases selectionnées dans un tableau
      With .Worksheets("Log")
         While .Cells(li_bd, 7) <> ""              'Pour chaque ligne
            ReDim Preserve BdTab(li_bd)            'On redimensionne le tableau avec la nouvelle ligne
            BdTab(li_bd) = .Cells(li_bd, 7).Value  'Et le contenu de la ligne est stocké dans le tableau
            li_bd = li_bd + 1                      'On passe a la ligne suivante
         Wend
      End With
      '----------------------------------------------------
   
      'Requete pour rechercher la base la plus récente parmis les bases selectionnées de la liste dans la premiere base(par defaut)
      req_last_bd = "SELECT top 1 [UpdDT],'" & BdTab(1) & _
                    "' as bdd FROM [" & BdTab(1) & _
                    "].[dbo].[USER]  order by UpdDT ASC "
      
      '----------------------------------------------------
      If UBound(BdTab) > 1 Then  'Si on a choisi plusieur bases
         For li_bd = 2 To UBound(BdTab)
            req_last_bd = req_last_bd & " UNION ALL " & _
                          "SELECT top 1 [UpdDT],'" & BdTab(li_bd) & _
                          "' as bdd FROM [" & BdTab(li_bd) & _
                          "].[dbo].[USER]  order by UpdDT ASC "
         Next li_bd
      End If
      '----------------------------------------------------
      
      req_last_bd_final = "SELECT top 1 bdd FROM ( " & req_last_bd & " ) res  order by res.[UpdDT] DESC"
      Call ConnexionSQL(req_last_bd_final, lig, 2)
      DbSelect = .Worksheets("Log").Cells(1, 2).Value                         'DbSelect = la premiere base de la colonne 2
   
   
   '-------------------------------------------------------
      While .Worksheets("Donnees_Entree").Cells(li, 1) <> ""                  'Pour chaque identifiant collé dans "Donnes Entree"
         
         SampIdNa = .Worksheets("Donnees_Entree").Cells(li, 1).Value          'SampIdNa = id collés
         
         '-------------------------------------------------
         For i = 0 To UBound(Locus_tab)                                       'Pour chaque locus dans Locus_tab

            'Requete pour recherhe Serologie dans la premiere base(par defaut)
            req_sero = "SELECT Value01  COLLATE DATABASE_DEFAULT as Value01 FROM [" & BdTab(1) & _
                       "].[dbo].[WELL_RESULT] WR, [" & BdTab(1) & "].[dbo].[WELL] WE, [" & BdTab(1) & _
                       "].[dbo].[SAMPLE] SA where WR.WellID = WE.WellID  and WE.[SampleID]=SA.[SampleID]  and SA.SampleIDName='" & SampIdNa & _
                       "'  and ResultType='08'  and Value02='" & Locus_tab(i) & "' "

            'Requete pour recherche Ambiguité dans la premiere base(par defaut)
            req_amb = "SELECT  Value01  COLLATE DATABASE_DEFAULT as Value01  FROM [" & BdTab(1) & _
                      "].[dbo].[WELL_RESULT] WR, [" & BdTab(1) & "].[dbo].[WELL] WE, [" & BdTab(1) & _
                      "].[dbo].[SAMPLE] SA where WR.WellID = WE.WellID and WE.[SampleID]=SA.[SampleID] and SA.SampleIDName='" & SampIdNa & _
                      "' and ResultType='06' and Value02='" & Locus_tab(i) & "' "

            '----------------------------------------------
            If UBound(BdTab) > 1 Then  'Si on a choisi plusieur bases
            
               '-------------------------------------------
               'On fait un UNON ALL avec le reste a partir de chaque deuxieme base
               For li_bd = 2 To UBound(BdTab)

                  req_sero = req_sero & " UNION ALL " & _
                             "SELECT Value01  COLLATE DATABASE_DEFAULT as Value01 FROM [" & BdTab(li_bd) & _
                             "].[dbo].[WELL_RESULT] WR, [" & BdTab(li_bd) & "].[dbo].[WELL] WE, [" & BdTab(li_bd) & _
                             "].[dbo].[SAMPLE] SA where WR.WellID = WE.WellID  and WE.[SampleID]=SA.[SampleID]  and SA.SampleIDName='" & SampIdNa & _
                             "'  and ResultType='08'  and Value02='" & Locus_tab(i) & "' "

                  req_amb = req_amb & " UNION ALL " & _
                            "SELECT  Value01  COLLATE DATABASE_DEFAULT as Value01  FROM [" & BdTab(li_bd) & _
                            "].[dbo].[WELL_RESULT] WR, [" & BdTab(li_bd) & "].[dbo].[WELL] WE, [" & BdTab(li_bd) & _
                            "].[dbo].[SAMPLE] SA where WR.WellID = WE.WellID and WE.[SampleID]=SA.[SampleID] and SA.SampleIDName='" & SampIdNa & _
                            "' and ResultType='06' and Value02='" & Locus_tab(i) & "' "


               Next li_bd
               '-------------------------------------------
               
               'Puis on écrit la requete complete
               req_sero_final = "SELECT  distinct Value01  COLLATE DATABASE_DEFAULT as Value01 from ( " & req_sero & " )res"
               req_amb_final = "SELECT  distinct Value01  COLLATE DATABASE_DEFAULT as Value01 from ( " & req_amb & " )res"
               
               'Et on execute la requete complete
               Call ConnexionSQL(req_sero_final, lig, 4)
               Call ConnexionSQL(req_amb_final, lig, 5)

            Else                       'Si on a une seule base selectionnée
               'On écrit la requete complete
               req_sero_final = "SELECT  distinct Value01  COLLATE DATABASE_DEFAULT as Value01 from ( " & req_sero & " )res"
               req_amb_final = "SELECT  distinct Value01  COLLATE DATABASE_DEFAULT as Value01 from ( " & req_amb & " )res"
               
               'Et on éxécute la requete complete
               Call ConnexionSQL(req_sero_final, lig, 4)
               Call ConnexionSQL(req_amb_final, lig, 5)

            End If
            '----------------------------------------------
            
            '----------------------------------------------
            If .Worksheets("Log").Cells(lig, 5).Value <> "" Then                'Si ambiguité est trouvé
               res_ambig = .Worksheets("Log").Cells(lig, 5).Value
               
               '-------------------------------------------
               Set regex1 = CreateObject("vbscript.regexp")
               With regex1
                 .Pattern = "^[a-zA-Z]{1,2}\*((\d{1,2})[^a-zA-Z]?([a-zA-Z]{2,5}))\s[a-zA-Z]{1,2}\*((\d{1,2})[^a-zA-Z]?([a-zA-Z]{2,5}$))"
                 .Global = True
                 .MultiLine = False
               End With
               Set matches = regex1.Execute(res_ambig)
               For Each Match In matches
                  al_inc1 = Match.SubMatches(0) ' al_inc: allele inconnu  'spl_al1
                  al_inc2 = Match.SubMatches(3) 'spl_al2
                  two_digit1 = Match.SubMatches(1)
                  two_digit2 = Match.SubMatches(4)
                  ambig1 = Match.SubMatches(2)  'ambig1
                  ambig2 = Match.SubMatches(5)  'spl_amb2
               Next Match
               '-------------------------------------------
               
               '-------------------------------------------
               Set regex2 = CreateObject("vbscript.regexp")
               regex2.Pattern = "^[a-zA-Z]{2,5}$" 'pattern: commençant par un mot de 2 a 5 lettre et se terminant aussi par une lettre
               '-------------------------------------------
               
               '-------------------------------------------
               If regex2.test(ambig1) = True Then 'Si correspond a une vrai ambiguité d'apres le pattern
                  Amb1 = ambig1
                  'Recherche de la liste des amibiguités correspondant a Amb1
                  req_list_amb1 = " SELECT [NmdpDef] FROM [" & DbSelect & "].[dbo].[NMDP_CODE_DETAIL] where [NmdpID] = '" & Amb1 & "' "
                  Call ConnexionSQL(req_list_amb1, lig, 6)
                  list_ambig1 = .Worksheets("Log").Cells(lig, 6).Value                'On definit la liste des ambiguités 1
                  spl_alle1 = Split(.Worksheets("Log").Cells(lig, 6).Value, "/")(0)   'On definit l'allele1
      
                  If Len(spl_alle1) > 2 Then
                     .Worksheets("Result").Cells(li + 1, col + 2) = spl_alle1                              'Copie dans Result
                  Else
                     .Worksheets("Result").Cells(li + 1, col + 2) = two_digit1 & "::" & spl_alle1
                  End If
                     .Worksheets("Result").Cells(li + 1, col + 4) = ambig1 & "#" & list_ambig1           'Copie dans Result

               Else  'Si pas d'ambiguité trouvé
                  With .Worksheets("Result")
                     .Cells(li + 1, col + 2) = al_inc1                                'Copie dans Result
                     .Cells(li + 1, col + 4) = "-"                                    'Copie dans Result
                  End With
               End If
               '-------------------------------------------
               
               '-------------------------------------------
               If regex2.test(ambig2) = True Then   'Si correspond a une vrai ambiguité d'apres le pattern
                  Amb2 = ambig2
                  'Recherche de la liste des amibiguités correspondant a Amb2
                  req_list_amb2 = " SELECT [NmdpDef] FROM [" & DbSelect & "].[dbo].[NMDP_CODE_DETAIL] where [NmdpID] = '" & Amb2 & "' "
                  Call ConnexionSQL(req_list_amb2, lig + 1, 6)
                  list_ambig2 = .Worksheets("Log").Cells(lig + 1, 6).Value              'On definit la liste des ambiguité 2
                  spl_alle2 = Split(.Worksheets("Log").Cells(lig + 1, 6).Value, "/")(0) 'On definit l'allele2
                  
                  If Len(spl_alle2) > 2 Then
                     .Worksheets("Result").Cells(li + 1, col + 3) = spl_alle2                              'Copie dans Result
                  Else
                     .Worksheets("Result").Cells(li + 1, col + 3) = two_digit2 & "::" & spl_alle2
                  End If
                     .Worksheets("Result").Cells(li + 1, col + 5) = ambig2 & "#" & list_ambig2           'Copie dans Result

               Else              'Si pas d'ambiguité trouvé
                   With .Worksheets("Result")
                     .Cells(li + 1, col + 3) = al_inc2                                  'Copie dans Result
                     .Cells(li + 1, col + 5) = "-"                                      'Copie dans Result
                  End With
               End If
               '-------------------------------------------
               
            Else  'Si pas de résultats pour ce locus on mets des "-" a la place
                  With .Worksheets("Result")
                     .Cells(li + 1, col + 2) = "-"
                     .Cells(li + 1, col + 3) = "-"
                     .Cells(li + 1, col + 4) = "-"
                     .Cells(li + 1, col + 5) = "-"
                  End With
            End If
            '----------------------------------------------
            
            '----------------------------------------------
            If .Worksheets("Log").Cells(lig, 4).Value <> "" Then  'Si une serologie est trouvé
               With .Worksheets("Log")
                 spl_sero1 = Split(.Cells(lig, 4).Value, " ")(0)  'On definit les serologies
                 spl_sero2 = Split(.Cells(lig, 4).Value, " ")(1)  'On definit les serologies
               End With
               With .Worksheets("Result")
                 .Cells(li + 1, col) = spl_sero1                  'Et on les copie des sero dans Result
                 .Cells(li + 1, col + 1) = spl_sero2              'Et on les copie des sero dans Result
               End With
            Else                                                  'Si pas de sérologie trouvé (dans le cas des DQ/DP.. le + souvent)
               With .Worksheets("Result")
                  .Cells(li + 1, col) = "-"                       'On remplace par des tirets
                  .Cells(li + 1, col + 1) = "-"                   'On remplace par des tirets
               End With
            End If
            '----------------------------------------------
            
            '----------------------------------------------
            Worksheets("Result").Cells(li + 1, 1) = SampIdNa      'On copie chaque Id dans Result
            col = col + 6                                         'Pour passer au locus suivant dans Result
            lig = lig + 2                                         'On saute une ligne dans Log pour traiter le locus suivant
            
            Debug.Print res_ambig, al_inc1, al_inc2, two_digit1, two_digit2, ambig1, ambig2, Len(spl_alle1), Len(spl_alle2)
            
         Next i
         '-------------------------------------------------
         col = 2                                                  'On reinitialise col pour ecrire a la ligne
         li = li + 1                                              'On passe a l'Id suivant
      Wend
      '----------------------------------------------------
   End With
   
   '-------------------------------------------------------
   Application.ScreenUpdating = True
   Application.EnableCancelKey = xlInterrupt
   Application.StatusBar = False
   Application.EnableEvents = True
   '-------------------------------------------------------
   
   
   
   '***
   Duree = Timer - Temps_deb
   MsgBox Duree
   '***
   
End Sub



