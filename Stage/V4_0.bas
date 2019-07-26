Dim ProgressIndicator As USFBarProgress


'=============================================
'SCRIPT DE CONNEXION AUX BASES ET DE REQUETAGE
'=============================================
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
   Workbooks("DataExtract_V4.xlsm").Worksheets("Log").Cells(Log_li, Log_col).CopyFromRecordset rst
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

'===============================
'SCRIPT D'EXTRACTION DONNEES SSO
'===============================
Sub SSO()

   '*** TIMER ET BARRE DE PROGRESSION ***
   Dim Temps_deb As Single, Duree As Single, compt&, comptMAX&, progress#, nb_li&, li_id&
   Temps_deb = Timer
   li_id = 2
   While Worksheets("Donnees_Entree").Cells(li_id, 1) <> ""
      li_id = li_id + 1
   Wend
   nb_li = li_id - 2
   comptMAX = nb_li * 9
   compt = 0
   '*** TIMER ET BARRE DE PROGRESSION ***

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
   'regex: objet regexp pour recherche de pattern
   Dim regex1 As Object, regex2 As Object
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
   
   '***
   Set ProgressIndicator = New USFBarProgress
   ProgressIndicator.Show vbModeless
   '***
   
   
   
   With Workbooks("DataExtract_V4.xlsm")
   
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
      'DbSelect: la base la plus récente des bases selectionnées
      DbSelect = .Worksheets("Log").Cells(1, 2).Value
   
   
      '----------------------------------------------------
      'Pour chaque identifiant collé dans "Donnes Entree"
      While .Worksheets("Donnees_Entree").Cells(li, 1) <> ""
         
         'SampIdNa = id collés
         SampIdNa = .Worksheets("Donnees_Entree").Cells(li, 1).Value
         
         '-------------------------------------------------
         'Pour chaque locus dans Locus_tab
         For i = 0 To UBound(Locus_tab)
            
            '***
            compt = compt + 1
            '***

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
            'Si plusieurs bases sont selectionnées
            If UBound(BdTab) > 1 Then
            
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
            
            'Si on a une seule base selectionnée
            Else
               'On écrit la requete complete
               req_sero_final = "SELECT  distinct Value01  COLLATE DATABASE_DEFAULT as Value01 from ( " & req_sero & " )res"
               req_amb_final = "SELECT  distinct Value01  COLLATE DATABASE_DEFAULT as Value01 from ( " & req_amb & " )res"
               
               'Et on éxécute la requete complete
               Call ConnexionSQL(req_sero_final, lig, 4)
               Call ConnexionSQL(req_amb_final, lig, 5)

            End If
            '----------------------------------------------
            
            '----------------------------------------------
            'Si ambiguité est trouvé
            If .Worksheets("Log").Cells(lig, 5).Value <> "" Then
               'On stocke le contenu de la cellule dans res_ambig
               res_ambig = .Worksheets("Log").Cells(lig, 5).Value
               
               '-------------------------------------------
               'Objet regex pour decouper la chaine d'amibiguité
               Set regex1 = CreateObject("vbscript.regexp")
               With regex1
                 .Pattern = "^[a-zA-Z]{1,3}\d*\*((\d{1,2})\:?(([a-zA-Z]{2,5}\d*|\d+)))\s[a-zA-Z]{1,3}\d*\*((\d{1,2})\:?(([a-zA-Z]{2,5}\d*|\d+)))$"
                 .Global = True
               End With
               Set matches = regex1.Execute(res_ambig)
               For Each Match In matches
                  al_inc1 = Match.SubMatches(0)          'al_inc1: allele inconnu = la chaine apres les ":" dans les cas ou on par exp XX1
                  al_inc2 = Match.SubMatches(4)          'al_inc2: allele inconnu = la chaine apres les ":" dans les cas ou on par exp XX2
                  two_digit1 = Match.SubMatches(1)       'two_digit1: les 2 1er chiffres de l'ambiguité1
                  two_digit2 = Match.SubMatches(5)       'two_digit1: les 2 1er chiffres de l'ambiguité2
                  ambig1 = Match.SubMatches(2)           'ambig1: 1er code ambiguité a rechercher dans la derniere selectionnée
                  ambig2 = Match.SubMatches(6)           'ambig2: 2iem code ambiguité2 a rechercher dans la derniere selectionnée
               Next Match
               '-------------------------------------------
               
               '-------------------------------------------
               'Objet regex pour vérifier si le code d'ambiguité  correspond bien a un vrai code d'ambiguité
               Set regex2 = CreateObject("vbscript.regexp")
               'pattern: commençant par un mot de 2 a 5 lettre et se terminant aussi par une lettre
               regex2.Pattern = "^[a-zA-Z]{2,5}$"
               '-------------------------------------------
               
               '-------------------------------------------
               If regex2.test(ambig1) Then  'Si ambig 1 correspond a une vrai ambiguité d'apres le pattern
                  Amb1 = ambig1
                  'Recherche de la liste des amibiguités correspondant a Amb1
                  req_list_amb1 = " SELECT [NmdpDef] FROM [" & DbSelect & "].[dbo].[NMDP_CODE_DETAIL] where [NmdpID] = '" & Amb1 & "' "
                  Call ConnexionSQL(req_list_amb1, lig, 6)
                  'On definit la liste des ambiguités 1
                  list_ambig1 = .Worksheets("Log").Cells(lig, 6).Value
                  'On definit  le 1er élément comme l'allele1
                  spl_alle1 = Split(.Worksheets("Log").Cells(lig, 6).Value, "/")(0)
                       
                  '----------------------------------------
                  'Si l'allele 1 trouvé n'est pas constitué que de deux digit
                  If Len(spl_alle1) > 2 Then
                     'Copie de l'allele1 dans Result
                     .Worksheets("Result").Cells(li + 1, col + 2) = spl_alle1
                  'Si non si c'est constitué que de deux digit
                  Else
                     'On concatene l'allele1(2 digit) avec la variable two_digit
                     .Worksheets("Result").Cells(li + 1, col + 2) = two_digit1 & ":" & spl_alle1
                  End If
                  '----------------------------------------
                  'Et dans tous les cas copie du code ambiguité concaténé avec la liste des amiguités correspondante dans Result
                  .Worksheets("Result").Cells(li + 1, col + 4) = ambig1 & "#" & list_ambig1

               Else  'Si le code d'ambiguité ne corresond pas un code NMDP
                  With .Worksheets("Result")
                     'Copie de al_inc1 dans Result
                     .Cells(li + 1, col + 2) = al_inc1
                     'Des tirets a la place de la liste des ambiguités  dans Result
                     .Cells(li + 1, col + 4) = "-"
                  End With
               End If
               '-------------------------------------------
               
               '-------------------------------------------
               'Si ambig2 correspond a une vrai ambiguité d'apres le pattern
               If regex2.test(ambig2) Then
                  Amb2 = ambig2
                  'Recherche de la liste des amibiguités correspondant a Amb2
                  req_list_amb2 = " SELECT [NmdpDef] FROM [" & DbSelect & "].[dbo].[NMDP_CODE_DETAIL] where [NmdpID] = '" & Amb2 & "' "
                  Call ConnexionSQL(req_list_amb2, lig + 1, 6)
                  'On definit la liste des ambiguité 2
                  list_ambig2 = .Worksheets("Log").Cells(lig + 1, 6).Value
                  'On definit le premier élément comme l'allele2
                  spl_alle2 = Split(.Worksheets("Log").Cells(lig + 1, 6).Value, "/")(0)
                  
                  '----------------------------------------
                  'Si l'allele2 trouvé n'est pas constitué que de deux digit
                  If Len(spl_alle2) > 2 Then
                     'Copie de l'allele2 dans Result
                     .Worksheets("Result").Cells(li + 1, col + 3) = spl_alle2
                  'Si non si c'est constitué que de deux digit
                  Else
                     'On concatene l'allele2(2 digit) avec la variable two_digit2
                     .Worksheets("Result").Cells(li + 1, col + 3) = two_digit2 & ":" & spl_alle2
                  End If
                  '----------------------------------------
                  'Et dans tous les cas copie du code ambiguité concaténé avec la liste des amiguités correspondante dans Result
                  .Worksheets("Result").Cells(li + 1, col + 5) = ambig2 & "#" & list_ambig2

               Else      'Si le code d'ambiguité ne corresond pas un code NMDP
                   With .Worksheets("Result")
                     'Copie de al_inc2 dans Result
                     .Cells(li + 1, col + 3) = al_inc2
                     'Des tirets a la place de la liste des ambiguités  dans Result
                     .Cells(li + 1, col + 5) = "-"
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
            'Si une serologie est trouvé
            If .Worksheets("Log").Cells(lig, 4).Value <> "" Then
               With .Worksheets("Log")
                  'Récupere dans spl_sero1 le 1er élémt du split par espace
                  spl_sero1 = Split(.Cells(lig, 4).Value, " ")(0)
                  'Récupere dans spl_sero2 le 2iem élémt du split par espace
                  spl_sero2 = Split(.Cells(lig, 4).Value, " ")(1)
               End With
               'Et on les copie dans Result
               With .Worksheets("Result")
                 .Cells(li + 1, col) = spl_sero1
                 .Cells(li + 1, col + 1) = spl_sero2
               End With
            'Si pas de sérologie trouvé (dans le cas des DQ/DP.. le + souvent)
            Else
               With .Worksheets("Result")
                  'On remplace par des tirets
                  .Cells(li + 1, col) = "-"
                  .Cells(li + 1, col + 1) = "-"
               End With
            End If
            '----------------------------------------------
            
            '***
            Debug.Print "res: " & res_ambig, "sero1: " & spl_sero1, "sero2: " & spl_sero2, "al_inc1: " & al_inc1, "al_inc2: " & al_inc2, "digit1: " & two_digit1, "digit2: " & two_digit2, "ambig1: " & Amb1, "amb2: " & ambig2
            '***
            
            'On copie chaque Id dans Result
            Worksheets("Result").Cells(li + 1, 1) = SampIdNa
            'Pour passer au locus suivant dans Result
            col = col + 6
            'On saute une ligne dans Log pour traiter le locus suivant
            lig = lig + 2
            'Et on réinitialise tous les variables
            res_ambig = ""
            al_inc1 = ""
            al_inc2 = ""
            two_digit1 = ""
            two_digit2 = ""
            ambig1 = ""
            ambig2 = ""
            
            
            '***
            progress = compt / comptMAX
            Call UpdateProgress(progress)
            '***
            
         Next i
         '-------------------------------------------------
         col = 2      'On reinitialise col pour ecrire a la ligne
         li = li + 1  'On passe a l'Id suivant
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
   Unload ProgressIndicator
   Set ProgressIndicator = Nothing
   Duree = Timer - Temps_deb
   Workbooks("DataExtract_V4.xlsm").Worksheets("Log").Cells(1, 3) = Duree
   '***
   
End Sub


'======================================
'MISE A JOUR DE LA BARRE DE PROGRESSION
'======================================
Sub UpdateProgress(prog)
   With ProgressIndicator
      .LabelProgress.Caption = "Traitement..  " & CInt(prog * 100) & " %"
      If CInt(prog * 100) > 99 Then .LabelProgress.Caption = "Terminé !"
      .ImgProgress.Width = prog * 332
   End With
   DoEvents
End Sub
