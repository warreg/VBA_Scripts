'Choix de la base la plus récente parmis les bases selectionnées de la liste
            'Call ConnexionSQL("select top 1 [UpdDT], bdd FROM (SELECT top 1 [UpdDT],'FUSIONDB_03012017_BM' as bdd FROM [FUSIONDB_03012017_BM].[dbo].[USER]  order by UpdDT ASC Union all SELECT top 1 [UpdDT],'FUSIONDB_20090310_BM' as bdd FROM [FUSIONDB_20090310_BM].[dbo].[USER] order by UpdDT ASC) res  order by res.[UpdDT] DESC",1, 2)
            
            requete = "SELECT Value01 FROM [" & DbSelect & "].[dbo].[WELL_RESULT] WR, [" & DbSelect & "].[dbo].[WELL] WE, [" & DbSelect & "].[dbo].[SAMPLE] SA "
            requete = requete & " WR.WellID = WE.WellID  and WE.[SampleID]=SA.[SampleID]  and SA.SampleIDName='" & SampIdNa & "' "
            requete = requete & " and ResultType='08'  and Value02='" & Locus_tab(i) & "' "
            
            requete_last_bdd = "SELECT top 1 [UpdDT],'" & DbSelect & "' as bdd FROM [" & DbSelect & "].[dbo].[USER]  order by UpdDT ASC "
            
           ' For
            
            requete = " union all SELECT Value01 FROM [" & DbSelect & "].[dbo].[WELL_RESULT] WR, [" & DbSelect & "].[dbo].[WELL] WE, [" & DbSelect & "].[dbo].[SAMPLE] SA "
            requete = requete & " WR.WellID = WE.WellID  and WE.[SampleID]=SA.[SampleID]  and SA.SampleIDName='" & SampIdNa & "' "
            requete = requete & " and ResultType='08'  and Value02='" & Locus_tab(i) & "' "
           
           requete_last_bdd = "union all SELECT top 1 [UpdDT],'" & DbSelect & "' as bdd FROM [" & DbSelect & "].[dbo].[USER]  order by UpdDT ASC "
           
           ' next
           
           requete = "select  distinct Value01 from ( " & requete & " ) res "
           
            requete_last_bdd = "select top 1 [UpdDT], bdd FROM ( " & requete_last_bdd & " ) res  order by res.[UpdDT] DESC"
            
