Sub fusion_ligne()
'FUSIONNER LES LIGNES IDENTIQUES ENTRE ELLES

  'li= num lign ; li_fin= num lign fin;
  Dim li&, li_fin&
  li = 2
  'Empecher affichage msg d'erreur d'excel
  Application.DisplayAlerts = False
  While Cells(li, 1) <> ""
    li_fin = li
    While Cells(li_fin, 1) = Cells(li_fin + 1, 1)
      li_fin = li_fin + 1
    Wend
    'Instruction de fusion+align centre+gauche
    With Range(Cells(li, 1), Cells(li_fin, 1))
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .MergeCells = True
    End With
    li = li_fin + 1
    Debug.Print "valeur de li:" & li
    Debug.Print "valeur de li_fin:" & li_fin
  Wend
  'Autoriser affichage msg d'erreur
  Application.DisplayAlerts = True
  
End Sub
