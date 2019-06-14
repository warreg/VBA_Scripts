Sub Mettre_meme_lign()
'ORGANISER SON TABLEAU EN PRESENTANT SUR UNE MEME LIGNE, UNE DATE ET TOUS LES PRENOMS CORRESPONDANTS A CETTE DATE
  
  'li: lign; li_w: ligne write (ligne d'écriture ); col_w: col write
  Dim li&, li_w&, col_w&
  li = 1
  li_w = 1
  While Cells(li, 1) <> ""
    col_w = 5
    Do
      Cells(li_w, col_w) = Cells(li, 2)
      col_w = col_w + 1
      li = li + 1
    Loop While Cells(li - 1, 1) = Cells(li, 1)
    Cells(li_w, 4) = Cells(li - 1, 1)
    li_w = li_w + 1
  Wend
  
End Sub
