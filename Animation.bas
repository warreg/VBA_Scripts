Sub DeplaceImage(Img As Integer)
Dim i As Integer, T As Double

For i = 1 To 26 ' Boucle pour faire une rotation de la voiture.
   ActiveSheet.Shapes(Img).IncrementRotation (i) ' Rotation.
   T = Timer: While T + 0.05 > Timer: Wend ' Pour faire une pause.
   DoEvents                                ' Actualise l'écran.
Next i
ActiveSheet.Shapes(Img).Rotation = 0  ' Remet à 0.

For i = 1 To 300 ' Boucle pour déplacer horizontalement la voiture.
   ActiveSheet.Shapes(Img).IncrementLeft 1  ' Déplace d'un pixel à droite.
   T = Timer: While T + 0.01 > Timer: Wend  ' Pour faire une pause.
   DoEvents                                 ' Actualise l'écran.
Next i

End Sub
