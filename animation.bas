' ROTATION ET DEPLACEMENT IMAGE 

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
'Lancer l'animation par:
Call DéplaceImage(i) ' où i est le numéro de l'image (vaut 1 pour la première image créée).



' BARRE DE PROGRESSION

' Sur le même principe, on peut  faire une barre de progression originale.
' Ajoutez un formulaire, nommé « BarreProgression » qui contient l'image de votre choix:
' L'affichage du formulaire se fait avec:
Call BarreProgression.Show(False) 'où False rend le formulaire non modal, 
' Le code garde ainsi la main. La progression se fait avec les instructions suivantes :
BarreProgression.Image1.Left = (BarreProgression.Width - BarreProgression.Image1.Width) * x
BarreProgression.Repaint
' Où x est un pourcentage entre 0 et 1, et l'image est nommée « Image1 ».
' À la fin du traitement, fermez la barre de progression avec l'instruction 
Unload BarreProgression.
