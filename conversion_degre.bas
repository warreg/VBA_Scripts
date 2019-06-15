Sub Conversion_Degrés()
Dim Num_Ligne As Long
    Num_Ligne = 2
    While Cells(Num_Ligne, 1) <> ""
        Cells(Num_Ligne, 3) = Fahrenheit(Cells(Num_Ligne, 2))
        Num_Ligne = Num_Ligne + 1
    Wend
End Sub

Function Fahrenheit(Valeur_Degrés As Double)
    Fahrenheit = Valeur_Degrés * 9 / 5 + 32
End Function
