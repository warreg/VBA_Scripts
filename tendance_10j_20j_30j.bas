Sub Pgm_Principal()
Dim NumLig As Long
Dim NumCol As Long
Dim EnteteCol As Long
    For NumCol = 4 To 6
        NumLig = 2
        EnteteCol = Cells(1, NumCol) 'Valeur des entêtes
        While Cells(NumLig, 1) <> ""
           If NumLig > EnteteCol Then
                Cells(NumLig, NumCol) = _
                  La_Moyenne(Range(Cells(NumLig - EnteteCol + 1, 3), Cells(NumLig, 3)))
            End If
        NumLig = NumLig + 1
        Wend
    Next
End Sub

Function La_Somme(Tab_Valeur As Variant) As Double
Dim Indice_Tab As Long
   For Indice_Tab = 1 To UBound(Tab_Valeur.Value)
        La_Somme = La_Somme + Tab_Valeur(Indice_Tab)
    Next
End Function
Function La_Moyenne(Tab_Valeur As Variant) As Double
     La_Moyenne = La_Somme(Tab_Valeur) / UBound(Tab_Valeur.Value)
End Function
