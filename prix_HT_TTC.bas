Sub Pgm_Principal()
Dim NumLig As Long
    NumLig = 2
    While Cells(NumLig, 1) <> ""
        Cells(NumLig, 5) = Prix_HT_TTC(Cells(NumLig, 4), Cells(NumLig, 3))
        Cells(NumLig, 6) = Prix_HT_TTC(Cells(NumLig, 4), Cells(NumLig, 3), Cells(NumLig, 2), 0.2)
        NumLig = NumLig + 1
    Wend
End Sub

Function Prix_HT_TTC(Prix_Unitaire As Double, Qté As Long, Optional Pays As String, Optional TVA As Double)
    If Pays <> "" Then
        If Pays = "FR" Then
            Prix_HT_TTC = (Prix_Unitaire * Qté) * (1 + TVA)
        Else
            Prix_HT_TTC = "-"
        End If
    Else
        Prix_HT_TTC = Prix_Unitaire * Qté
    End If
End Function
