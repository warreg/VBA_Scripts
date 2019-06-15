Sub Appel_Fonction()
    Cells(12, 1) = Ma_Somme(Range("A1:A10"))
End Sub
Function Ma_Somme(Tab_Cell As Object)
Dim Indice_Tab As Long
   For Indice_Tab = 1 To UBound(Tab_Cell.Value)
        Ma_Somme = Ma_Somme + Tab_Cell(Indice_Tab)
    Next
End Function

