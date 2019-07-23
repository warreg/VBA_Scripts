Option Explicit

Dim ProgressIndicator As UserForm1


Sub Start()
'   la sub UserForm1_Activate sub fait appel
    UserForm1.LabelProgress.Width = 0
    UserForm1.Show
End Sub

Sub EnterRandomNumbers()
'   Insère des nombres aléatoires dans la feuille active
    Dim Counter As Integer
    Dim RowMax As Integer, ColMax As Integer
    Dim r As Integer, c As Integer
    Dim PctDone As Single
    
'   Crée une copie du forme dans une variable
    Set ProgressIndicator = New UserForm1
    
'   Affiche Barreprogression dans état modeless
    ProgressIndicator.Show vbModeless
    If TypeName(ActiveSheet) <> "Worksheet" Then
        Unload ProgressIndicator
        Exit Sub
    End If
    Cells.Clear
    Counter = 1
    RowMax = 650 'rajout yvouille : à la base, il y avait 200, max 650
    ColMax = 50
    For r = 1 To RowMax
        For c = 1 To ColMax
            Cells(r, c) = Int(Rnd * 1000)
            Counter = Counter + 1
        Next c
        PctDone = Counter / (RowMax * ColMax)
        Call UpdateProgress(PctDone)
    Next r
    Unload ProgressIndicator
    Set ProgressIndicator = Nothing
End Sub


Sub UpdateProgress(pct)
    With ProgressIndicator
        .FrameProgress.Caption = Format(pct, "0%")
        .LabelProgress.Width = pct * (.FrameProgress _
           .Width - 10)
    End With
'   L'instruction DoEvents est responsable fde la mise à jour de l'Userform
    DoEvents
End Sub

