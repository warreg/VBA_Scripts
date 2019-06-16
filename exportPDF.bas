' Pour plus d'exemples sur la création de fichier en PDF, visitez le site de Ron de Bruin
' http://www.rondebruin.nl/tips.htm

Sub RDB_Workbook_To_PDF()
    Dim FileName As String

    'Appel de la function avec tous 4 arguments attendus.
    FileName = RDB_Create_PDF(ActiveWorkbook, "", True, True)

    'Si le fichier doit être écrasé à chaque fois que vous lancez le traitement
    'enlevez le commentaire de la ligne ci-dessous.
    'RDB_Create_PDF(ActiveWorkbook, " C:\MesDocuments\FichierPDF ", True, True)

    If FileName <> "" Then
    'Enlevez le commentaire suivant si vous voulez envoyer par courriel le fichier pdf.
        'RDB_Mail_PDF_Outlooka FileName, "destin@tai.re", "Votre sujet", _
           "Texte de votre message", False
    Else
        MsgBox "Il n'est pas possible de créer de PDF pour l'une des raisons suivantes :" & vbNewLine & _
               "La dll n'est pas installée" & vbNewLine & _
               "Vous avez annulé l'Enregistrer sous" & vbNewLine & _
               "Le chemin indiqué n'est pas correct" & vbNewLine & _
               "Vous ne voulez pas écraser le PDF existant."
    End If
End Sub

Function RDB_Create_PDF(Myvar As Object, FixedFilePathName As String, _
         OverwriteIfFileExist As Boolean, OpenPDFAfterPublish As Boolean) As String
    Dim FileFormatstr As String
    Dim Fname As Variant

    'Test pour vérifier si la dll est installée
    If Dir(Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" _
         & Format(Val(Application.Version), "00") & "\EXP_PDF.DLL") <> "" Then

        If FixedFilePathName = "" Then
            'Ouverture de la boîte Enregistrer sous et entrer le nom du fichier PDF
            FileFormatstr = "Fichiers PDF (*.pdf), *.pdf"
            Fname = Application.GetSaveAsFilename("", filefilter:=FileFormatstr, _
                  Title:="Create PDF")

            'Si vous annulez cette boîte, vous quittez la fonction
            If Fname = False Then Exit Function
        Else
            Fname = FixedFilePathName
        End If

        'Si le paramètre OverwriteIfFileExist = Faux, alors un test est réalisé
        'pour savoir si le fichier n'existe pas déjà dans le répertoire
        If OverwriteIfFileExist = False Then
            If Dir(Fname) <> "" Then Exit Function
        End If

        'Maintenant Exportation du fichier PDF.
        On Error Resume Next
        Myvar.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                FileName:=Fname, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=OpenPDFAfterPublish
        On Error GoTo 0

        'Si l'export est réussi, le nom du fichier PDF est renvoyé à la fonction.
        If Dir(Fname) <> "" Then RDB_Create_PDF = Fname
    End If
End Function


