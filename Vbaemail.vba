Sub EnvoyerEmailAvecHTML()
    Dim OutlookApp As Object
    Dim Courrier As Object

    ' Création d'une instance Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set Courrier = OutlookApp.CreateItem(0)

    With Courrier
        .To = "destinataire@exemple.com" ' Remplacez par l'adresse e-mail du destinataire
        .Subject = "E-mail avec contenu HTML" ' Modifiez l'objet de l'e-mail
        ' Insérer du contenu HTML dans le corps de l'e-mail
        .HTMLBody = "<html><body>" & _
                    "<h1 style='color:blue;'>Bonjour !</h1>" & _
                    "<p>Ceci est un e-mail avec du contenu <b>HTML</b>.</p>" & _
                    "<p><a href='https://example.com'>Cliquez ici</a> pour en savoir plus.</p>" & _
                    "</body></html>"
        .Send ' Envoyer l'e-mail
    End With

    ' Libérer les objets
    Set Courrier = Nothing
    Set OutlookApp = Nothing
End Sub
