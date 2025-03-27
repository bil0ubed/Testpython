Sub EnvoyerFichiersEMLConfigures()
    Dim OutlookApp As Object
    Dim Explorateur As Object
    Dim dossierSource As String
    Dim fichier As String

    ' Définir le chemin du dossier contenant les fichiers .eml
    dossierSource = "C:\chemin\vers\le\dossier"

    ' Initialiser Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set Explorateur = OutlookApp.GetNamespace("MAPI")

    ' Parcourir tous les fichiers .eml dans le dossier
    fichier = Dir(dossierSource & "\*.eml")
    Do While fichier <> ""
        ' Ouvrir le fichier .eml
        Dim MailItem As Object
        Set MailItem = Explorateur.OpenSharedItem(dossierSource & "\" & fichier)

        ' Envoyer l'email
        MailItem.Send

        ' Passer au fichier suivant
        fichier = Dir
    Loop

    ' Libérer les objets
    Set MailItem = Nothing
    Set Explorateur = Nothing
    Set OutlookApp = Nothing

    MsgBox "Tous les fichiers .eml configurés ont été envoyés avec succès !"
End Sub
