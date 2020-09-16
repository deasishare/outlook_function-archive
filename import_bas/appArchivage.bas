Attribute VB_Name = "appArchivage"
'====================================================================================================
' Procédure : Archivages des pieces jointes et des images des emails
' Auteur    : Didier Maes
' Date      : 19 avril 2020
' Détail    : Fonction d'archivage des pieces jointes avec un lien dans le message pour les ouvrir
'====================================================================================================

' Détail de la procédure :
' ======================

' - Traiter les messages selectionner
' - Inscrire à la fin du message {pièce jointe enlevée:} et indiquer le chemin d'accès au fichier complet (avec hyperlink)
' - Copier les pièces jointes dans un dossier local {c:\ol_archivage\}
' - Supprimer les pièces jointes du message de l'email reçu
' - Etablir une catégorie et l'attribuer au message : couleur = {olCategoryColorOlive} nom = {Évolution_précipitées}
' - Alimenter un fichier {journal.log} si situant dans le dossier {c:\ol_archivage\}


' - Restoration du message.


'--------------------------------------------------------------------------------------------------

Sub appArchivagesPieceJointe()

    'Déclaration
    Dim myOlApp As New Outlook.Application
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim myItems, myItem, myAttachments, myAttachment As Object 'pièces jointes


    Dim idmessage As String
    Dim myOrt As String 'chemin de sauvegarde
    Dim MsgTxt As String 'contenu du message
    Dim X As Integer 'boucle
    Dim ItemBody As String
    Dim extensionpj As String


Dim pathfilelog As String
Dim logstring As String

pathfilelog = ReadOption("pathfilelog")


Dim olappname As String  ' Archivés
Dim catColor As String     ' 5

olappname = ReadOption("olappname")
catColor = ReadOption("catColor")


    'Ajouter une catégorie à outlook :
    checkCategory CStr(olappname), CInt(catColor), CInt(0)


'Ouvrir le fichier journal des actions
Open pathfilelog For Append As #1

    'Boîte de dialogue simple pour le chemin de sauvegarde
    myOrt = ReadOption("folder")

    On Error Resume Next

    'Actions sur les objets sélectionnés
    Set myOlExp = myOlApp.ActiveExplorer
    Set myOlSel = myOlExp.Selection 'ensemble des myItem

    'boucle
    For Each myItem In myOlSel

        Set myAttachments = myItem.Attachments
        If myAttachments.Count > 0 Then
        
        'Ajouter une catégorie
        myItem.Categories = olappname
        
        ItemBody = "<HTML><BODY><p>" & myItem.HTMLBody & "</p><br><br><hr>"
            
            'Ajoute une remarque dans le corps du message
            
            myItem.BodyFormat = olFormatHTML
            ItemBody = ItemBody & "<table style=""" & "border-collapse: collapse; border: 2px solid rgb(200, 200, 200);letter-spacing: 1px;" & "><tr style=""" & "background-color: #696969;" & """><th style=""" & "border: 1px solid rgb(190, 190, 190);padding: 7px;background-color: #696969;color:whitesmoke; font-family: Arial, Helvetica, sans-serif; font-size: 0.8rem; width:400px;" & """>Liste des pièces jointes archivées</th></tr>"
      
      
            'Stockage des différentes informations utiles du mail
            With myItem
                DateRe = .ReceivedTime
                Expediteur = .SenderName
                Objet = .Subject
            End With
      
            'LOG 1 : Stockage des informations à utiliser pour les logs
            logstring = DateRe & " | " & Expediteur & " | " & Objet & " | "
      

      
             'pour toutes les pièces jointes
            For i = 1 To myAttachments.Count

            extensionpj = Mid(myAttachments(i), InStrRev(myAttachments(i), ".") + 1)
            
            If extensionpj = "png" Then

                ElseIf extensionpj = "jpg" Then
                
                Else
                'Enregistrer à destination
                idmessage = id_aleatoire(myItem.EntryID)
                myAttachments(i).SaveAsFile myOrt & idmessage & myAttachments(i).DisplayName
                
                ItemBody = ItemBody & "<tr><th style=""" & "font-weight: lighter; font-family: Arial, Helvetica, sans-serif; font-style: italic; font-size: 0.7rem;width:400px; border: 1px solid rgb(190, 190, 190);padding: 5px;" & """>" & "<a href=""" & myOrt & idmessage & myAttachments(i).DisplayName & """>" & myOrt & idmessage & myAttachments(i).DisplayName & "</a>" & "</th></tr>"
            
                'LOG 2 : Stockage également du nom des pièces jointes à supprimer dans le fichier des logs
                logstring = logstring & myOrt & idmessage & myAttachments(i).DisplayName
            End If
            Next i
          
            myItem.HTMLBody = ItemBody & "</table></p></BODY></HTML>"

                'Enlève les pièces jointes du message (boucle) Verifier si possible de le mettre dans le for next precedent
                
                For j = 1 To myAttachments.Count
                extensionpj = Mid(myAttachments(j), InStrRev(myAttachments(j), ".") + 1)
                If extensionpj = "png" Then
                ElseIf extensionpj = "jpg" Then
                Else
                myAttachments(j).Delete
                End If
                Next j

            'Sauvegarde le message sans ses pièces jointes
            myItem.Save
        End If

    Next

'Ecriture du fichier
Print #1, logstring
    
'fermeture du fichier journal
Close #1



    'On libere les instances

    Set myInbox = Nothing
    Set myItems = Nothing
    Set myDestFolder = Nothing
    Set myRestrictItems = Nothing
    Set myOlExp = Nothing
    Set myOlApp = Nothing
    Set myOlSel = Nothing
    Set myItems = Nothing
    Set myItem = Nothing
    Set myAttachments = Nothing
    Set myAttachment = Nothing


End Sub

Sub parametreshow()
UserForm1.Show
End Sub

Sub ouvrirlog()

Dim pathfilelog As String

pathfilelog = ReadOption("pathfilelog")

'ceci va lancer le fichier
OuvrirFichier (pathfilelog)

End Sub
