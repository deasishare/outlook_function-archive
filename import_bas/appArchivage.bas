Attribute VB_Name = "appArchivage"
'====================================================================================================
' Proc�dure : Archivages des pieces jointes et des images des emails
' Auteur    : Didier Maes
' Date      : 19 avril 2020
' D�tail    : Fonction d'archivage des pieces jointes avec un lien dans le message pour les ouvrir
'====================================================================================================

' D�tail de la proc�dure :
' ======================

' - Traiter les messages selectionner
' - Inscrire � la fin du message {pi�ce jointe enlev�e:} et indiquer le chemin d'acc�s au fichier complet (avec hyperlink)
' - Copier les pi�ces jointes dans un dossier local {c:\ol_archivage\}
' - Supprimer les pi�ces jointes du message de l'email re�u
' - Etablir une cat�gorie et l'attribuer au message : couleur = {olCategoryColorOlive} nom = {�volution_pr�cipit�es}
' - Alimenter un fichier {journal.log} si situant dans le dossier {c:\ol_archivage\}


' - Restoration du message.


'--------------------------------------------------------------------------------------------------

Sub appArchivagesPieceJointe()

    'D�claration
    Dim myOlApp As New Outlook.Application
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim myItems, myItem, myAttachments, myAttachment As Object 'pi�ces jointes


    Dim idmessage As String
    Dim myOrt As String 'chemin de sauvegarde
    Dim MsgTxt As String 'contenu du message
    Dim X As Integer 'boucle
    Dim ItemBody As String
    Dim extensionpj As String


Dim pathfilelog As String
Dim logstring As String

pathfilelog = ReadOption("pathfilelog")


Dim olappname As String  ' Archiv�s
Dim catColor As String     ' 5

olappname = ReadOption("olappname")
catColor = ReadOption("catColor")


    'Ajouter une cat�gorie � outlook :
    checkCategory CStr(olappname), CInt(catColor), CInt(0)


'Ouvrir le fichier journal des actions
Open pathfilelog For Append As #1

    'Bo�te de dialogue simple pour le chemin de sauvegarde
    myOrt = ReadOption("folder")

    On Error Resume Next

    'Actions sur les objets s�lectionn�s
    Set myOlExp = myOlApp.ActiveExplorer
    Set myOlSel = myOlExp.Selection 'ensemble des myItem

    'boucle
    For Each myItem In myOlSel

        Set myAttachments = myItem.Attachments
        If myAttachments.Count > 0 Then
        
        'Ajouter une cat�gorie
        myItem.Categories = olappname
        
        ItemBody = "<HTML><BODY><p>" & myItem.HTMLBody & "</p><br><br><hr>"
            
            'Ajoute une remarque dans le corps du message
            
            myItem.BodyFormat = olFormatHTML
            ItemBody = ItemBody & "<table style=""" & "border-collapse: collapse; border: 2px solid rgb(200, 200, 200);letter-spacing: 1px;" & "><tr style=""" & "background-color: #696969;" & """><th style=""" & "border: 1px solid rgb(190, 190, 190);padding: 7px;background-color: #696969;color:whitesmoke; font-family: Arial, Helvetica, sans-serif; font-size: 0.8rem; width:400px;" & """>Liste des pi�ces jointes archiv�es</th></tr>"
      
      
            'Stockage des diff�rentes informations utiles du mail
            With myItem
                DateRe = .ReceivedTime
                Expediteur = .SenderName
                Objet = .Subject
            End With
      
            'LOG 1 : Stockage des informations � utiliser pour les logs
            logstring = DateRe & " | " & Expediteur & " | " & Objet & " | "
      

      
             'pour toutes les pi�ces jointes
            For i = 1 To myAttachments.Count

            extensionpj = Mid(myAttachments(i), InStrRev(myAttachments(i), ".") + 1)
            
            If extensionpj = "png" Then

                ElseIf extensionpj = "jpg" Then
                
                Else
                'Enregistrer � destination
                idmessage = id_aleatoire(myItem.EntryID)
                myAttachments(i).SaveAsFile myOrt & idmessage & myAttachments(i).DisplayName
                
                ItemBody = ItemBody & "<tr><th style=""" & "font-weight: lighter; font-family: Arial, Helvetica, sans-serif; font-style: italic; font-size: 0.7rem;width:400px; border: 1px solid rgb(190, 190, 190);padding: 5px;" & """>" & "<a href=""" & myOrt & idmessage & myAttachments(i).DisplayName & """>" & myOrt & idmessage & myAttachments(i).DisplayName & "</a>" & "</th></tr>"
            
                'LOG 2 : Stockage �galement du nom des pi�ces jointes � supprimer dans le fichier des logs
                logstring = logstring & myOrt & idmessage & myAttachments(i).DisplayName
            End If
            Next i
          
            myItem.HTMLBody = ItemBody & "</table></p></BODY></HTML>"

                'Enl�ve les pi�ces jointes du message (boucle) Verifier si possible de le mettre dans le for next precedent
                
                For j = 1 To myAttachments.Count
                extensionpj = Mid(myAttachments(j), InStrRev(myAttachments(j), ".") + 1)
                If extensionpj = "png" Then
                ElseIf extensionpj = "jpg" Then
                Else
                myAttachments(j).Delete
                End If
                Next j

            'Sauvegarde le message sans ses pi�ces jointes
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
