VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Param�tres"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

If ComboBox1.Value = "Red" Then
        Image1.BackColor = RGB(217, 136, 137)
ElseIf ComboBox1.Value = "Orange" Then
        Image1.BackColor = RGB(241, 157, 90)
ElseIf ComboBox1.Value = "Peach" Then
        Image1.BackColor = RGB(235, 202, 103)
ElseIf ComboBox1.Value = "Yellow" Then
        Image1.BackColor = RGB(248, 242, 100)
ElseIf ComboBox1.Value = "Green" Then
        Image1.BackColor = RGB(124, 206, 110)
ElseIf ComboBox1.Value = "Teal" Then
        Image1.BackColor = RGB(117, 202, 177)
ElseIf ComboBox1.Value = "Olive" Then
        Image1.BackColor = RGB(171, 187, 141)
ElseIf ComboBox1.Value = "Blue" Then
        Image1.BackColor = RGB(116, 153, 225)
ElseIf ComboBox1.Value = "Purple" Then
        Image1.BackColor = RGB(147, 123, 209)
ElseIf ComboBox1.Value = "Maroon" Then
        Image1.BackColor = RGB(205, 141, 170)
ElseIf ComboBox1.Value = "Steel" Then
        Image1.BackColor = RGB(193, 191, 195)
ElseIf ComboBox1.Value = "Dark Steel" Then
        Image1.BackColor = RGB(83, 97, 125)
End If

End Sub


Private Sub CommandButton1_Click()


Dim pathfilelog As String
Dim logstring As String
Dim Indicateur As Boolean


pathfilelog = ReadOption("pathfilelog")

'Ouvrir le fichier journal des actions
Open pathfilelog For Append As #1

Indicateur = WriteOption("olappname", TextBox1.Value)
If Indicateur = True Then

'LOG 1
logstring = logstring & Date & " | " & " Modification du nom du tag " & " | " & TextBox1.Value & vbCrLf

End If


Indicateur = WriteOption("folder", TextBox3.Value)
If Indicateur = True Then

'LOG 2
logstring = logstring & Date & " | " & " Modification du chemin vers le dossier d'achivage " & " | " & TextBox3.Value & vbCrLf

End If


Indicateur = WriteOption("catColor", ComboBox1.ListIndex)
If Indicateur = True Then

'LOG 3
logstring = logstring & Date & " | " & " Modification du tag cat�gorie couleur " & " | " & ComboBox1.ListIndex

End If


'Ecriture du fichier
Print #1, logstring
    
'fermeture du fichier journal
Close #1

Unload Me


End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()


ComboBox1.AddItem "Red"             'ListIndex = 0      olCategoryColorRed          Color.rgb(217, 136, 137)
ComboBox1.AddItem "Orange"          'ListIndex = 1      olCategoryColorOrange       Color.rgb(241, 157, 90)
ComboBox1.AddItem "Peach"           'ListIndex = 2      olCategoryColorPeach        Color.rgb(235, 202, 103)
ComboBox1.AddItem "Yellow"          'ListIndex = 3      olCategoryColorYellow       Color.rgb(248, 242, 100)
ComboBox1.AddItem "Green"           'ListIndex = 4      olCategoryColorGreen        Color.rgb(124, 206, 110)
ComboBox1.AddItem "Teal"            'ListIndex = 5      olCategoryColorTeal         Color.rgb(117, 202, 177)
ComboBox1.AddItem "Olive"           'ListIndex = 6      olCategoryColorOlive        Color.rgb(171, 187, 141)
ComboBox1.AddItem "Blue"            'ListIndex = 7      olCategoryColorBlue         Color.rgb(116, 153, 225)
ComboBox1.AddItem "Purple"          'ListIndex = 8      olCategoryColorPurple       Color.rgb(147, 123, 209)
ComboBox1.AddItem "Maroon"          'ListIndex = 9      olCategoryColorMaroon       Color.rgb(205, 141, 170)
ComboBox1.AddItem "Steel"           'ListIndex = 10     olCategoryColorSteel        Color.rgb(193, 191, 195)
ComboBox1.AddItem "Dark Steel"      'ListIndex = 11     olCategoryColorDarkSteel    Color.rgb(83, 97, 125)


TextBox1.Value = ReadOption("olappname")
TextBox3.Value = ReadOption("folder")
ComboBox1.ListIndex = ReadOption("catColor")

Select Case ComboBox1.ListIndex
    Case 0
        Image1.BackColor = RGB(217, 136, 137)
    Case 1
        Image1.BackColor = RGB(241, 157, 90)
    Case 2
        Image1.BackColor = RGB(235, 202, 103)
    Case 3
        Image1.BackColor = RGB(248, 242, 100)
    Case 4
        Image1.BackColor = RGB(124, 206, 110)
    Case 5
        Image1.BackColor = RGB(117, 202, 177)
    Case 6
        Image1.BackColor = RGB(171, 187, 141)
    Case 7
        Image1.BackColor = RGB(116, 153, 225)
    Case 8
        Image1.BackColor = RGB(147, 123, 209)
    Case 9
        Image1.BackColor = RGB(205, 141, 170)
    Case 10
        Image1.BackColor = RGB(193, 191, 195)
    Case 11
        Image1.BackColor = RGB(83, 97, 125)
    Case Else
        Image1.BackColor = &H80000005
End Select


Image3.Visible = False
 
End Sub
 
Private Sub CommandButton4_Click()
    Dim objFSO As Object
    Dim objShell As Object
    Dim objFolder As Object
    Dim strFolderPath As String
    Dim blnIsEnd As Boolean

    blnIsEnd = False

    Set objShell = CreateObject("Shell.Application")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objShell.BrowseForFolder( _
                lHwnd, "Please Select Folder to:", _
                BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN, CSIDL_DESKTOP)


    If objFolder Is Nothing Then
        strFolderPath = ""
        blnIsEnd = True
        GoTo PROC_EXIT
    Else
        strFolderPath = CGPath(objFolder.Self.Path)
    End If
       
TextBox3.Value = strFolderPath
Image3.Visible = True

PROC_EXIT:
    Set objFSO = Nothing
    If blnIsEnd Then End
End Sub
