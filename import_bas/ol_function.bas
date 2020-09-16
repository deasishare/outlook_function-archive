Attribute VB_Name = "ol_function"
Option Explicit

Public Const file_ini As String = "c:\ol\config.ini"

' Déclarations pour fichier ini
#If VBA7 Then
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias _
        "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias _
        "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#Else
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
        "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
        "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#End If

'------------------------------------------------------------------------
' Lecture de la valeur d'une clé dans un fichier ini
'------------------------------------------------------------------------
Public Function ReadOption(pOption As String, Optional pDefault As String) As String
Dim lRet As Long
Dim lData As String
Dim lSize As Long
lData = Space$(8192)
lSize = 8192
lRet = GetPrivateProfileString("Options", pOption, pDefault, lData, lSize, file_ini)
If lSize > 0 Then
    ReadOption = Left$(lData, lRet)
Else
    ReadOption = ""
End If
End Function
            
'------------------------------------------------------------------------
' Ecriture de la valeur d'une clé dans un fichier ini
'------------------------------------------------------------------------
Public Function WriteOption(pOption As String, Optional pValue As String) As Boolean
Dim lRet As Long
Dim lSize As Long
If pValue = "" Then
    lRet = WritePrivateProfileString("Options", pOption, 0&, file_ini)
Else
    lRet = WritePrivateProfileString("Options", pOption, pValue, file_ini)
End If
WriteOption = (lRet <> 0)
End Function

Function checkCategory(strCtName As String, intColor As Integer, intKey As Integer)

    'Déclaration variable interne
    Dim objNS As NameSpace
    Dim objCat As Category
    Dim bolCat As Boolean
     
    'Instance
    Set objNS = Application.GetNamespace("MAPI")
    
    If objNS.Categories.Count > 0 Then
        
        For Each objCat In objNS.Categories

            If objCat.Name = strCtName Then
                  bolCat = True 'la catégorie existe
                  objCat.color = intColor 'redefini la couleur
            End If
        Next
        
    End If
    
    'On ajoute si elle n'existe pas
    If bolCat = False Then
        AddCategory strCtName, intColor, intKey
    End If

    
    'On libere les instances
    Set objCat = Nothing
    Set objNS = Nothing

End Function

Function AddCategory(strCategoryName As String, intColor As Integer, intKey As Integer)
    
    Dim objNS As NameSpace
 
    Set objNS = Application.GetNamespace("MAPI")
    On Error Resume Next
    
    objNS.Categories.Add strCategoryName, intColor, intKey
    
    'On libere les instances
    Set objNS = Nothing
    
End Function
 

Public Function id_aleatoire(carac As String) As String

    Randomize
 
    Dim code_alea As String
    Dim nombre_aleatoire As Integer
    Dim i
    
    For i = 1 To 5 '10 = longueur du code
        nombre_aleatoire = Int(Len(carac) * Rnd) + 1
        code_alea = code_alea + Mid(carac, nombre_aleatoire, 1)
    Next
    
    code_alea = code_alea & "_" & Fix(Timer)

    id_aleatoire = code_alea
    
End Function


Public Function OuvrirFichier(MonFichier As String)
   
On Error GoTo OuvertureFichierErreur
   
   'vérifie si le fichier existe
   If Len(Dir(MonFichier)) = 0 Then
    OuvrirFichier = False
    Exit Function
   Else
   End If
   
   'ouvre le fichier dans son application associée
   Dim MonApplication As Object
   Set MonApplication = CreateObject("Shell.Application")
   
    MonApplication.Open (MonFichier)
    OuvrirFichier = True
   Set MonApplication = Nothing
   
Exit Function
OuvertureFichierErreur:
   Set MonApplication = Nothing
    OuvrirFichier = False
    
End Function




