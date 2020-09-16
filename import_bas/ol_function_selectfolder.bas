Attribute VB_Name = "ol_function_selectfolder"
Option Explicit
' For Outlook 2010.

#If VBA7 Then
    ' The window handle of Outlook.
    Private lHwnd As LongPtr

    ' /* API declarations. */
    Private Declare PtrSafe Function FindWindow Lib "user32" _
            Alias "FindWindowA" (ByVal lpClassName As String, _
                                 ByVal lpWindowName As String) As LongPtr


' For the previous version of Outlook 2010.
#Else
    ' The window handle of Outlook.
    Private lHwnd As Long

    ' /* API declarations. */
    Private Declare Function FindWindow Lib "user32" _
            Alias "FindWindowA" (ByVal lpClassName As String, _
                                 ByVal lpWindowName As String) As Long
#End If
'
' Windows desktop -
' the virtual folder that is the root of the namespace.
Private Const CSIDL_DESKTOP = &H0

' Only return file system directories.
' If user selects folders that are not part of the file system,
' then OK button is grayed.
Private Const BIF_RETURNONLYFSDIRS = &H1

' Do not include network folders below
' the domain level in the dialog box's tree view control.
Private Const BIF_DONTGOBELOWDOMAIN = &H2

Public Function CGPath(ByVal Path As String) As String
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    CGPath = Path
End Function









