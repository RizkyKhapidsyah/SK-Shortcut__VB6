Attribute VB_Name = "BrowseForShortcutFolder"
' Disclaimer of Warranty:

' This software and the accompanying files are provided "as is"
' and without warranties as to performance of the software and
' the accompanying files or any other warranties whether expressed
' or implied.  No warranty of fitness for a particular purpose
' is offered.
'
' You MAY NOT sell this software or it's source code.
' You MAY use this code in any way you find useful.

Option Explicit

'Function Declarations, type structure, and constants to use the
'Browse for Folder dialog box.  For more information on these,
'consult the SDK, included with VB 6.0 Pro or Ent editions as
'part of the MSDN/VB Starter Kit.

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

'Below are the constants which can be specified in the ulFlags member
'of the BROWSEINFO structure.

'Only returns file system directories. If the user selects folders
'that are not part of the file system, the OK button is grayed.
'Eg... My Computer
Public Const BIF_RETURNONLYFSDIRS = &H1

'Does not include network folders below the domain level in the
'tree view control.
Public Const BIF_DONTGOBELOWDOMAIN = &H2

'Only returns file system ancestors. If the user selects anything
'other than a file system ancestor, the OK button is grayed.
Public Const BIF_RETURNFSANCESTORS = &H8

'Only returns computers. If the user selects anything other than
'a computer, the OK button is grayed.
Public Const BIF_BROWSEFORCOMPUTER = &H1000

'Only returns printers. If the user selects anything other than
'a printer, the OK button is grayed.
Public Const BIF_BROWSEFORPRINTER = &H2000

'Includes a status area in the dialog box. The callback function
'can set the status text by sending messages to the dialog box.
Const BIF_STATUSTEXT = &H4


