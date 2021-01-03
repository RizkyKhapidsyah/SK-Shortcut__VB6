Attribute VB_Name = "FileOpenforShortcut"
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

'Function Declarations, type structure, and constants for the Open and Save
'common dialog boxes.  For more information on these, consult the SDK,
'included with VB 6.0 Pro or Ent editions as part of the MSDN/VB Starter
'Kit.

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Functions and constants for the common dialog boxes
Public Declare Function GetOpenFileName Lib "comdlg32.dll" _
Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32.dll" _
Alias "GetSaveFileNameA" (pOPENFILENAME As OPENFILENAME) As Long

'These constants must be declared since the Common Dialog control is not
'part of the project and, therefore, the intrinsic constants are not
'defined.
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000         '  force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000            '  new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000          '  force long names for 3.x modules

