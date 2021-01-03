VERSION 5.00
Begin VB.Form CreateDeleteShortcut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create & Delete any File Shortcut:"
   ClientHeight    =   4545
   ClientLeft      =   3765
   ClientTop       =   1755
   ClientWidth     =   5625
   Icon            =   "Create Shortcut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtUrlname 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   4890
      End
      Begin VB.TextBox txtShortcutName 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   4890
      End
      Begin VB.TextBox txtExename 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   4890
      End
      Begin VB.TextBox txtShortcutDir 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   4890
      End
      Begin VB.CommandButton cmdShortcutPathDialog 
         Caption         =   "..."
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton cmdExeDialog 
         Caption         =   "..."
         Height          =   255
         Left            =   5040
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblUrlname 
         Caption         =   "Path for your Internet Shortcut. Eg... http://www.myshortcut.com"
         Height          =   285
         Left            =   165
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   4770
      End
      Begin VB.Label lblShortcutName 
         Alignment       =   2  'Center
         Caption         =   "What is the name for this shortcut? Eg. Notepad"
         Height          =   285
         Left            =   165
         TabIndex        =   23
         Top             =   720
         Width           =   4770
      End
      Begin VB.Label lblExename 
         Alignment       =   2  'Center
         Caption         =   "Path of file you want for Shortcut. Eg... C:\Windows\Notepad.exe"
         Height          =   285
         Left            =   165
         TabIndex        =   22
         Top             =   120
         Width           =   4770
      End
      Begin VB.Label lblShortcutDir 
         Alignment       =   2  'Center
         Caption         =   "What directory should the shortcut be created in? Please select one."
         Height          =   270
         Left            =   165
         TabIndex        =   21
         Top             =   1305
         Width           =   4890
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quick Access"
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5415
      Begin VB.CheckBox chkQuickLaunchBar 
         Caption         =   "Quick Launch Bar"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkFavorites 
         Alignment       =   1  'Right Justify
         Caption         =   "&Favorites"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkAccessories 
         Alignment       =   1  'Right Justify
         Caption         =   "&Accessories"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkStartup 
         Alignment       =   1  'Right Justify
         Caption         =   "S&tartup"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkPrograms 
         Caption         =   "&Programs"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkStartmenu 
         Caption         =   "&Start Menu"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkDesktop 
         Caption         =   "&Desktop"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frmOption 
      Caption         =   "                            Select type of Shortcut you want created"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   5415
      Begin VB.CheckBox chkUrlname 
         Alignment       =   1  'Right Justify
         Caption         =   "&Internet Shortcut (.url)"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   320
         Width           =   1935
      End
      Begin VB.CheckBox chkLnkname 
         Caption         =   "&Normal Shortcut (.lnk .pif)"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   320
         Width           =   2175
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   100
         X2              =   1365
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   120
         X2              =   1360
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   2805
         X2              =   2805
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   2805
         X2              =   2805
         Y1              =   240
         Y2              =   600
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Text            =   "5"
      Top             =   5400
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   5280
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete any Shortcut"
      Height          =   375
      Left            =   1965
      Picture         =   "Create Shortcut.frx":0442
      TabIndex        =   2
      Top             =   4080
      Width           =   1650
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3880
      Picture         =   "Create Shortcut.frx":074C
      TabIndex        =   1
      Top             =   4080
      Width           =   1650
   End
   Begin VB.CommandButton cmdCreateShortcut 
      Caption         =   "&Create Shortcut"
      Height          =   375
      Left            =   120
      Picture         =   "Create Shortcut.frx":0A56
      TabIndex        =   0
      Top             =   4080
      Width           =   1650
   End
End
Attribute VB_Name = "CreateDeleteShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'Functions from Setup Toolkit for creating links.  I'm not aware of any
'official documentation for these, but you can examine the setup program
'from VB's Setup Toolkit for additional examples.

'#########################################################################
'#  Could someone let me know how to add Startin, Shortcut Key & Run     #
'#  Eg...                                                                #
'#                                                                       #
'#  1: LNK or URL Name is txtlnkname - txturlnamename                    #
'#  2: Target which is txtExename                                        #
'#  3: Startin        Eg... Prog in in C:\Windows starup is in D:\       #
'#  4: Shortcut Key   Eg... CTRL+ALT+S                                   #
'#  5: Run            Eg... Min, Max                                     #
'#                                                                       #
'#  Please send E-Mail to bryce3@bigpond.com                             #
'#########################################################################

Private Declare Function fCreateShellGroup Lib "STKIT432.DLL" _
(ByVal lpstrDirName As String) As Long

Private Declare Function fCreateShellLink Lib "STKIT432.DLL" _
(ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, _
ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long

Private Declare Function fRemoveShellLink Lib "STKIT432.DLL" _
(ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long

'For more information on the functions, structure, and constants below,
'consult the SDK, included with VB 6.0 Pro and Ent editions as
'part of the MSDN/VB Starter Kit.

'Function to get Windows directory
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Structure used by SHFileOperation
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type
  
'Function and constants for performing operations on file system objects
Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'Tip: This same function can be used to delete a file to Win95's
'Recycle Bin. Just specify FO_DELETE for wFunc and FOF_ALLOWUNDO
'for fFlags.

Const FO_DELETE = &H3
Const FO_COPY = &H2
Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const FOF_CONFIRMMOUSE = &H2
Const FOF_FILESONLY = &H80
Const FOF_MULTIDESTFILES = &H1
Const FOF_NOCONFIRMATION = &H10
Const FOF_NOCONFIRMMKDIR = &H200
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_SILENT = &H4
Const FOF_SIMPLEPROGRESS = &H100
Const FOF_WANTMAPPINGHANDLE = &H20
Const FOF_ALLOWUNDO = &H40

'Variable to store the Windows directory
Dim msWinDir As String

'Buffer and constant used for API functions
Dim msBuffer As String * 255
Const BUFFERSIZE = 255

Private Sub CreateShortcut(ByVal sDir As String, ByVal sShortcut As String, ByVal sExecutable As String, ByVal sArguments As String)

'This sub creates a shortcut file (*.lnk or *.pif) in any directory.
'sDir: directory in which to create the shortcut
'sShortcut: file name of the shortcut
'sExecutable:  file name of the executable file to which the the shortcut points
'sArguments:  any optional command line arguments for the executable file

Dim lRet As Long
Dim udtFileOp As SHFILEOPSTRUCT

'Make sure there is no extension.  Since periods are valid in filenames, check
'for the specific extension, rather than just for the existance of a period.
If chkLnkname.Value = 1 Then
If LCase$(Left$(sShortcut, 4)) = ".lnk" Or LCase$(Left$(sShortcut, 4)) = ".pif" Then
    sShortcut = Left$(sShortcut, Len(sShortcut) - 4)
End If
End If
If chkUrlname.Value = 1 Then
If LCase$(Left$(txtShortcutDir & sShortcut, 4)) = ".url" Then
    sShortcut = Left$(txtShortcutDir & sShortcut, Len(txtShortcutDir & sShortcut) - 4)
End If
End If
    
'This function creates a shortcut in the \Start Menu\Programs folder by default.
'The function returns 0 if NOT successful.
'For the first argument, you can specify a relative path to the \Start Menu\Programs
'folder in which to create the shortcut.  For example, if you have
'"..\..Desktop", the shortcut will be created in the Windows\Desktop folder.
'Unfortunately, this can get real messy if you're allowing the user to specify
'the folder for the shortcut because the path MUST be relative to
'\Start Menu\Programs.
If chkLnkname.Value = 1 Then
lRet = fCreateShellLink("", sShortcut, sExecutable, sArguments)
ElseIf chkUrlname.Value = 1 Then
CreateInternetShortCut txtShortcutName & ".url", txtUrlname
End If
If lRet Then
    'If successful, move the shortcut to the specified directory
    udtFileOp.hwnd = Me.hwnd
    udtFileOp.wFunc = FO_MOVE
    
    'Need to determine if the shortcut is to a Windows or DOS program.
    'This can be done just by checking whether the file has an LNK or PIF extension
    If LCase$(Right$(Dir$(msWinDir & "\start menu\programs\" & sShortcut & ".*"), 3)) = "lnk" Then
        udtFileOp.pFrom = msWinDir & "\start menu\programs\" & sShortcut & ".lnk"
        udtFileOp.pTo = sDir & "\" & sShortcut & ".lnk"
    Else
        'Not an LNK extension, so it must be PIF
        udtFileOp.pFrom = msWinDir & "\start menu\programs\" & sShortcut & ".pif"
        udtFileOp.pTo = sDir & "\" & sShortcut & ".pif"
    End If
    
    'udtFileOp.fFlags = FOF_NOCONFIRMATION
    'If you uncomment the above line, "yes to all" will be assumed for any
    'prompts, such as one to confirm a file replace.
    
    udtFileOp.fFlags = FOF_RENAMEONCOLLISION
    'Uncomment the above line if you want the file automatically
    'renamed if a file of the same name already exists.
    
    'This function returns 0 if it IS successful.
    lRet = SHFileOperation(udtFileOp)
    
    'Do not display the message box if the user chose to abort.
    If lRet <> 0 And Not udtFileOp.fAnyOperationsAborted Then
       ' MsgBox "Unable to create shortcut for " & sExecutable & ".", vbInformation
    End If
Else
    'MsgBox "Unable to create shortcut for " & sExecutable & ".", vbInformation
End If

End Sub




Private Function ShowOpen(ByRef sFile As String, ByRef sFileTitle As String) As Boolean

Dim udtFile As OPENFILENAME
Dim lResult As Long
Dim nNullPos As Integer

'The function needs to know the size of the structure being passed to it
udtFile.lStructSize = Len(udtFile)

'Specify the window handle for the owner of the dialog box
udtFile.hwndOwner = Me.hwnd

'Set desired flags
udtFile.Flags = OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST + OFN_HIDEREADONLY

'Selected filename is filled in this element so
'create a long enough string to be the buffer.
udtFile.lpstrFile = Space$(255)
'Set the maximum size for the returned path\filename
udtFile.nMaxFile = 255

'Buffer to be filled with the filename only
udtFile.lpstrFileTitle = Space$(255)
'Specify the size of this buffer
udtFile.nMaxFileTitle = 255

'set the default directory; if not set, current directory is used
udtFile.lpstrInitialDir = msWinDir

'specify a filter
If chkLnkname.Value = 1 Then
udtFile.lpstrFilter = "Application Files" & vbNullChar & "*.EXE" & vbNullChar & "MS-DOS Application" & vbNullChar & "*.COM" & vbNullChar & "EXE & COM" & vbNullChar & "*.EXE;*.COM" & vbNullChar & "MS-DOS Batch File" & vbNullChar & "*.BAT" & vbNullChar & "Any File" & vbNullChar & "*.*"
udtFile.nFilterIndex = 1
ElseIf chkUrlname.Value = 1 Then
udtFile.lpstrFilter = "Internet Files" & vbNullChar & "*.URL" & vbNullChar & "Any File" & vbNullChar & "*.*"
udtFile.nFilterIndex = 1
End If
'dialog box caption
udtFile.lpstrTitle = "Open"

'Call the function; a non-zero value is returned if successful
lResult = GetOpenFileName(udtFile)

If lResult <> 0 Then
    'Find position of null character
    nNullPos = InStr(udtFile.lpstrFileTitle, vbNullChar)
    If nNullPos > 0 Then
        sFileTitle = Left$(udtFile.lpstrFileTitle, nNullPos - 1)
    End If
    
    nNullPos = InStr(udtFile.lpstrFile, vbNullChar)
    If nNullPos > 0 Then
        sFile = Left$(udtFile.lpstrFile, nNullPos - 1)
    End If
    
    ShowOpen = True
    
    'nFileOffset is the number of characters from the beginning of the
    'full path to the start of the file name
    
    'nFileExtension is the number of characters from the beginning of the
    'full path to the file's extension, including the (.)
End If

End Function
'I Know there is an easier way, but i have done it this way
'to save the hassels of changing any codes

Private Sub chkDesktop_Click()
If chkDesktop.Value = 1 Then
chkQuickLaunchBar.Value = 0
chkStartmenu.Value = 0
chkPrograms.Value = 0
chkStartup.Value = 0
chkAccessories.Value = 0
chkFavorites.Value = 0
txtShortcutDir = msWinDir & "\Desktop"
End If
End Sub

Private Sub chkQuickLaunchBar_Click()
If chkQuickLaunchBar.Value = 1 Then
chkDesktop.Value = 0
chkPrograms.Value = 0
chkStartup.Value = 0
chkAccessories.Value = 0
chkFavorites.Value = 0
txtShortcutDir = msWinDir & "\Application Data\Microsoft\Internet Explorer\Quick Launch"
End If
End Sub

Private Sub chkStartmenu_Click()
If chkStartmenu.Value = 1 Then
chkQuickLaunchBar.Value = 0
chkStartmenu.Value = 0
chkDesktop.Value = 0
chkPrograms.Value = 0
chkStartup.Value = 0
chkAccessories.Value = 0
chkFavorites.Value = 0
txtShortcutDir = msWinDir & "\Start Menu"
End If
End Sub
Private Sub chkPrograms_Click()
If chkPrograms.Value = 1 Then
chkQuickLaunchBar.Value = 0
chkStartmenu.Value = 0
chkDesktop.Value = 0
chkStartup.Value = 0
chkAccessories.Value = 0
chkFavorites.Value = 0
txtShortcutDir = msWinDir & "\Start Menu\Programs"
End If
End Sub
Private Sub chkStartup_Click()
If chkStartup.Value = 1 Then
chkQuickLaunchBar.Value = 0
chkPrograms.Value = 0
chkStartmenu.Value = 0
chkDesktop.Value = 0
chkPrograms.Value = 0
chkAccessories.Value = 0
chkFavorites.Value = 0
txtShortcutDir = msWinDir & "\Start Menu\Programs\Startup"
End If
End Sub
Private Sub chkAccessories_Click()
If chkAccessories.Value = 1 Then
chkQuickLaunchBar.Value = 0
chkStartup.Value = 0
chkPrograms.Value = 0
chkStartmenu.Value = 0
chkDesktop.Value = 0
chkPrograms.Value = 0
chkFavorites.Value = 0
txtShortcutDir = msWinDir & "\Start Menu\Programs\Accessories"
End If
End Sub
Private Sub chkFavorites_Click()
If chkFavorites.Value = 1 Then
chkQuickLaunchBar.Value = 0
chkAccessories.Value = 0
chkStartup.Value = 0
chkPrograms.Value = 0
chkStartmenu.Value = 0
chkDesktop.Value = 0
chkPrograms.Value = 0
txtShortcutDir = msWinDir & "\Favorites"
End If
End Sub

Private Sub chkLnkname_Click()
If chkLnkname.Value = 1 Then
txtExename.Text = ""
txtShortcutName.Text = ""
cmdExeDialog.Visible = True
lblExename.Visible = True
txtExename.Visible = True
lblUrlname.Visible = False
txtUrlname.Visible = False
chkUrlname.Value = 0
cmdExeDialog.Visible = True
End If
End Sub

Private Sub chkUrlname_Click()
If chkUrlname.Value = 1 Then
txtUrlname.Text = "http://www."
txtShortcutName.Text = ""
'cmdExeDialog.Visible = False
lblExename.Visible = False
txtExename.Visible = False
lblUrlname.Visible = True
txtUrlname.Visible = True
chkLnkname.Value = 0
End If
End Sub

Private Sub cmdCreateShortcut_Click()
If chkLnkname.Value = 1 Then
If Len(txtExename) = 0 Then
    MsgBox "Please enter a valid Path and Filename" & vbCrLf & vbCrLf & _
    "Eg... C:\Windows\Notepad.exe", vbQuestion
    txtExename.SetFocus
    Exit Sub
Else
If chkUrlname.Value = 1 Then
If Len(txtUrlname) = 0 Then
    MsgBox "Please enter a valid Path and Filename" & vbCrLf & vbCrLf & _
    "Eg... http://www.myshortcut.com", vbQuestion
    txtUrlname.SetFocus
Exit Sub
End If
End If
End If
End If

If Len(txtShortcutName) = 0 Then
    MsgBox "Please enter a Shortcut Name. Eg... Notepad", vbQuestion
    txtShortcutName.SetFocus
    Exit Sub
End If

If Len(txtShortcutDir) = 0 Then
    MsgBox "Please enter a Directory to create you new shortcut in." & vbCrLf & vbCrLf & _
    "Eg... C:\Windows\Desktop", vbQuestion
    txtShortcutDir.SetFocus
    Exit Sub
End If

'This loop does the exact same thing as the code above.  However, if you have
'a lot of controls, the loop would probably be slower since it must check
'every control to determine if it's a textbox; but, in some circumstances, a
'loop like this might be easier to use.
'Dim Obj As Object
'For Each Obj In CreateDeleteShortcut
'    If TypeOf Obj Is TextBox Then
'        If Len(Obj) = 0 Then
'            MsgBox "Did you forget something?", vbQuestion
'            Obj.SetFocus
'            Exit Sub
'        End If
'    End If
'Next Obj
If chkLnkname.Value = 1 Then
CreateShortcut txtShortcutDir, txtShortcutName, txtExename, ""
ElseIf chkUrlname.Value = 1 Then
CreateShortcut txtShortcutDir, txtShortcutName & ".url", txtUrlname, ""
'CreateInternetShortCut txtShortcutDir, txtShortcutName & ".url", txtUrlname
'"C:\windows\desktop\test.url", "http://www.vbcode.com"
End If
Timer1.Enabled = True
End Sub
Function FileOrDirExists(Optional ByVal sFileName As Variant, Optional ByVal sPath As Variant) As Boolean

On Error GoTo Oops
If IsMissing(sPath) Then
    'Only a file name was passed.
    If Len(Dir$(sFileName)) Then FileOrDirExists = True
Else
    'A directory was passed
    'Append a backslash to the pathname, if necessary.
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    If IsMissing(sFileName) Then
        'Directory was passed, but not a file, so determine if
        'the directory exists
        If Len(Dir$(sPath, vbDirectory)) Then FileOrDirExists = True
    Else
        'Both a directory and a file were passed, so determine
        'if the file exists in the specified directory.
        If Len(Dir$(sPath & sFileName)) Then FileOrDirExists = True
    End If
End If
Exit Function

Oops:
Exit Function

End Function
Private Sub cmdExeDialog_Click()

Dim sFileName As String
Dim sFileTitle As String
Dim nDot As Integer

'txtExename.SetFocus

If ShowOpen(sFileName, sFileTitle) Then
 If chkLnkname.Value = 1 Then
    txtExename = sFileName
    Else: chkUrlname.Value = 1
    txtUrlname = sFileName
    End If
    nDot = InStr(sFileTitle, Chr$(46))
    With txtShortcutName
        .Text = Left$(sFileTitle, nDot - 1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End If

End Sub

Private Sub cmdExit_Click()
Me.Hide
Unload Me
End
End Sub

Private Sub cmdShortcutPathDialog_Click()

Dim udtBrowseInfo As BROWSEINFO
Dim lRet As Long
Dim lPathID As Long
Dim sPath As String
Dim nNullPos As Integer

txtShortcutDir.SetFocus

'Specify the window handle for the owner of the dialog box
udtBrowseInfo.hOwner = Me.hwnd

'Specify the root to start browsing from;
'if null, My Computer is the root
udtBrowseInfo.pidlRoot = 0&

'Specify a title.  This is not the caption of the dialog.  Useful for
'adding any kind of additional information or instructions
udtBrowseInfo.lpszTitle = "Select a folder for Shortcut"

'Specify any flags; See Declarations section
udtBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS

'Call the function.
'The return value is a pointer to an item identifier list that
'specifies the location of the selected folder.
'If the user cancels the dialog box, the return value is 0.
lPathID = SHBrowseForFolder(udtBrowseInfo)

sPath = Space$(512)
lRet = SHGetPathFromIDList(lPathID, sPath)

If lRet Then
    nNullPos = InStr(sPath, vbNullChar)
    txtShortcutDir = Left(sPath, nNullPos - 1)
End If


End Sub

Private Sub cmdDelete_Click()
DeleteShortcut.Show
End Sub

Private Sub Form_Load()
chkDesktop.Value = 1
chkLnkname.Value = 1
Text1.Text = 5

Dim lBytes As Long

lBytes = GetWindowsDirectory(msBuffer, BUFFERSIZE)
msWinDir = Left$(msBuffer, lBytes)

txtShortcutDir = msWinDir & "\Desktop"
If txtShortcutDir = msWinDir & "\Desktop\" Then
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
Text1.Text = Text1.Text - 1
If Text1.Text = 0 Then
txtExename.Text = ""
txtShortcutName.Text = ""
End If
If txtExename.Text = "" Then
Text1.Text = 5
Timer1.Enabled = False
End If
End Sub

Private Sub txtExeName_LostFocus()
If chkLnkname.Value = 1 Then
If Not FileOrDirExists(txtExename) And ActiveControl.Name <> "cmdExeDialog" Then
    MsgBox "File Not Found" & vbCrLf & vbCrLf & txtExename, vbInformation
    With txtExename
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    Else
If chkUrlname.Value = 1 Then
If Not FileOrDirExists(txtUrlname) And ActiveControl.Name <> "cmdExeDialog" Then
    MsgBox "File Not Found" & vbCrLf & vbCrLf & txtUrlname, vbInformation
    With txtUrlname
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End If
End If
End If
End If
End Sub


Private Sub txtShortcutDir_LostFocus()

If Not FileOrDirExists(, txtShortcutDir) And ActiveControl.Name <> "cmdShortcutPathDialog" Then
    MsgBox "Path Not Found" & vbCrLf & vbCrLf & txtShortcutDir, vbInformation
    With txtShortcutDir
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End If

End Sub
Private Sub txtShortcutName_KeyPress(KeyAscii As Integer)

'Disallow characters which are not valid in file names
Select Case KeyAscii
    Case 92, 47, 58, 42, 63, 34, 60, 62, 124
        MsgBox "A filename cannot contain any of the following characters:" & vbCrLf & vbCrLf & _
             Space$(15) & Space$(15) & "/ \ : * ? "" < > |", vbCritical, App.Title
        KeyAscii = 0
End Select

End Sub

Sub CreateInternetShortCut(URLFile As String, URLTarget As String)
'"C:\windows\desktop\test.url", "http://www.vbcode.com"
'An Internet Shortcut takes on the form of
'   [InternetShortcut]
'   URL=http://www.YourWebSiteAddress.com

    Dim intFreeFile As Integer

    'get free file number
    intFreeFile = FreeFile
    
    'print the URL file
    Open txtShortcutDir & "\" & URLFile For Output As intFreeFile
    Print #intFreeFile, "[InternetShortcut]"
    Print #intFreeFile, "URL=" & URLTarget
    Close intFreeFile
End Sub

