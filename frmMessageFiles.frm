VERSION 5.00
Begin VB.Form frmMessageFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message File Selection"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreateFolder 
      Caption         =   "Create &Folder"
      Height          =   330
      Left            =   2100
      TabIndex        =   14
      Top             =   5010
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   4935
      TabIndex        =   13
      Top             =   7005
      Width           =   1485
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   465
      Left            =   3345
      TabIndex        =   11
      Top             =   7005
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   285
      TabIndex        =   9
      Top             =   6060
      Width           =   6150
      Begin VB.Label lblInstr 
         Caption         =   $"frmMessageFiles.frx":0000
         Height          =   630
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   5910
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   360
      Left            =   285
      TabIndex        =   7
      Top             =   5685
      Width           =   6150
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   0
      ScaleHeight     =   1380
      ScaleWidth      =   6645
      TabIndex        =   3
      Top             =   0
      Width           =   6675
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   270
         Picture         =   "frmMessageFiles.frx":00DB
         Stretch         =   -1  'True
         Top             =   150
         Width           =   1635
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H8000000E&
         Caption         =   "Selection/Creation of Message Files"
         Height          =   420
         Left            =   2835
         TabIndex        =   4
         Top             =   480
         Width           =   2985
      End
   End
   Begin VB.FileListBox flFiles 
      Height          =   3240
      Left            =   3555
      TabIndex        =   2
      Top             =   1725
      Width           =   2895
   End
   Begin VB.DirListBox dirDirs 
      Height          =   2490
      Left            =   285
      TabIndex        =   1
      Top             =   2445
      Width           =   3015
   End
   Begin VB.DriveListBox drDrives 
      Height          =   330
      Left            =   285
      TabIndex        =   0
      Top             =   1770
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Files:"
      Height          =   345
      Left            =   3570
      TabIndex        =   12
      Top             =   1500
      Width           =   570
   End
   Begin VB.Label Label4 
      Caption         =   "Path:"
      Height          =   210
      Left            =   285
      TabIndex        =   8
      Top             =   5400
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "Folder:"
      Height          =   225
      Left            =   285
      TabIndex        =   6
      Top             =   2190
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "Drive:"
      Height          =   345
      Left            =   285
      TabIndex        =   5
      Top             =   1500
      Width           =   570
   End
End
Attribute VB_Name = "frmMessageFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Message files form
'------------------------------------------------------------------------------------
Option Explicit

Private m_strMessageFilePath As String
Private m_bReturn As Boolean
Private m_strPathIncoming As String
Private m_strMessageFileType As String

Public Property Get MessageFilePath() As String

    MessageFilePath = m_strMessageFilePath
    
End Property

Public Property Let MessageFilePath(ByVal strNewValue As String)

    m_strMessageFilePath = strNewValue
    
End Property

Private Sub BuildPath(ByVal FTP As Boolean)

    'build path from components
    Dim strPath As String
    Dim strDrive As String
    Dim strSep As String
    
    strSep = ""
    strDrive = drDrives.Drive & "\"
    
    If Not LCase$(dirDirs.Path) = strDrive Then
        strSep = "\"
    End If
    
    If FTP Then
        strPath = dirDirs.Path
    Else
        strPath = dirDirs.Path & strSep & flFiles.FileName
    End If
        
    m_strMessageFilePath = strPath
    
    txtPath.Text = m_strMessageFilePath
    txtPath.ToolTipText = m_strMessageFilePath

End Sub

Private Sub cmdCancel_Click()

    m_strMessageFilePath = ""
    Me.Hide

End Sub

Private Sub cmdCreateFolder_Click()

    Dim fs As New FileSystemObject
    Dim sFolder As String
    Dim sSep As String
    Dim sPath As String
    Dim sNewFolder As String
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdCreateFolder_Click_Error
    
    sSep = "\"
    
    'get the current folder
    sFolder = dirDirs.Path
    
    If dirDirs.Path = "c:" Then
        sSep = ""
    End If
    
    sNewFolder = InputBox("Please type in the folder name into the box below:", "New Folder")
    If Len(Trim$(sNewFolder)) > 0 Then
        sPath = sFolder & sSep & sNewFolder
        fs.CreateFolder sPath
        dirDirs.Refresh
    End If
    
    
Exit_Properly:
    Exit Sub
    
cmdCreateFolder_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMessageFiles:cmdCreateFolder_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMessageFiles:cmdCreateFolder_Click - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdDone_Click()

    Dim bFTP As Boolean
    
    bFTP = False
    
    If LCase$(Trim$(m_strMessageFileType)) = "ftp" Then
        bFTP = True
    End If
    
    If Not ValidateFile(bFTP) Then
        If Not m_bReturn Then
            Exit Sub
        End If
    End If
    
    Me.Hide

End Sub

Private Sub dirDirs_Change()
    Dim lErrno As Long
    Dim sSource As String, sDesc As String

    On Error GoTo dirDirs_Change_Error
    
    flFiles.Path = dirDirs.Path
    flFiles.Pattern = "*.*"
    flFiles.Refresh
    
    BuildPath False

Exit_Properly:
    Exit Sub
    
dirDirs_Change_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMessageFiles:dirDirs_Change:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMessageFiles:dirDirs_Change - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Sub

Private Sub drDrives_Change()

    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo drDrives_Change_Error
    
    dirDirs.Path = drDrives.Drive
    dirDirs.Refresh
    
    BuildPath False

Exit_Properly:
    Exit Sub
    
drDrives_Change_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMessageFiles:drDrives_Change:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMessageFiles:drDrives_Change - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Sub

Private Sub flFiles_Click()

    BuildPath False

End Sub

Private Function ValidateFile(ByVal FTP As Boolean) As Boolean

    Dim fs As New FileSystemObject
    Dim lRet As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo ValidateFile_Error
    
    If Not fs.FileExists(m_strMessageFilePath) Then
        'is this just a folder
        If fs.FolderExists(m_strMessageFilePath) Then
            If Not FTP Then
                MsgBox "Please enter a valid file name for the Message File!", vbExclamation + vbOKOnly, "Missing file name"
                ValidateFile = False
                m_bReturn = False
            Else
                'exists and is valid for ftp path
                ValidateFile = True
                m_bReturn = True
            End If
            GoTo Exit_Properly
        End If
        lRet = MsgBox("The file '" & m_strMessageFilePath & "' does not exist!" & vbCrLf _
        & "Would you like to create this file now?", vbYesNoCancel, "File does not exist")
        If lRet = vbYes Then
            'create this file, then return
            fs.CreateTextFile m_strMessageFilePath, True
            If Not fs.FileExists(m_strMessageFilePath) Then
                MsgBox "File error - Unable to create file '" & m_strMessageFilePath & "'!", vbCritical & vbOKOnly, "File create error!"
                ValidateFile = False
                m_bReturn = True
                GoTo Exit_Properly
            End If
            ValidateFile = True
            m_bReturn = True
            GoTo Exit_Properly
        End If
        If lRet = vbNo Then
            m_strMessageFilePath = ""
            ValidateFile = False
            m_bReturn = True
            GoTo Exit_Properly
        End If
        If lRet = vbCancel Then
            ValidateFile = False
            m_bReturn = False
            GoTo Exit_Properly
        End If
    Else
        lRet = MsgBox("The file '" & m_strMessageFilePath & "' already exists - continue using this file?", vbYesNo, "File exists")
        If lRet = vbYes Then
            ValidateFile = True
            m_bReturn = True
            GoTo Exit_Properly
        End If
        If lRet = vbNo Then
            ValidateFile = False
            m_bReturn = False
        End If
    End If
    
Exit_Properly:
    Set fs = Nothing
    Exit Function
    
ValidateFile_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMessageFiles:ValidateFile:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMessageFiles:ValidateFile - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Function

Private Sub Form_Load()

    Dim lPos As Long
    Dim lLastPos As Long
    Dim fs As New FileSystemObject
    Dim iIndex As Integer
    Dim sItem As String
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo Form_Load_Error
    
    'sort out the main caption
    If LCase$(Trim$(m_strMessageFileType)) = "logging" Then
        Me.Caption = m_strMessageFileType & " - File Selection"
        lblTitle.Caption = "Selection/Creation of Logging File"
        lblInstr.Caption = "Select drive and folder and, optionally, select a new filename from the file list.  Should you wish to create a new logging file, select drive and directory, then type in the name of the new file, then click 'Done'."
    ElseIf LCase$(Trim$(m_strMessageFileType)) = "ftp" Then
        Me.Caption = m_strMessageFileType & " File Path Selection"
        lblTitle.Caption = "Selection/Creation of FTP File Path"
        lblInstr.Caption = "Select drive and folder then click 'Done'."
    Else
        Me.Caption = m_strMessageFileType & " - Message File Selection"
        lblTitle.Caption = "Selection/Creation of Message Files"
        lblInstr.Caption = "Select drive and folder and, optionally, select a new filename from the file list.  Should you wish to create a new message file, select drive and directory, then type in the name of the new file, then click 'Done'."
    End If
    
    m_bReturn = False
    
    'is the incoming path to a valid file
    If LCase$(Trim$(m_strMessageFileType)) <> "ftp" Then
        flFiles.Enabled = True
        If fs.FileExists(m_strPathIncoming) Then
            txtPath.Text = m_strPathIncoming
            
            drDrives.Drive = Mid$(m_strPathIncoming, 1, InStr(m_strPathIncoming, ":"))
            lPos = 1
            Do
                lPos = InStr(lPos + 1, m_strPathIncoming, "\")
                If Not lPos = 0 Then
                    lLastPos = lPos
                End If
            Loop While lPos > 0
            
            'dirdirs will use the full path
            dirDirs.Path = Mid$(m_strPathIncoming, 1, lLastPos)
            flFiles.Path = Mid$(m_strPathIncoming, 1, lLastPos)
            flFiles.Refresh
            For iIndex = 0 To flFiles.ListCount - 1
                sItem = flFiles.List(iIndex)
                If sItem = Mid$(m_strPathIncoming, lLastPos + 1) Then
                    flFiles.Selected(iIndex) = True
                    Exit For
                End If
            Next iIndex
        End If
        BuildPath False
    Else
        'this is an FTP file path
        flFiles.Enabled = False
        If fs.FolderExists(m_strPathIncoming) Then
            txtPath.Text = m_strPathIncoming
            drDrives.Drive = Mid$(m_strPathIncoming, 1, InStr(m_strPathIncoming, ":"))
            dirDirs.Path = m_strPathIncoming
        End If
        BuildPath True
    End If
            

Exit_Properly:
    If Not fs Is Nothing Then
        Set fs = Nothing
    End If
    Exit Sub
    
Form_Load_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMessageFiles:Form_Load:-" & vbCr & vbCr & _
    "Error number: " & Err.Number & vbCr & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Critical load error"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMessageFiles:Form_Load - " & sSource
    Log "**************************************"
    GoTo Exit_Properly
    
End Sub

Private Sub txtPath_Change()

    m_strMessageFilePath = txtPath.Text

End Sub

Public Property Let PathIncoming(ByVal strNew As String)

    m_strPathIncoming = strNew
    
End Property

Public Property Let MessageFileType(ByVal strNew As String)

    m_strMessageFileType = strNew
    
End Property
