VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFTPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCommandDate As Date
Private mvarCommandTime As Date
Private mvarHostIP As String
Private mvarHostPort As Long
Private mvarSessionID As String

Private mvarCommandEvents As CFTPCommandEvents

Private WithEvents moFTPClient As Inet
Attribute moFTPClient.VB_VarHelpID = -1

Private mvarFilePath As String
Private mvarFTPHostName As String
Private mvarFTPUserName As String
Private mvarFTPPassword As String

Private Const MODULE As String = "CFTPClient::"

Public Event StateChanged(ByVal StateString As String, ByVal State As Integer)

Public Sub SendFTP(ByVal FileName As String)

    Dim sSep As String
    Dim sFileNameAndPath As String
    
    On Error GoTo SendFTP_Error
    
    sSep = ""
    
    If Trim$(mvarFilePath) = "" Then
        sFileNameAndPath = App.Path & "\" & FileName & ".dat"
    Else
        If InStr(Trim$(mvarFilePath), "\") = 0 Then
            sSep = "\"
        End If
        sFileNameAndPath = Trim$(mvarFilePath) & sSep & Trim$(FileName) & ".dat"
    End If
        
    moFTPClient.url = g_oGoddess.FTPRemoteHost
    moFTPClient.Username = g_oGoddess.FTPUserName
    moFTPClient.Password = g_oGoddess.FTPPassword

    'moFTPClient.OpenURL
    
    moFTPClient.Execute , "PUT """ & sFileNameAndPath & """ """ & Trim$(FileName) & ".dat"""
    
    Do Until moFTPClient.StillExecuting = False
        DoEvents
    Loop
    
    Log "SendFTP success, file '" & Trim$(FileName) & "' sent to '" & moFTPClient.url & "'"
    
Exit_Properly:
    moFTPClient.Execute , "close"
    Do Until moFTPClient.StillExecuting = False
        DoEvents
    Loop
    moFTPClient.Execute , "BYE"
    Do Until moFTPClient.StillExecuting = False
        DoEvents
    Loop
    Exit Sub
    
SendFTP_Error:
    Err.Raise Err.Number, MODULE & "SendFTP", Err.Description
    
End Sub

Public Property Let FTPPassword(ByVal vData As String)
    
    mvarFTPPassword = Trim$(vData)
    moFTPClient.Password = mvarFTPPassword
    
End Property

Public Property Get FTPPassword() As String
    
    FTPPassword = mvarFTPPassword
    
End Property

Public Property Let FTPUserName(ByVal vData As String)
    
    mvarFTPUserName = Trim$(vData)
    moFTPClient.Username = mvarFTPUserName
    
End Property

Public Property Get FTPUserName() As String
    
    FTPUserName = mvarFTPUserName
    
End Property

Public Property Let FTPHostName(ByVal vData As String)
    
    mvarFTPHostName = Trim$(vData)
    moFTPClient.url = mvarFTPHostName
    
End Property

Public Property Get FTPHostName() As String
    
    FTPHostName = mvarFTPHostName
    
End Property

Public Property Let FilePath(ByVal vData As String)
    
    mvarFilePath = Trim(vData)
    moFTPClient.Document = mvarFilePath
    
End Property

Public Property Get FilePath() As String
    
    FilePath = mvarFilePath
    
End Property

Public Sub InitFTPClient(FTPClient As Inet)

    Set moFTPClient = FTPClient
    
End Sub

Public Property Set CommandEvents(ByVal vData As CFTPCommandEvents)
    
    Set mvarCommandEvents = vData
    
End Property

Public Property Get CommandEvents() As CFTPCommandEvents
    
    Set CommandEvents = mvarCommandEvents
    
End Property

Public Property Let HostPort(ByVal vData As Long)
    
    mvarHostPort = vData
    
End Property

Public Property Get HostPort() As Long
    
    HostPort = mvarHostPort
    
End Property

Public Property Let HostIP(ByVal vData As String)
    
    mvarHostIP = vData
    
End Property

Public Property Get HostIP() As String
    
    HostIP = mvarHostIP
    
End Property

Public Property Let CommandTime(ByVal vData As Date)
    
    mvarCommandTime = vData
    
End Property

Public Property Get CommandTime() As Date
    
    CommandTime = mvarCommandTime
    
End Property

Public Property Let CommandDate(ByVal vData As Date)
    
    mvarCommandDate = vData
    
End Property

Public Property Get CommandDate() As Date
    
    CommandDate = mvarCommandDate
    
End Property

Private Sub Class_Initialize()

    Set mvarCommandEvents = New CFTPCommandEvents
    
End Sub

Private Sub Class_Terminate()

    Set moFTPClient = Nothing
    
    Set mvarCommandEvents = Nothing
    
End Sub

Private Sub moFTPClient_StateChanged(ByVal State As Integer)

    Dim sState As String
    
    Select Case State
        Case 1:
            'resolving host
            sState = "Resolving Host '" & mvarFTPHostName & "'"
            
        Case 2:
            'host resolved
            sState = "Host '" & mvarFTPHostName & "' resolved"
            
        Case 3:
            'connecting
            sState = "Connecting to Host"
            
        Case 4:
            'connected
            sState = "Connected successfully"
            
        Case 5:
            'requesting
            sState = "Requesting"
            
        Case 6:
            'request sent
            sState = "Request sent"
            
        Case 7:
            'receiving response
            sState = "Receiving response"
            
        Case 8:
            'response received
            sState = "Response received"
            
        Case 9:
            'disconnecting
            sState = "Disconnecting from '" & mvarFTPHostName & "'..."
            
        Case 10:
            'disconnected
            sState = "Disconnected"
            
        Case 11:
            'error
            sState = "Error occured - " & moFTPClient.ResponseInfo
            
        Case 12:
            'response completed
            sState = "Response completed"
            
        Case Else:
            'undefined response
            sState = "Undefined state change"
            
    End Select
       
    RaiseEvent StateChanged(sState, State)
    
End Sub
