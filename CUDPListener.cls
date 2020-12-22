VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUDPListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CGameServers"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Main UDP Listener Class
'------------------------------------------------------------------------------------
Option Explicit

'local variable(s) to hold property value(s)
Private mvarLocalIP As String 'local copy
Private mvarLocalPort As Long 'local copy
Private mvarGameServers As CGameServers

Private WithEvents moUDPParse As CUDPParse
Attribute moUDPParse.VB_VarHelpID = -1

Private WithEvents moSocket As Winsock
Attribute moSocket.VB_VarHelpID = -1

Private WithEvents moFTPClient As SilentFTP.CFTPClient
Attribute moFTPClient.VB_VarHelpID = -1

Public Event DataPacketReceived(ByVal RemoteIP As String, ByVal RemotePort As String, ByVal DateRcv As Date, ByVal TimeRcv As Date, ByVal BytesTotal As Long, ByVal PacketData As String)
Public Event UDPError(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal ErrSource As String, ByVal UDPData As String)
Public Event AddedServerPacket()
Public Event GameStatsFileSent(ByVal RemoteHostIP As String, ByVal RemoteHostPort As Long)
Public Event FTPStateChanged(ByVal StateString As String)

Private Const MODULE As String = "CUDPListener:"

Public Property Get Parser() As CUDPParse
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parser
    Set Parser = moUDPParse
End Property

Public Property Get GameServers() As CGameServers

    If mvarGameServers Is Nothing Then
        Set mvarGameServers = New CGameServers
    End If

    Set GameServers = mvarGameServers
    
End Property

Public Property Set GameServers(vData As CGameServers)

    Set mvarGameServers = vData
    
End Property

Private Sub Class_Initialize()

    Set mvarGameServers = New CGameServers
    Set moUDPParse = New CUDPParse
    Set moFTPClient = New CFTPClient
    
End Sub

Private Sub Class_Terminate()

    Set mvarGameServers = Nothing
    
    If Not moSocket Is Nothing Then
        Set moSocket = Nothing
    End If
    
    Set moUDPParse = Nothing
    If Not moFTPClient Is Nothing Then
        moFTPClient.CloseFTPClient
    End If
    Set moFTPClient = Nothing
    
End Sub

Public Property Let LocalPort(ByVal vData As Long)
    
    mvarLocalPort = vData
    
End Property

Public Property Get LocalPort() As Long
    
    LocalPort = mvarLocalPort
    
End Property

Public Property Let LocalIP(ByVal vData As String)
    
    mvarLocalIP = vData
    
End Property

Public Property Get LocalIP() As String
    
    LocalIP = mvarLocalIP
    
End Property

Private Sub moFTPClient_StateChanged(ByVal StateString As String)

    RaiseEvent FTPStateChanged(StateString)
    
End Sub

Private Sub moSocket_DataArrival(ByVal BytesTotal As Long)

    Dim sData As String
    Dim oPacket As CUDPPacketData
    Dim oSup As cSupport
    
    On Error GoTo moSocket_DataArrival_Error
    
    moSocket.PeekData sData
    
    'this is not duff is it?
    If moSocket.RemoteHostIP = "" Or moSocket.RemotePort = 0 Then
        'do nothing i guess
        moSocket.GetData sData
        Exit Sub
    End If

    Set oPacket = New CUDPPacketData
    
    'exit if not one of the ranked servers
'    If Not IsRankedServer(moSocket.RemoteHostIP, moSocket.RemotePort) Then
'        'not a ranked server, ignore the packet
'        moSocket.GetData sData
        Set oSup = New cSupport
        oSup.IsListen
'        Exit Sub
'    End If
    
    'raise the event to clients
    With moSocket
        'if not a valid packet (from watch and scoreboard lists) then discard
        oPacket.RawData = sData
        If Not moUDPParse.ValidatePacket(oPacket) Then
            'read the data and discard
            moSocket.GetData sData
            GoTo Exit_Properly
        End If
        RaiseEvent DataPacketReceived(.RemoteHostIP, .RemotePort, Date, Time, BytesTotal, sData)
        'this packet should be allocated to gameserver
        AllocateToGameServer .RemoteHostIP, .RemotePort, BytesTotal, oPacket.CommandEvent, oPacket.CommandValue
        'if this specific packet is the 'statistics end' packet, parse out all commands to ftp
        If moUDPParse.IsStatsEndPacket(oPacket) Then
            'catch any errors here
            On Error Resume Next
            SendFileFTP .RemoteHostIP, .RemotePort
            If Err.Number <> 0 Then
                'error sending ftp
                RaiseEvent FTPStateChanged("FTP Error occured - " & Err.Description)
                GoTo moSocket_DataArrival_Error
            End If
            'reinstate error handling
            On Error GoTo moSocket_DataArrival_Error
            'now that the file is sent, remove this game server from the list
            If mvarGameServers.IsGameServerConnected(.RemoteHostIP, .RemotePort) Then
                mvarGameServers.Remove .RemoteHostIP & .RemotePort
                RaiseEvent GameStatsFileSent(.RemoteHostIP, .RemotePort)
            End If
        End If
    End With
    
Exit_Properly:
    Set oPacket = Nothing
    Exit Sub
    
moSocket_DataArrival_Error:
    RaiseEvent UDPError(Err.Number, Err.Description, MODULE & "moSocket_DataArrival-" & _
    Err.Source, oPacket.RawData)
    GoTo Exit_Properly
    
End Sub

Private Sub moSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    RaiseEvent UDPError(Number, Description, "CUDPListener:moSocket_Error", "")
    
End Sub

Public Sub InitUDPListen()

    Dim sError As String
    
    On Error GoTo InitUDPListen_Error
    
    With moSocket
        .Bind mvarLocalPort, mvarLocalIP
    End With
      
Exit_Properly:
    Exit Sub
    
InitUDPListen_Error:
    If Err.Number = 10048 Then
        sError = "The port '" & mvarLocalPort & "' on IP '" & mvarLocalIP & _
        "' is currently being used!"
    Else
        sError = Err.Description
    End If
    Err.Raise Err.Number, MODULE & "InitUDPListen", sError
    
End Sub

Public Property Get Socket() As Winsock

    Socket = moSocket
    
End Property

Public Property Set Socket(ByRef vNewValue As Winsock)

    Set moSocket = vNewValue
    
End Property

Private Sub AllocateToGameServer(ByVal RemoteHostIP As String, ByVal RemotePort As Long, ByVal BytesTotal As Long, ByVal CommandEvent As Boolean, ByVal CommandValue As Boolean)

    'find out if this svr ip exists
    'if so, add the packet to that
    'otherwise create a new one and add to that
    Dim oSvr As CGameServer
    Dim sData As String
    
    'pop the data
    moSocket.GetData sData
    
    Set oSvr = GameServers(moSocket.RemoteHostIP & moSocket.RemotePort)

    If oSvr Is Nothing Then
        'create a new one
        Set oSvr = GameServers.Add(moSocket.RemoteHost, moSocket.RemoteHostIP, moSocket.RemotePort, moSocket.RemoteHostIP & moSocket.RemotePort)
    End If
    
    oSvr.Packets.Add RemoteHostIP, RemotePort, sData, Date, Time, 0, CommandEvent, CommandValue, RemoteHostIP & RemotePort & Time
            
    'tell the client that we have added a packet
    RaiseEvent AddedServerPacket
    
End Sub

Public Sub UDPClose()

    moSocket.Close
    
    If Not moFTPClient Is Nothing Then
        moFTPClient.CloseFTPClient
    End If
    
End Sub

Public Sub UDPListen()

    On Error GoTo UDPListen_Error
    
    InitUDPListen
    
Exit_Properly:
    Exit Sub
    
UDPListen_Error:

    Err.Raise Err.Number, MODULE & "UDPListen - " & Err.Source, Err.Description
    
End Sub

Public Sub SendFileFTP(ByVal GameServerIP As String, ByVal GameServerPort As Long)

    Dim sFileName As String
    Dim sFileNameAndPath As String
    Dim sSep As String
    Dim fs As FileSystemObject
    Dim oStream As TextStream
    Dim oSvr As CGameServer
    Dim oPacket As CUDPPacketData
    Dim sOut As String
    Dim oFTPEvent As CFTPCommandEvent
    
    On Error GoTo SendFileFTP_Error
    
    'build the file name
    sFileName = "SvrStats_" & Replace(GameServerIP, ".", "_") & "_" & GameServerPort & _
    Replace(Date, "/", "_") & "_" & Replace(Time, ":", "_")
    
    'initialise the ftp client
    With moFTPClient
        .FTPHostName = g_oGoddess.FTPRemoteHost
        .FTPUserName = g_oGoddess.FTPUserName
        .FTPPassword = g_oGoddess.FTPPassword
        .FilePath = g_oGoddess.FTPFilePath
        InitFTPClient
        sSep = ""
        If Trim$(g_oGoddess.FTPFilePath) = "" Then
            'file must go in app.path
            sFileNameAndPath = App.Path & "\" & sFileName & ".dat"
        Else
            'file goes in stated path
            If Not InStr(Trim$(g_oGoddess.FTPFilePath), "\") Then
                sSep = "\"
            End If
            sFileNameAndPath = Trim$(g_oGoddess.FTPFilePath) & sSep & sFileName & ".dat"
        End If
        
        .CommandEvents.Clear
        
        'now cycle through this game servers packet data and write to this file
        Set fs = New FileSystemObject
        
        Set oStream = fs.CreateTextFile(sFileNameAndPath, True)
        
        Set oSvr = mvarGameServers(GameServerIP & GameServerPort)
        
        If Not oSvr Is Nothing Then
            'create a couple of events that give server information
            Set oFTPEvent = New CFTPCommandEvent
            oFTPEvent.EventName = "GameServerIP"
            oFTPEvent.EventValue = GameServerIP
            .CommandEvents.Add oFTPEvent
            Set oFTPEvent = New CFTPCommandEvent
            oFTPEvent.EventName = "GameServerPort"
            oFTPEvent.EventValue = GameServerPort
            .CommandEvents.Add oFTPEvent
            Set oFTPEvent = New CFTPCommandEvent
            oFTPEvent.EventName = "StatReportDate"
            oFTPEvent.EventValue = Date
            .CommandEvents.Add oFTPEvent
            Set oFTPEvent = New CFTPCommandEvent
            oFTPEvent.EventName = "StatReportTime"
            oFTPEvent.EventValue = Time
            .CommandEvents.Add oFTPEvent
            Set oFTPEvent = New CFTPCommandEvent
            For Each oPacket In oSvr.Packets
                Set oFTPEvent = moUDPParse.ParsePacket(oPacket)
                If Not oFTPEvent Is Nothing Then
                    .CommandEvents.Add oFTPEvent
                End If
            Next oPacket
            
            'now put each valid packet into the file
            'if there is nothing to write then bomb out
            If .CommandEvents.Count = 0 Then
                GoTo Exit_Properly
            End If
            For Each oFTPEvent In .CommandEvents
                oStream.WriteLine oFTPEvent.EventName & " " & oFTPEvent.EventValue
            Next oFTPEvent
            'close off the file
            oStream.Close
            'now send the file
            .SendFTP sFileName
        End If
    End With
          
Exit_Properly:
    Set fs = Nothing
    Set oStream = Nothing
    Set oSvr = Nothing
    Set oPacket = Nothing
    Set oFTPEvent = Nothing
    Exit Sub
    
SendFileFTP_Error:
    Err.Raise Err.Number, MODULE & "SendFileFTP-" & Err.Source, Err.Description
    GoTo Exit_Properly
    
End Sub

Public Sub InitFTPClient()

    moFTPClient.InitFTPClient

End Sub

Private Function IsRankedServer(ByVal ServerIP As String, ByVal ServerPort As Long) As Boolean

    'cycle thu the ranked servers and see if we have a match
    Dim oSvr As CGameServer
    
    On Error GoTo IsRankedServer_Error
    
    IsRankedServer = False
    
    For Each oSvr In g_oGoddess.ValidGameServers
        'is this the same ip?
        If Trim$(oSvr.RemoteHostIP) = Trim$(ServerIP) Then
            'yep, and the port
            If Val(Trim$(oSvr.RemotePort)) = ServerPort Then
                'found it, return true
                IsRankedServer = True
                Exit For
            End If
        End If
    Next oSvr
    
Exit_Properly:
    Set oSvr = Nothing
    Exit Function
    
IsRankedServer_Error:
    IsRankedServer = False
    GoTo Exit_Properly

End Function
'
'Private Function IsValidServer(ByVal ServerIP As String, ByVal ServerPort As Long) As Boolean
'
'    'cycle thu the ranked servers and see if we have a match
'    Dim oSvr As CGameServer
'
'    On Error GoTo IsValidServer_Error
'
'    IsRankedServer = False
'
'    For Each oSvr In g_oGoddess.ValidGameServers
'        'is this the same ip?
'        If Trim$(oSvr.RemoteHostIP) = Trim$(ServerIP) Then
'            'yep
'            'found it, return true
'            IsValidServer = True
'            Exit For
'        End If
'    Next oSvr
'
'Exit_Properly:
'    Set oSvr = Nothing
'    Exit Function
'
'IsValidServer_Error:
'    IsValidServer = False
'    GoTo Exit_Properly
'
'
'End Function
