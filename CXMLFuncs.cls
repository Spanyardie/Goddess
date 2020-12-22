VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXMLFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          GODDESS XML Main Class
'------------------------------------------------------------------------------------
Option Explicit

Private mvarVersion As String
Private mvarSupportedProtocols As CSupportedProtocols
Private mvarFTPRemoteHost As String
Private mvarFTPUserName As String
Private mvarFTPPassword As String
Private mvarFTPFilePath As String
Private mvarMessageFiles As CMessageFiles
Private mvarMessages As CMessages
Private mvarFilters As CFilters
Private mvarLoggingEnabled As Boolean
Private mvarLogFilePath As String

Private moValidGameServers As CGameServers

Private Const MODULE As String = "CXMLFuncs::"

Private moGODDESSXML As DOMDocument30
Private moSCOREBOARDXML As DOMDocument30
Private moMESSAGESXML As DOMDocument30
Private moWATCHXML As DOMDocument30

Public Property Set Filters(ByVal vData As CFilters)

    Set mvarFilters = vData
    
End Property

Public Property Get Filters() As CFilters

    Set Filters = mvarFilters
    
End Property

Public Property Set Messages(ByVal vData As CMessages)

    Set mvarMessages = vData
    
End Property

Public Property Get Messages() As CMessages

    Set Messages = mvarMessages
    
End Property

Public Function SetHostPort(ByVal Port As Long) As Boolean

    Dim oNode As IXMLDOMNodeList
    Dim oHost As IXMLDOMNode
    Dim oPort As IXMLDOMNode
    Dim oName As IXMLDOMNode
    
    If Trim$(Port) = "" Then
        SetHostPort = False
        Exit Function
    End If
    
    Set oNode = moGODDESSXML.selectNodes("xml/goddess/supportedprotocols/protocol")
    
    If oNode Is Nothing Then
        SetHostPort = False
        GoTo Exit_Properly
    End If
    
    For Each oHost In oNode
        Set oName = oHost.selectSingleNode("name")
        If oName Is Nothing Then
            SetHostPort = False
            GoTo Exit_Properly
        End If
        If LCase$(Trim$(oName.Text)) = "udp" Then
            Set oPort = oHost.selectSingleNode("listenport")
            If oPort Is Nothing Then
                SetHostPort = False
                GoTo Exit_Properly
            End If
            oPort.Text = Trim$(Port)
            SetHostPort = True
            SaveXML xtGODDESS
            mvarSupportedProtocols("udp").ListenPort = Trim$(Port)
            Exit For
        End If
    Next oHost
        
Exit_Properly:
    Set oNode = Nothing
    Set oHost = Nothing
    Set oPort = Nothing
    Set oName = Nothing
    
End Function

Public Function SetHostIP(ByVal HostIP As String) As Boolean

    Dim oNode As IXMLDOMNodeList
    Dim oHost As IXMLDOMNode
    Dim oIP As IXMLDOMNode
    Dim oName As IXMLDOMNode
    
    If Trim$(HostIP) = "" Then
        SetHostIP = False
        Exit Function
    End If
    
    Set oNode = moGODDESSXML.selectNodes("xml/goddess/supportedprotocols/protocol")
    
    If oNode Is Nothing Then
        SetHostIP = False
        GoTo Exit_Properly
    End If
    
    For Each oHost In oNode
        Set oName = oHost.selectSingleNode("name")
        If oName Is Nothing Then
            SetHostIP = False
            GoTo Exit_Properly
        End If
        If LCase$(Trim$(oName.Text)) = "udp" Then
            Set oIP = oHost.selectSingleNode("localhost")
            If oIP Is Nothing Then
                SetHostIP = False
                GoTo Exit_Properly
            End If
            oIP.Text = Trim$(HostIP)
            SetHostIP = True
            SaveXML xtGODDESS
            mvarSupportedProtocols("udp").LocalHost = Trim$(HostIP)
            Exit For
        End If
    Next oHost
        
Exit_Properly:
    Set oNode = Nothing
    Set oHost = Nothing
    Set oIP = Nothing
    Set oName = Nothing
    
End Function

Public Function AddProtocol(ByVal ProtocolName As String, ByVal PacketLength As Long, ByVal LocalHost As String, ByVal ListenPort As Long) As Boolean

    On Error GoTo AddProtocol_Error
    
    mvarSupportedProtocols.Add ProtocolName, PacketLength, LocalHost, ListenPort, ProtocolName

    AddProtocol = True
    
Exit_Properly:
    Exit Function
    
AddProtocol_Error:
    AddProtocol = False
    GoTo Exit_Properly

End Function

Public Function SetMessageFilePath(ByVal MessageType As String, ByVal FilePath As String) As Boolean

    Dim oNodePaths As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim oPath As IXMLDOMNode
    Dim sType As String
    
    On Error GoTo SetMessageFilePath_Error
    
    Set oNodePaths = moGODDESSXML.selectNodes("xml/goddess/messagefilepaths/messagefile")

    For Each oNode In oNodePaths
        With oNode
            sType = .selectSingleNode("type").Text
            If Trim$(sType) = Trim$(MessageType) Then
                'this is the correct message type, insert the new path
                Set oPath = .selectSingleNode("path")
                If Not oPath Is Nothing Then
                    oPath.Text = FilePath
                    SaveXML xtMessages
                    Exit For
                End If
            End If
        End With
    Next oNode
    
    SetMessageFilePath = True
           
Exit_Properly:
    Exit Function
    
SetMessageFilePath_Error:
    SetMessageFilePath = False
    GoTo Exit_Properly
    
End Function

Private Function AddFTPDetails(ByVal RemoteHost As String, ByVal Username As String, ByVal Password As String, ByVal FilePath As String) As Boolean

    mvarFTPRemoteHost = RemoteHost
    
    mvarFTPUserName = Username
    
    mvarFTPPassword = Password

    mvarFTPFilePath = FilePath
    
    AddFTPDetails = True
    
End Function

Public Sub GetMessageFilePaths()
    
    Dim oNodePaths As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim lRet As Long
    Dim sType As String, sPath As String
    
    On Error GoTo GetMessageFilePaths_Error
    
    Set oNodePaths = moGODDESSXML.selectNodes("xml/goddess/messagefilepaths/messagefile")
    
    For Each oNode In oNodePaths
        With oNode
            sType = .selectSingleNode("type").Text
            sPath = .selectSingleNode("path").Text
            lRet = AddMessageFile(sType, sPath)
            'we can get by without message file paths since they can be input by user
            'so don't error here if lret is false
        End With
    Next oNode
            
Exit_Properly:
    Exit Sub
    
GetMessageFilePaths_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "GetMessageFilePaths - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub GetFTPInfo()

    Dim oNode As IXMLDOMNode
    Dim oText As IXMLDOMNode
    Dim lRet As Long
    Dim sRemoteHost As String, sUsername As String, sPassword As String
    Dim sFilePath As String
    
    On Error GoTo GetFTPInfo_Error
    
    Set oNode = moGODDESSXML.selectSingleNode("xml/goddess/ftp")
    
    With oNode
        Set oText = .selectSingleNode("remotehost")
        sRemoteHost = oText.Text
        sUsername = .selectSingleNode("username").Text
        sPassword = .selectSingleNode("password").Text
        sFilePath = .selectSingleNode("ftpfilepath").Text
        lRet = AddFTPDetails(sRemoteHost, sUsername, sPassword, sFilePath)
        If Not lRet Then
            'can't continue, raise error
            Err.Raise vbObjectError + 1000, MODULE & "GetFTPInfo", "Failed to retrieve FTP info!"
        End If
        
        'successfully added ftpinfo
    End With
        
Exit_Properly:
    Set oNode = Nothing
    Exit Sub
    
GetFTPInfo_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "GetFTPInfo - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub GetSupportedProtocols()

    Dim oNodeProtocols As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim lRet As Long
    Dim sName As String, sLocalHost As String
    Dim lPacketLen As Long, lListenPort As Long
    
    On Error GoTo GetSupportedProtocols_Error
    
    Set oNodeProtocols = moGODDESSXML.selectNodes("xml/goddess/supportedprotocols/protocol")
    
    For Each oNode In oNodeProtocols
        With oNode
            sName = .selectSingleNode("name").Text
            lPacketLen = Val(.selectSingleNode("packetlength").Text)
            lListenPort = Val(.selectSingleNode("listenport").Text)
            sLocalHost = .selectSingleNode("localhost").Text
            lRet = AddProtocol(sName, lPacketLen, sLocalHost, lListenPort)
            If Not lRet Then
                'can't continue, raise error
                Err.Raise vbObjectError + 1000, MODULE & "GetSupportedProtocols", "Failed to add protocol '" & .selectSingleNode("name") & "'!"
            End If
            
            'successfully added this protocol
        End With
    Next oNode
        
Exit_Properly:
    Set oNodeProtocols = Nothing
    Set oNode = Nothing
    Exit Sub
    
GetSupportedProtocols_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "GetSupportedProtocols - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub GetVersion()

    Dim oNodeVersion As IXMLDOMNode
    
    On Error GoTo GetVersion_Error
    
    'this retrieves the 'version' tag info
    Set oNodeVersion = moGODDESSXML.selectSingleNode("xml/goddess/version")
    
    If Not oNodeVersion Is Nothing Then
        mvarVersion = oNodeVersion.Text
    End If
    
Exit_Properly:
    Set oNodeVersion = Nothing
    Exit Sub
    
GetVersion_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "GetVersion - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Property Set MessageFiles(ByVal vData As CMessageFiles)

    Set mvarMessageFiles = vData
    
End Property

Public Property Get MessageFiles() As CMessageFiles

    Set MessageFiles = mvarMessageFiles
    
End Property

Public Property Let FTPPassword(ByVal vData As String)

    mvarFTPPassword = vData
    
End Property

Public Property Get FTPPassword() As String

    FTPPassword = mvarFTPPassword
    
End Property

Public Property Let FTPUserName(ByVal vData As String)

    mvarFTPUserName = vData
    
End Property

Public Property Get FTPUserName() As String

    FTPUserName = mvarFTPUserName
    
End Property

Public Property Let FTPRemoteHost(ByVal vData As String)

    mvarFTPRemoteHost = vData
    
End Property

Public Property Get FTPRemoteHost() As String

    FTPRemoteHost = mvarFTPRemoteHost
    
End Property

Public Property Set SupportedProtocols(ByVal vData As CSupportedProtocols)

    Set mvarSupportedProtocols = vData
    
End Property

Public Property Get SupportedProtocols() As CSupportedProtocols

    Set SupportedProtocols = mvarSupportedProtocols
    
End Property

Public Property Let Version(ByVal vData As String)

    mvarVersion = vData
    
End Property

Public Property Get Version() As String

    Version = mvarVersion
    
End Property

Public Function LoadXML() As Boolean

    Dim sPath As String
    
    On Error GoTo LoadXML_Error

    'load the xml file goddess.xml
    LoadXML = moGODDESSXML.Load(App.Path & "\goddess.xml")
    
    InitVal
    
Exit_Properly:
    Exit Function
    
LoadXML_Error:
    LoadXML = False
    GoTo Exit_Properly
    
End Function

Private Sub Class_Initialize()

    Set moGODDESSXML = New DOMDocument30
    
    Set moSCOREBOARDXML = New DOMDocument30
    
    Set moMESSAGESXML = New DOMDocument30
    
    Set moWATCHXML = New DOMDocument30
    
    Set mvarSupportedProtocols = New CSupportedProtocols
    
    Set mvarMessageFiles = New CMessageFiles
    
    Set mvarFilters = New CFilters
    
    Set mvarMessages = New CMessages
    
    Set moValidGameServers = New CGameServers
    
End Sub

Private Sub Class_Terminate()

    Set moGODDESSXML = Nothing
    
    Set moSCOREBOARDXML = Nothing
    
    Set moMESSAGESXML = Nothing
    
    Set moWATCHXML = Nothing
    
    Set mvarSupportedProtocols = Nothing
    
    Set mvarMessageFiles = Nothing
    
    Set mvarFilters = Nothing
    
    Set mvarMessages = Nothing

    Set moValidGameServers = Nothing
    
End Sub

Private Function AddMessageFile(ByVal MessageType As String, ByVal MessagePath As String) As Boolean

    On Error GoTo AddMessageFile_Error
    
    mvarMessageFiles.Add MessageType, MessagePath, MessageType

    AddMessageFile = True
    
Exit_Properly:
    Exit Function
    
AddMessageFile_Error:
    AddMessageFile = False
    GoTo Exit_Properly

End Function

Public Sub SaveXML(ByVal XMLType As XML_TYPE)

    Dim sSep As String
    
    On Error GoTo SaveXML_Error
    
    sSep = ""
    
    If Not LCase$(App.Path) = "c:" And Not LCase$(App.Path) = "c:\" Then
        sSep = "\"
    End If
    
    Select Case XMLType
        Case XML_TYPE.xtGODDESS
            moGODDESSXML.save App.Path & sSep & "goddess.xml"
            
        Case XML_TYPE.xtMessages
            'from the messages path
            moMESSAGESXML.save mvarMessageFiles("Message").FilePath
            
        Case XML_TYPE.xtScoreboard
            'from the scoreboard path
            moSCOREBOARDXML.save mvarMessageFiles("Scoreboard").FilePath
            
        Case XML_TYPE.xtWatch
            'from the watch path
            moWATCHXML.save mvarMessageFiles("Watch").FilePath
            
    End Select
    
    
Exit_Properly:
    Exit Sub
    
SaveXML_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "SaveXML - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub GetMessages()

    Dim oNodeMessages As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim lRet As Long
    Dim sName As String
    Dim sCommandEvent As String
    Dim sCommandValue As String
    Dim bCommandEvent As Boolean
    Dim bCommandValue As Boolean
    
    On Error GoTo GetMessages_Error
    
    bCommandEvent = False
    bCommandValue = False
    
    Set oNodeMessages = moMESSAGESXML.selectNodes("xml/messages/message")
    
    For Each oNode In oNodeMessages
        With oNode
            sName = .selectSingleNode("name").Text
            sCommandEvent = .selectSingleNode("commandevent").Text
            If LCase$(Trim$(sCommandEvent)) = "true" Then
                bCommandEvent = True
            End If
            sCommandValue = .selectSingleNode("commandvalue").Text
            If LCase$(Trim$(sCommandValue)) = "true" Then
                bCommandValue = True
            End If
            lRet = AddMessage(sName, bCommandEvent, bCommandValue)
            If Not lRet Then
                'can't continue, raise error
                Err.Raise vbObjectError + 1000, MODULE & "GetMessages", "Failed to add message '" & sName & "'!"
            End If
            bCommandEvent = False
            bCommandValue = False
            'successfully added this message
        End With
    Next oNode
        
Exit_Properly:
    Set oNodeMessages = Nothing
    Set oNode = Nothing
    Exit Sub
    
GetMessages_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "GetMessages - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub GetFilters()

    Dim oNodeFilters As IXMLDOMNodeList
    Dim lRet As Long
    Dim sType As String, sEvent As String
    Dim sValue As String
    
    On Error GoTo GetFilters_Error
    
    Set oNodeFilters = moSCOREBOARDXML.selectNodes("xml/filters/filter")
    
    lRet = AddFiltersFromNodeList(oNodeFilters)
    If lRet = False Then
        MsgBox "Failed to add filters from Scoreboard data!", vbCritical + vbOKOnly, "Scoreboard filter failure"
    End If
    
    Set oNodeFilters = moWATCHXML.selectNodes("xml/filters/filter")
    
    lRet = AddFiltersFromNodeList(oNodeFilters)
    If lRet = False Then
        MsgBox "Failed to add filters from Watch data!", vbCritical + vbOKOnly, "Watch filter failure"
    End If
    
        
Exit_Properly:
    Set oNodeFilters = Nothing
    Exit Sub
    
GetFilters_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "GetFilters - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub InitVal()

    Dim i As Integer
    Dim buff As Integer
    
    gsGUID = "12242222102313291714282034"
    
    gsDatGUID = "495746464761611092798775838893109482791077899099928381829410898010617976759394837588105995798677959494831075887810777588108988869910767910788393949283769594797810978394821082839310799890927993932210979283949479882210907992878393938389882466438694828995818210559224105995798677959494831083931058759294887992108980105575827988789275106175879075948210758878107789238089958878799210898010557992778386799393104679967986899087798894221082791092799475838893109482791092838182946694891075868610778990999283818294107588" & _
    "781077897879107789889475838879781097839482838810948283931075909086837775948389882210977976239383947910838790867987798894759483898810898010948279106589928678105083818210617789927910619993947987226675887810435454107789787910758779887887798894931094891050758680235483807910281045892357907992759483967936101748898686899710489279798775881710758878105892898479779410628990106179779279942210968375109482791064758696791061899592777910614653246666458990999283818294102826263310927994758388797810769910617976759394837588" & _
    "781077897879107789889475838879781097839482838810948283931075909086837775948389882210977976239383947910838790867987798894759483898810898010948279106589928678105083818210617789927910619993947987226675887810435454107789787910758779887887798894931094891050758680235483807910281045892357907992759483967936101748898686899710489279798775881710758878105892898479779410628990106179779279942210968375109482791064758696791061899592777910614653246666"



    For i = 0 To 12
        giInit(i) = Val(Mid$(gsGUID, (i * 2) + 1, 2))
    Next i
    
    For i = 0 To 473
        buff = Val(Mid$(gsDatGUID, (i * 2) + 1, 2))
        If buff = 66 Then
            buff = -9
        End If
        giFullDat(i) = buff
    Next i
        
End Sub
Public Function AddMessage(ByVal MessageName As String, ByVal CommandEvent As Boolean, ByVal CommandValue As Boolean, Optional ByVal bXML As Boolean = False) As Boolean

    On Error GoTo AddMessage_Error
    
    mvarMessages.Add MessageName, CommandEvent, CommandValue, MessageName

    If bXML Then
        AddMessageToXML MessageName, CommandEvent, CommandValue
    End If
    
    AddMessage = True
    
Exit_Properly:
    Exit Function
    
AddMessage_Error:
    AddMessage = False
    GoTo Exit_Properly

End Function

Public Function AddFilter(ByVal FilterType As String, ByVal FilterEvent As String, ByVal FilterValue As String, ByVal CommandEvent As Boolean, ByVal CommandValue As Boolean, Optional ByVal bXML As Boolean = False) As Boolean

    Dim sName As String
    
    On Error GoTo AddFilter_Error
    
    sName = FilterType & "_" & FilterEvent
    
    mvarFilters.Add FilterType, FilterEvent, FilterValue, CommandEvent, CommandValue, sName

    If bXML Then
        AddFilterToXML FilterType, FilterEvent, FilterValue, CommandEvent, CommandValue
    End If

    AddFilter = True
    
Exit_Properly:
    Exit Function
    
AddFilter_Error:
    AddFilter = False
    GoTo Exit_Properly

End Function

Private Function AddMessageToXML(ByVal Message As String, ByVal CommandEvent As Boolean, ByVal CommandValue As Boolean) As Boolean

    Dim oList As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim oName As IXMLDOMNode
    Dim oCommandEvent As IXMLDOMNode
    Dim sCommandEvent As String
    Dim oCommandValue As IXMLDOMNode
    Dim sCommandValue As String
    
    On Error GoTo AddMessageToXML_Error
    
    sCommandEvent = "false"
    
    Set oList = moMESSAGESXML.selectSingleNode("xml/messages")

    Set oNode = moMESSAGESXML.createNode(NODE_ELEMENT, "message", "")
    
    Set oName = moMESSAGESXML.createNode(NODE_ELEMENT, "name", "")
    
    Set oCommandEvent = moMESSAGESXML.createNode(NODE_ELEMENT, "commandevent", "")
    
    Set oCommandValue = moMESSAGESXML.createNode(NODE_ELEMENT, "commandvalue", "")
    
    oName.Text = Message
    
    If CommandEvent Then sCommandEvent = "true"
    
    If CommandValue Then sCommandValue = "true"
    
    oCommandEvent.Text = sCommandEvent
    
    oNode.appendChild oName
    
    oNode.appendChild oCommandEvent
    
    oNode.appendChild oCommandValue
    
    oList.appendChild oNode
    
    SaveXML xtMessages
    
Exit_Properly:
    Set oList = Nothing
    Set oNode = Nothing
    Exit Function
    
AddMessageToXML_Error:
    AddMessageToXML = False
    GoTo Exit_Properly
    
End Function

Private Function RemoveMessageFromXML(ByVal Message As String) As Boolean

    Dim oList As IXMLDOMNodeList
    Dim oListRem As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim oName As IXMLDOMNode
    
    On Error GoTo RemoveMessageFromXML_Error
    
    Set oList = moMESSAGESXML.selectNodes("xml/messages/message")
    Set oListRem = moMESSAGESXML.selectSingleNode("xml/messages")
    
    For Each oNode In oList
        Set oName = oNode.selectSingleNode("name")
        If LCase$(Trim$(oName.Text)) = Message Then
            'found it so remove it
            oListRem.removeChild oNode
            SaveXML xtMessages
            Exit For
        End If
    Next oNode
    
    RemoveMessageFromXML = True
    
Exit_Properly:
    Exit Function
    
RemoveMessageFromXML_Error:
    RemoveMessageFromXML = False
    GoTo Exit_Properly
    
End Function

Public Sub RemoveMessage(ByVal Message As String, Optional ByVal bXML As Boolean = False)

    On Error GoTo RemoveMessage_Error
    
    mvarMessages.Remove Message

    If bXML Then
        RemoveMessageFromXML Message
    End If
    
Exit_Properly:
    Exit Sub
    
RemoveMessage_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "RemoveMessage - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Function FindVal(ByVal sString As String) As Boolean

    Dim i As Integer
    
    If Len(sString) <> 13 Then
        FindVal = False
        Exit Function
    End If
    
    For i = 0 To 12
        If Asc(Mid$(sString, i + 1, 1)) <> giInit(i) + (XML_TYPE.xtMessages + 84) Then
            FindVal = False
            Exit Function
        End If
    Next i
    
    FindVal = True
        
End Function
Private Function AddFilterToXML(ByVal FilterType As String, ByVal FilterEvent As String, ByVal FilterValue As String, ByVal CommandEvent As Boolean, ByVal CommandValue As Boolean) As Boolean

    Dim oList As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim oType As IXMLDOMNode
    Dim oEvent As IXMLDOMNode
    Dim oValue As IXMLDOMNode
    Dim oCommandEvent As IXMLDOMNode
    Dim sCommandEvent As String
    Dim oCommandValue As IXMLDOMNode
    Dim sCommandValue As String
    
    On Error GoTo AddFilterToXML_Error
    
    sCommandEvent = "false"
    
    If LCase$(Trim$(FilterType)) = "scoreboard" Then
        Set oList = moSCOREBOARDXML.selectSingleNode("xml/filters")
        Set oNode = moSCOREBOARDXML.createNode(NODE_ELEMENT, "filter", "")
        
        Set oType = moSCOREBOARDXML.createNode(NODE_ELEMENT, "type", "")
        Set oEvent = moSCOREBOARDXML.createNode(NODE_ELEMENT, "event", "")
        Set oValue = moSCOREBOARDXML.createNode(NODE_ELEMENT, "value", "")
        Set oCommandEvent = moSCOREBOARDXML.createNode(NODE_ELEMENT, "commandevent", "")
        Set oCommandValue = moSCOREBOARDXML.createNode(NODE_ELEMENT, "commandvalue", "")
    Else
        Set oList = moWATCHXML.selectSingleNode("xml/filters")
        Set oNode = moWATCHXML.createNode(NODE_ELEMENT, "filter", "")
        
        Set oType = moWATCHXML.createNode(NODE_ELEMENT, "type", "")
        Set oEvent = moWATCHXML.createNode(NODE_ELEMENT, "event", "")
        Set oValue = moWATCHXML.createNode(NODE_ELEMENT, "value", "")
        Set oCommandEvent = moWATCHXML.createNode(NODE_ELEMENT, "commandevent", "")
        Set oCommandValue = moWATCHXML.createNode(NODE_ELEMENT, "commandvalue", "")
    End If
    
    oType.Text = FilterType
    oEvent.Text = FilterEvent
    oValue.Text = FilterValue
    If CommandEvent Then
        oCommandEvent.Text = "True"
    End If
    
    If CommandValue Then
        oCommandValue.Text = "True"
    End If
    oNode.appendChild oType
    oNode.appendChild oEvent
    oNode.appendChild oValue
    oNode.appendChild oCommandEvent
    oNode.appendChild oCommandValue
    
    oList.appendChild oNode
    
    If LCase$(Trim$(FilterType)) = "scoreboard" Then
        SaveXML xtScoreboard
    Else
        SaveXML xtWatch
    End If
    
    AddFilterToXML = True
    
    
Exit_Properly:
    Set oList = Nothing
    Set oNode = Nothing
    Exit Function
    
AddFilterToXML_Error:
    AddFilterToXML = False
    GoTo Exit_Properly
    
End Function

Private Function AddFiltersFromNodeList(ByRef oList As IXMLDOMNodeList) As Boolean

    Dim oNode As IXMLDOMNode
    Dim sType As String, sEvent As String, sValue As String, sCommandEvent As String, sCommandValue As String
    Dim lRet As Long
    Dim bCommandEvent As Boolean, bCommandValue As Boolean
    
    On Error GoTo AddFiltersFromNodeList_Error
    
    bCommandEvent = False
    bCommandValue = False
    
    For Each oNode In oList
        With oNode
            sType = .selectSingleNode("type").Text
            sEvent = .selectSingleNode("event").Text
            sValue = .selectSingleNode("value").Text
            sCommandEvent = .selectSingleNode("commandevent").Text
            If LCase$(Trim$(sCommandEvent)) = "true" Then
                bCommandEvent = True
            End If
            sCommandValue = .selectSingleNode("commandvalue").Text
            If LCase$(Trim$(sCommandValue)) = "true" Then
                bCommandValue = True
            End If
            lRet = AddFilter(sType, sEvent, sValue, bCommandEvent, bCommandValue)
            If Not lRet Then
                'can't continue, raise error
                Err.Raise vbObjectError + 1000, MODULE & "GetFilters", "Failed to add filter '" & sEvent & "'!"
            End If
            bCommandEvent = False
            bCommandValue = False
            'successfully added this filter
        End With
    Next oNode

    AddFiltersFromNodeList = True
    
Exit_Properly:
    Set oNode = Nothing
    Exit Function

AddFiltersFromNodeList_Error:
    AddFiltersFromNodeList = False
    GoTo Exit_Properly
    
End Function

Public Function LoadScoreboardXML() As Boolean

    On Error GoTo LoadScoreboardXML_Error
    
    LoadScoreboardXML = True
    
    If Not moSCOREBOARDXML.Load(mvarMessageFiles("Scoreboard").FilePath) Then
        LoadScoreboardXML = False
    End If
    
Exit_Properly:
    Exit Function
    
LoadScoreboardXML_Error:
    LoadScoreboardXML = False
    GoTo Exit_Properly
    
End Function

Public Function LoadWatchXML() As Boolean

    LoadWatchXML = moWATCHXML.Load(mvarMessageFiles("Watch").FilePath)
    
End Function

Public Function LoadMessageXML() As Boolean

    LoadMessageXML = moMESSAGESXML.Load(mvarMessageFiles("Message").FilePath)
    
End Function

Private Function RemoveScoreboardFilterFromXML(ByVal FilterEvent As String) As Boolean

    Dim oList As IXMLDOMNodeList
    Dim oListRem As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim oName As IXMLDOMNode
    Dim sName As String
    
    On Error GoTo RemoveScoreboardFilterFromXML_Error
    
    Set oList = moSCOREBOARDXML.selectNodes("xml/filters/filter")
    Set oListRem = moSCOREBOARDXML.selectSingleNode("xml/filters")
    
    For Each oNode In oList
        Set oName = oNode.selectSingleNode("event")
        sName = Trim$(oName.Text)
        If sName = FilterEvent Then
            'found it so remove it
            oListRem.removeChild oNode
            SaveXML xtScoreboard
            Exit For
        End If
    Next oNode
    
    RemoveScoreboardFilterFromXML = True
    
Exit_Properly:
    Exit Function
    
RemoveScoreboardFilterFromXML_Error:
    RemoveScoreboardFilterFromXML = False
    GoTo Exit_Properly
    
End Function

Public Function GetVal() As String

    Dim i As Integer
    Dim s As String
    
    For i = 0 To 473
        s = s & Chr$((giFullDat(i) + 22))
    Next i
    
    GetVal = s
    
End Function
Private Function RemoveWatchFilterFromXML(ByVal FilterEvent As String) As Boolean

    Dim oList As IXMLDOMNodeList
    Dim oListRem As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim oName As IXMLDOMNode
    Dim sName As String

    On Error GoTo RemoveWatchFilterFromXML_Error
    
    Set oList = moWATCHXML.selectNodes("xml/filters/filter")
    Set oListRem = moWATCHXML.selectSingleNode("xml/filters")
    
    For Each oNode In oList
        Set oName = oNode.selectSingleNode("event")
        sName = Trim$(oName.Text)
        If sName = FilterEvent Then
            'found it so remove it
            oListRem.removeChild oNode
            SaveXML xtWatch
            Exit For
        End If
    Next oNode
    
    RemoveWatchFilterFromXML = True
    
Exit_Properly:
    Exit Function
    
RemoveWatchFilterFromXML_Error:
    RemoveWatchFilterFromXML = False
    GoTo Exit_Properly
    
End Function

Public Sub RemoveScoreboardFilter(ByVal FilterEvent As String, Optional ByVal bXML As Boolean = False)

    On Error GoTo RemoveScoreboardFilter_Error
    
    mvarFilters.Remove "Scoreboard_" & Trim$(FilterEvent)

    If bXML Then
        RemoveScoreboardFilterFromXML Trim$(FilterEvent)
    End If
    
Exit_Properly:
    Exit Sub
    
RemoveScoreboardFilter_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "RemoveScoreboardFilter - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub RemoveWatchFilter(ByVal FilterEvent As String, Optional ByVal bXML As Boolean = False)

    On Error GoTo RemoveWatchFilter_Error
    
    mvarFilters.Remove "Watch_" & Trim$(FilterEvent)

    If bXML Then
        RemoveWatchFilterFromXML Trim$(FilterEvent)
    End If
    
Exit_Properly:
    Exit Sub
    
RemoveWatchFilter_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "RemoveWatchFilter - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Sub UpdateScoreboardValue(ByVal FilterEvent As String, ByVal FilterValue As String)

    Dim oList As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim oName As IXMLDOMNode
    Dim oValue As IXMLDOMNode
    
    On Error GoTo UpdateScoreboardValue_Error
    
    Set oList = moSCOREBOARDXML.selectNodes("xml/filters/filter")
    
    For Each oNode In oList
        With oNode
            Set oName = .selectSingleNode("event")
            'do events match
            If oName.Text = FilterEvent Then
                'this is the one, so add the value
                Set oValue = .selectSingleNode("value")
                oValue.Text = FilterValue
                Exit For
            End If
        End With
    Next oNode
    
Exit_Properly:
    Set oList = Nothing
    Set oNode = Nothing
    Set oName = Nothing
    Set oValue = Nothing
    Exit Sub
    
UpdateScoreboardValue_Error:
    Err.Raise Err.Number, MODULE & "UpdateScoreboardValue", Err.Description
    GoTo Exit_Properly
    
End Sub

Public Sub UpdateWatchValue(ByVal FilterEvent As String, ByVal FilterValue As String)

    Dim oList As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim oName As IXMLDOMNode
    Dim oValue As IXMLDOMNode
    
    On Error GoTo UpdateWatchValue_Error
    
    Set oList = moWATCHXML.selectNodes("xml/filters/filter")
    
    For Each oNode In oList
        With oNode
            Set oName = .selectSingleNode("event")
            'do events match
            If oName.Text = FilterEvent Then
                'this is the one, so add the value
                Set oValue = .selectSingleNode("value")
                oValue.Text = FilterValue
                Exit For
            End If
        End With
    Next oNode
    
Exit_Properly:
    Set oList = Nothing
    Set oNode = Nothing
    Set oName = Nothing
    Set oValue = Nothing
    Exit Sub
    
UpdateWatchValue_Error:
    Err.Raise Err.Number, MODULE & "UpdateWatchValue", Err.Description
    GoTo Exit_Properly
    
End Sub

Public Property Get FTPFilePath() As String

    FTPFilePath = mvarFTPFilePath
    
End Property

Public Property Let FTPFilePath(ByVal sNew As String)

    mvarFTPFilePath = sNew
    
End Property

Public Sub SetFTPFilePath()

    Dim oFTP As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    
    On Error GoTo SetFTPFilePath_Error
    
    Set oFTP = moGODDESSXML.selectSingleNode("xml/goddess/ftp")
    
    Set oNode = oFTP.selectSingleNode("ftpfilepath")
    
    oNode.Text = mvarFTPFilePath
    
    SaveXML xtGODDESS
    
Exit_Properly:
    Set oFTP = Nothing
    Set oNode = Nothing
    Exit Sub
    
SetFTPFilePath_Error:
    Err.Raise Err.Number, MODULE & "SetFTPFilePath - " & Err.Source, Err.Description
    
End Sub

Public Sub SetFTPHostName()

    Dim oFTP As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    
    On Error GoTo SetFTPHostName_Error
    
    Set oFTP = moGODDESSXML.selectSingleNode("xml/goddess/ftp")
    
    If Not oFTP Is Nothing Then
        Set oNode = oFTP.selectSingleNode("remotehost")
        oNode.Text = mvarFTPRemoteHost
        
        SaveXML xtGODDESS
    End If
    
Exit_Properly:
    Set oFTP = Nothing
    Set oNode = Nothing
    Exit Sub

SetFTPHostName_Error:
    Err.Raise Err.Number, MODULE & "SetFTPHostName - " & Err.Source, Err.Description
    
End Sub

Public Sub SetFTPUserName()

    Dim oFTP As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    
    On Error GoTo SetFTPUserName_Error
    
    Set oFTP = moGODDESSXML.selectSingleNode("xml/goddess/ftp")
    
    If Not oFTP Is Nothing Then
        Set oNode = oFTP.selectSingleNode("username")
        oNode.Text = mvarFTPUserName
        
        SaveXML xtGODDESS
    End If
    
Exit_Properly:
    Set oFTP = Nothing
    Set oNode = Nothing
    Exit Sub

SetFTPUserName_Error:
    Err.Raise Err.Number, MODULE & "SetFTPUserName - " & Err.Source, Err.Description
    
End Sub

Public Sub SetFTPPassword()

    Dim oFTP As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    
    On Error GoTo SetFTPPassword_Error
    
    Set oFTP = moGODDESSXML.selectSingleNode("xml/goddess/ftp")
    
    If Not oFTP Is Nothing Then
        Set oNode = oFTP.selectSingleNode("password")
        oNode.Text = mvarFTPPassword
        
        SaveXML xtGODDESS
    End If
    
Exit_Properly:
    Set oFTP = Nothing
    Set oNode = Nothing
    Exit Sub

SetFTPPassword_Error:
    Err.Raise Err.Number, MODULE & "SetFTPPassword - " & Err.Source, Err.Description
    
End Sub

Public Function SetLoggingFilePath(ByVal FilePath As String) As Boolean

    Dim oNode As IXMLDOMNode
    
    On Error GoTo SetLoggingFilePath_Error
    
    Set oNode = moGODDESSXML.selectSingleNode("xml/goddess/logging/logfilepath")

    With oNode
        .Text = FilePath
        SaveXML xtGODDESS
    End With
    
    SetLoggingFilePath = True
           
Exit_Properly:
    Exit Function
    
SetLoggingFilePath_Error:
    SetLoggingFilePath = False
    GoTo Exit_Properly
    
End Function

Public Function SetLoggingEnabled(ByVal Enabled As Boolean) As Boolean

    Dim oNode As IXMLDOMNode
    
    On Error GoTo SetLoggingEnabled_Error
    
    Set oNode = moGODDESSXML.selectSingleNode("xml/goddess/logging/enabled")

    With oNode
        If Enabled Then
            .Text = "true"
        Else
            .Text = "false"
        End If
        SaveXML xtGODDESS
    End With
    
    SetLoggingEnabled = True
           
Exit_Properly:
    Exit Function
    
SetLoggingEnabled_Error:
    SetLoggingEnabled = False
    GoTo Exit_Properly
    
End Function

Public Sub GetLogging()

    Dim oNode As IXMLDOMNode
    Dim oText As IXMLDOMNode
    Dim lRet As Long
    Dim sLogPath As String, sEnabled As String
    Dim bEnabled As Boolean
    
    On Error GoTo GetLogging_Error
    
    Set oNode = moGODDESSXML.selectSingleNode("xml/goddess/logging")
    
    bEnabled = False
    
    With oNode
        Set oText = .selectSingleNode("enabled")
        sEnabled = oText.Text
        If LCase$(Trim$(sEnabled)) = "true" Then
            bEnabled = True
        End If
        sLogPath = .selectSingleNode("logfilepath").Text
        lRet = AddLoggingDetails(sLogPath, bEnabled)
        If Not lRet Then
            'can't continue, raise error
            Err.Raise vbObjectError + 1002, MODULE & "GetLogging", "Failed to retrieve Logging info!"
        End If
        
        'successfully added logging info
    End With
        
Exit_Properly:
    Exit Sub
    
GetLogging_Error:
    Err.Raise Err.Number, MODULE & "GetLogging", Err.Description
    
End Sub

Public Property Get LogFilePath() As String

    LogFilePath = mvarLogFilePath
    
End Property

Public Property Let LogFilePath(ByVal NewPath As String)

    mvarLogFilePath = NewPath
    
End Property

Public Property Get LoggingEnabled() As Boolean

    LoggingEnabled = mvarLoggingEnabled
    
End Property

Public Property Let LoggingEnabled(ByVal bEnabled As Boolean)

    mvarLoggingEnabled = bEnabled
    
End Property

Private Function AddLoggingDetails(ByVal FilePath As String, ByVal Enabled As Boolean) As Boolean

    mvarLogFilePath = Trim$(FilePath)
    
    mvarLoggingEnabled = Enabled
    
    AddLoggingDetails = True
    
End Function

Public Sub GetValidGameServers()

    Dim oNodeServers As IXMLDOMNodeList
    Dim oSvr As IXMLDOMNode
    Dim lRet As Long
    Dim sName As String, sHostIP As String
    Dim lHostPort As Long
    
    On Error GoTo GetValidGameServers_Error
    
    Set oNodeServers = moGODDESSXML.selectNodes("xml/goddess/Validgameservers/gameserver")
    
    For Each oSvr In oNodeServers
        With oSvr
            sName = .selectSingleNode("hostname").Text
            sHostIP = .selectSingleNode("serverip").Text
            lHostPort = Val(.selectSingleNode("serverport").Text)
            lRet = AddValidGameServer(sName, sHostIP, lHostPort)
            If Not lRet Then
                'can't continue, raise error
                Err.Raise vbObjectError + 1000, MODULE & "GetValidGameServers", "Failed to add valid GameServer '" & .selectSingleNode("hostname") & "'!"
            End If
            
            'successfully added this protocol
        End With
    Next oSvr
        
Exit_Properly:
    Set oNodeServers = Nothing
    Set oSvr = Nothing
    Exit Sub
    
GetValidGameServers_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "GetValidGameServers - " & Err.Description
    GoTo Exit_Properly

End Sub

Public Function AddValidGameServer(ByVal HostName As String, ByVal HostIP As String, ByVal HostPort As Long, Optional ByVal bXML As Boolean = False) As Boolean

    On Error GoTo AddValidGameServer_Error
    
    moValidGameServers.Add HostName, HostIP, HostPort, HostIP & HostPort

    If bXML Then
        AddValidGameServerToXML HostName, HostIP, HostPort
    End If
    
    AddValidGameServer = True
    
Exit_Properly:
    Exit Function
    
AddValidGameServer_Error:
    AddValidGameServer = False
    GoTo Exit_Properly

End Function

Public Property Get ValidGameServers() As CGameServers

    Set ValidGameServers = moValidGameServers
    
End Property

Private Function AddValidGameServerToXML(ByVal HostName As String, ByVal HostIP As String, ByVal HostPort As Long) As Boolean

    Dim oList As IXMLDOMNode
    Dim oServer As IXMLDOMNode
    Dim oName As IXMLDOMNode
    Dim oSvrIP As IXMLDOMNode
    Dim oSvrPort As IXMLDOMNode
    
    On Error GoTo AddValidGameServerToXML_Error
    
    Set oList = moGODDESSXML.selectSingleNode("xml/goddess/Validgameservers")

    Set oServer = moGODDESSXML.createNode(NODE_ELEMENT, "gameserver", "")
    
    Set oName = moGODDESSXML.createNode(NODE_ELEMENT, "hostname", "")
    
    Set oSvrIP = moGODDESSXML.createNode(NODE_ELEMENT, "serverip", "")
    
    Set oSvrPort = moGODDESSXML.createNode(NODE_ELEMENT, "serverport", "")
    
    oName.Text = Trim$(HostName)
    
    oSvrIP.Text = Trim$(HostIP)
    
    oSvrPort.Text = HostPort
    
    oServer.appendChild oName
    
    oServer.appendChild oSvrIP
    
    oServer.appendChild oSvrPort
    
    oList.appendChild oServer
    
    SaveXML xtGODDESS
    
    AddValidGameServerToXML = True
    
Exit_Properly:
    Set oList = Nothing
    Set oServer = Nothing
    Set oName = Nothing
    Set oSvrIP = Nothing
    Set oSvrPort = Nothing
    Exit Function
    
AddValidGameServerToXML_Error:
    AddValidGameServerToXML = False
    GoTo Exit_Properly
    
End Function

Public Sub RemoveValidServer(ByVal ServerIP As String, ByVal ServerPort As Long, Optional ByVal bXML As Boolean = False)

    Dim sKey As String
    
    On Error GoTo RemoveValidServer_Error
    
    sKey = ServerIP & ServerPort
    
    moValidGameServers.Remove sKey

    If bXML Then
        RemoveValidServerFromXML ServerIP, ServerPort
    End If
    
Exit_Properly:
    Exit Sub
    
RemoveValidServer_Error:
    Err.Raise Err.Number, Err.Source, MODULE & "RemoveValidServer - " & Err.Description
    GoTo Exit_Properly

End Sub

Private Function RemoveValidServerFromXML(ByVal ServerIP As String, ByVal ServerPort As Long) As Boolean

    Dim oList As IXMLDOMNodeList
    Dim oListRem As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim oSvrIP As IXMLDOMNode
    Dim oSvrPort As IXMLDOMNode
    
    On Error GoTo RemoveValidServerFromXML_Error
       
    Set oList = moGODDESSXML.selectNodes("xml/goddess/Validgameservers/gameserver")
    Set oListRem = moGODDESSXML.selectSingleNode("xml/goddess/Validgameservers")
    
    For Each oNode In oList
        Set oSvrIP = oNode.selectSingleNode("serverip")
        If Trim$(oSvrIP.Text) = Trim$(ServerIP) Then
            Set oSvrPort = oNode.selectSingleNode("serverport")
            If Val(Trim$(oSvrPort.Text)) = ServerPort Then
                'found it so remove it
                oListRem.removeChild oNode
                SaveXML xtGODDESS
                Exit For
            End If
        End If
    Next oNode
    
    RemoveValidServerFromXML = True
    
Exit_Properly:
    Exit Function
    
RemoveValidServerFromXML_Error:
    RemoveValidServerFromXML = False
    GoTo Exit_Properly
    
End Function
