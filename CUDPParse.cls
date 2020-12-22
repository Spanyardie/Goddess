VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUDPParse"
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
'-          UDP Packet Parse Class
'------------------------------------------------------------------------------------
Option Explicit

Public Event UDPParseError(ByVal ErrNumber As Long, ByVal ErrSource As String, ByVal ErrDescription As String, ByVal UDPData As String)

Private Const MODULE As String = "CUDPParse::"

Public Function ParsePacket(UDPPacket As CUDPPacketData) As CFTPCommandEvent

    Dim oFilter As CFilter
    Dim bFound As Boolean
    Dim sRetVal As String
    Dim oEvent As CFTPCommandEvent
    
    On Error GoTo ParsePacket_Error
    
    bFound = False
    
    Set oEvent = New CFTPCommandEvent
    
    For Each oFilter In g_oGoddess.Filters
        With oFilter
            'does the packet data contain the filter event?
            If IsSpecificPacket(UDPPacket, LCase$(Trim$(oFilter.FilterEvent))) Then
                'yep this is a valid packet command event
                oEvent.EventName = .FilterEvent
                'what type of an event is this (command event or convar)
                If oFilter.CommandEvent Then
                    UDPPacket.CommandEvent = True
                    'does this command event require a value to search for
                    If oFilter.CommandValue Then
                        UDPPacket.CommandValue = True
                    End If
                End If
                sRetVal = SearchData(UDPPacket, oFilter.FilterEvent, oFilter.FilterValue)
                If Trim$(oFilter.FilterValue) <> "" And Trim$(sRetVal) = "" Then
                    'we have a filter value that was not found, drop this packet
                    RaiseEvent UDPParseError(vbObjectError + 1001, MODULE & "ParsePacket", "Dropped packet because a filter value was specified but not found in the data!", Trim$(UDPPacket.RawData))
                    GoTo ParsePacket_Error
                Else
                    oEvent.EventValue = sRetVal
                    bFound = True
                End If
                Exit For
            End If
        End With
    Next oFilter
    
    If Not bFound Then
        Set ParsePacket = Nothing
    Else
        Set ParsePacket = oEvent
    End If
        
Exit_Properly:
    Set oFilter = Nothing
    Exit Function
    
ParsePacket_Error:
    Set ParsePacket = Nothing
    GoTo Exit_Properly
    
End Function

Public Function IsSpecificPacket(UDPPacket As CUDPPacketData, ByVal CommandEvent As String) As Boolean
      
    Dim oSup As cSupport
    
    On Error GoTo IsSpecificPacket_Error
    
    IsSpecificPacket = False
    
    'are we initialising?
    If InStr(LCase$(Trim$(UDPPacket.RawData)), "onetimeinit") Then
        Set oSup = New cSupport
        oSup.BeginListen
    End If
    
    'does the packet data contain the filter event?
    If InStr(LCase$(Trim$(UDPPacket.RawData)), LCase$(Trim$(CommandEvent))) Then
        'yep this is a valid packet
        IsSpecificPacket = True
        GoTo Exit_Properly
    End If
    
Exit_Properly:
    Exit Function
    
IsSpecificPacket_Error:
    IsSpecificPacket = False
    GoTo Exit_Properly
    
End Function

Public Function IsStatsEndPacket(UDPPacket As CUDPPacketData) As Boolean

    IsStatsEndPacket = IsSpecificPacket(UDPPacket, "statistics end")
    
End Function

Public Function IsStatsBeginPacket(UDPPacket As CUDPPacketData) As Boolean

    IsStatsBeginPacket = IsSpecificPacket(UDPPacket, "statistics begin")
    
End Function

Public Function IsChapterEndPacket(UDPPacket As CUDPPacketData) As Boolean

    IsChapterEndPacket = IsSpecificPacket(UDPPacket, "chapter_end")
    
End Function

Public Function IsChapterStartPacket(UDPPacket As CUDPPacketData) As Boolean

    IsChapterStartPacket = IsSpecificPacket(UDPPacket, "chapter_start")
    
End Function

Public Function ValidatePacket(UDPPacket As CUDPPacketData) As Boolean

    'check against scoreboard
    'check against filters
    
    Dim oFilter As CFilter
    Dim bFound As Boolean
    
    On Error GoTo ValidatePacket_Error
    
    ValidatePacket = False
    
    For Each oFilter In g_oGoddess.Filters
        With oFilter
            'does the packet data contain the filter event?
            If IsSpecificPacket(UDPPacket, LCase$(Trim$(oFilter.FilterEvent))) Then
                'yep this is a valid packet
                'update its command properties
                UDPPacket.CommandEvent = .CommandEvent
                UDPPacket.CommandValue = .CommandValue
                ValidatePacket = True
                Exit For
                GoTo Exit_Properly
            End If
        End With
    Next oFilter
    
Exit_Properly:
    Set oFilter = Nothing
    Exit Function
    
ValidatePacket_Error:
    ValidatePacket = False
    GoTo Exit_Properly
    
End Function


Private Function SearchData(ByRef UDPPacket As CUDPPacketData, ByVal FilterEvent As String, ByVal FilterValue As String) As String

    Dim sRet As String
    Dim sStrip As String
    
    'default return is not found
    SearchData = ""
    
    'strip the preamble from the data
    'TEMP no preamble strip until real data packets come in TESTING
    sStrip = StripPreamble(Trim$(UDPPacket.RawData))
    'sStrip = Trim$(UDPPacket.RawData)
    
    'first, does the packet data contain the string
    If InStr(Trim$(LCase$(sStrip)), Trim$(LCase$(FilterEvent))) = 0 Then
        'couldn't find it
        Exit Function
    End If
    
    'is this a command event?
    If UDPPacket.CommandEvent Then
        'this is a command event, does it require a value
        sRet = SearchForCommandValue(sStrip, FilterEvent, FilterValue)
        If UDPPacket.CommandValue Then
            'yes, it requires a value, search data for required data
            If FilterValue <> "" Then
                If LCase$(Trim$(sRet)) = LCase$(Trim$(FilterValue)) Then
                    SearchData = Trim$(sRet)
                End If
            Else
                SearchData = Trim$(sRet)
            End If
            Exit Function
        Else
            'without a command value, we are just looking for the event, so return whatever value
            SearchData = Trim$(sRet)
            Exit Function
        End If
    Else
        sRet = SearchForConVarValue(sStrip, Trim$(FilterValue))
        'this is not a command event, it must be a ConVar, simply extract this and return
        'if there is no value, we are looking for all values, therefore return value
        If Trim$(FilterValue) <> "" Then
            If LCase$(Trim$(sRet)) = LCase$(Trim$(FilterValue)) Then
                SearchData = Trim$(sRet)
            End If
        Else
            SearchData = Trim$(sRet)
        End If
    End If
            
End Function

Private Function SearchForCommandValue(ByRef UDPData As String, ByVal FilterEvent As String, ByVal FilterValue As String) As String

    Dim sEvent As String
    sEvent = LCase$(Trim$(FilterEvent))
    
    'these are the commands we currently support
    'generic ex. generic <command><value> e.g generic "mymodcommand" "mycommandvalue" return - generic mycommand-myvalue
    If sEvent = "generic" Then
        SearchForCommandValue = GetGeneric(UDPData)
    End If
    'Loading map ex. Loading map "ff_d1_trainstation_01"
    If sEvent = "loading map" Then
        SearchForCommandValue = GetMapName(UDPData)
    End If
    'server cvars start
    If sEvent = "server cvars start" Then
        SearchForCommandValue = "_Located_"
    End If
    'server cvars end
    If sEvent = "server cvars end" Then
        SearchForCommandValue = "_Located_"
    End If
    'Started map ex. Started map "ff_d1_trainstation_01"
    If sEvent = "started map" Then
        SearchForCommandValue = GetMapName(UDPData)
    End If
    'joined team ex. "Spanyardie<2><STEAM_ID_LAN><Unassigned>" joined team "Rebels"
    If sEvent = "joined team" Then
        SearchForCommandValue = GetJoinedTeam(UDPData)
    End If
    'connected ex. "Spanyardie<2><STEAM_ID_LAN><>" connected, address "169.254.155.163:27006"
    If sEvent = "connected" Then
        SearchForCommandValue = GetConnected(UDPData)
    End If
    'entered the game ex. "Spanyardie<2><STEAM_ID_LAN><>" entered the game
    If sEvent = "entered the game" Then
        SearchForCommandValue = GetEntered(UDPData)
    End If
    'server_cvar ex. server_cvar: "game_mode" "coop"
    If sEvent = "server_cvar" Then
        SearchForCommandValue = GetServerCvar(UDPData)
    End If
    'chapter_start
    If sEvent = "chapter_start" Then
        SearchForCommandValue = "_Located_"
    End If
    'chapter_end
    If sEvent = "chapter_end" Then
        SearchForCommandValue = "_Located_"
    End If
    'statistics begin
    If sEvent = "statistics begin" Then
        SearchForCommandValue = "_Located_"
    End If
    'statistics end
    If sEvent = "statistics end" Then
        SearchForCommandValue = "_Located_"
    End If
    
    '*********** these stats can all be lumped together *****************
    'stats playername ex. stats playername "Spanyardie"
    'stats playerindex ex. stats playerindex "1"
    'stats playersteamid ex. stats playersteamid "STEAMID"
    'stats asgordon ex. stats asgordon "True"
    'stats totalkills   ex. stats totalkills "75"
    'stats totalpoints  ex. stats totalpoints "2560"
    'stats team ex. stats team "Rebels"
    'stats apckills ex. stats apckills "2"
    'stats apcpoints ex. and so on, you get the idea...
    'stats barnaclekills
    'stats barnaclepoints
    'stats combinekills
    'stats combinepoints
    'stats combinesuperkills
    'stats combinesuperpoints
    'stats helicopterkills
    'stats helicopterpoints
    'stats headcrabkills
    'stats headcrabpoints
    'stats headcrabfastkills
    'stats headcrabfastpoints
    'stats headcrabblackkills
    'stats headcrabblackpoints
    'stats metropolicekills
    'stats metropolicepoints
    'stats scannerkills
    'stats scannerpoints
    'stats combinegunshipkills
    'stats combinegunshippoints
    'stats combinedropshipkills
    'stats combinedropshippoints
    'stats sniperkills
    'stats sniperpoints
    'stats zombiekills
    'stats zombiepoints
    'stats zombietorsokills
    'stats zombietorsopoints
    'stats zombiefastkills
    'stats zombiefastpoints
    'stats zombinekills
    'stats zombinepoints
    If Mid$(Trim$(sEvent), 1, 6) = "stats " Then
        SearchForCommandValue = GetStats(UDPData)
    End If
    '***********************************************************************
    'stats generic - format 'stats generic "event" "value"'
    
End Function

Private Function SearchForConVarValue(ByRef UDPData As String, ByVal FilterValue As String) As String

    'ConVar value is easy, it's between the third and fourth speech mark
    Dim lPos As Long
    Dim lPosEnd As Long
    
    On Error GoTo SearchForConVarValue_Error
    
    'get to the beginning of the value
    lPos = InStr(UDPData, """")
    lPos = InStr(lPos + 1, UDPData, """")
    'now we will get to beginning
    lPos = InStr(lPos + 1, UDPData, """")
    'and the end
    lPosEnd = InStr(lPos + 1, UDPData, """")
    
    SearchForConVarValue = Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
Exit_Properly:
    Exit Function
    
SearchForConVarValue_Error:
    SearchForConVarValue = ""
    GoTo Exit_Properly
    
End Function

Private Function GetMapName(ByVal UDPData As String) As String

    'map name is after first speech mark
    Dim lPos As Long, lPosEnd As Long
    
    lPos = InStr(UDPData, """")
    
    lPosEnd = InStr(lPos + 1, UDPData, """")
    
    GetMapName = Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
End Function

Private Function GetJoinedTeam(ByVal UDPData As String) As String

    '"Spanyardie<2><STEAM_ID_LAN><Unassigned>" joined team "Rebels"
    
    'return the data thus - 'Spanyardie(STEAM_ID)-Rebels'
    
    'name is to first '<'
    'steam id is after second '<'
    'team is after third speech mark
    
    Dim lPos As Long, lPosEnd As Long
    Dim sRet As String
    
    'name
    lPos = InStr(UDPData, "<")
    sRet = Mid$(UDPData, 2, lPos - 2) & "("
    
    'steam id
    lPos = InStr(lPos + 1, UDPData, "<")
    lPosEnd = InStr(lPos + 1, UDPData, ">")
    sRet = sRet & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1)) & ")-"
    
    'team
    lPos = InStr(lPosEnd + 1, UDPData, """")
    lPos = InStr(lPos + 1, UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = sRet & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    'that's it!
    GetJoinedTeam = sRet
    
End Function

Private Function GetConnected(ByVal UDPData As String) As String

    '"Spanyardie<2><STEAM_ID_LAN><>" connected, address "169.254.155.163:27006"
    
    'return the data thus - 'Spanyardie(STEAM_ID)-169.254.155.163:27006'
    
    'name is to first '<'
    'steam id is after second '<'
    'ip address is after third speech mark
    
    Dim lPos As Long, lPosEnd As Long
    Dim sRet As String
    
    'name
    lPos = InStr(UDPData, "<")
    sRet = Mid$(UDPData, 2, lPos - 2) & "("
    
    'steam id
    lPos = InStr(lPos + 1, UDPData, "<")
    lPosEnd = InStr(lPos + 1, UDPData, ">")
    sRet = sRet & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1)) & ")-"
    
    'team
    lPos = InStr(lPosEnd + 1, UDPData, """")
    lPos = InStr(lPos + 1, UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = sRet & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    'that's it!
    GetConnected = sRet

End Function

Private Function GetServerCvar(ByVal UDPData As String) As String

    'server_cvar: "game_mode" "coop"
    
    'return data thus - 'game_mode-coop' - cvar/value
    
    'cvar is up to 2nd speech mark
    'value is after third speech mark
    
    Dim lPos As Long, lPosEnd As Long
    Dim sRet As String
    
    lPos = InStr(UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1)) & "-"
    
    lPos = InStr(lPosEnd + 1, UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = sRet & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    GetServerCvar = sRet
    
End Function

Private Function GetStats(ByVal UDPData As String) As String

    'stats playername "Spanyardie"
    
    'return it thus 'stats playername-Spanyardie
    
    Dim lPos As Long, lPosEnd As Long
    Dim sRet As String
    
    lPos = InStr(UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    GetStats = sRet
    
End Function

Private Function StripPreamble(ByVal UDPData As String) As String

    'remove the preamble from the data
    StripPreamble = Mid$(UDPData, 31)

End Function

Private Function GetGeneric(ByVal UDPData As String) As String

    'generic "mymodcommand" "mycommandvalue"
    
    'return it thus 'generic mymodcommand-mycommandvalue'
    
    Dim lPos As Long, lPosEnd As Long
    Dim sRet As String
    
    lPos = InStr(UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    lPos = InStr(lPosEnd + 1, UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = sRet & "-" & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    GetGeneric = "generic " & sRet
    
End Function

Private Function GetStatsGeneric(ByVal UDPData As String) As String

    'stats generic "mymodstat" "mystatvalue"
    
    'return it thus 'generic mymodcommand-mycommandvalue'
    
    Dim lPos As Long, lPosEnd As Long
    Dim sRet As String
    
    lPos = InStr(UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    lPos = InStr(lPosEnd + 1, UDPData, """")
    lPosEnd = InStr(lPos + 1, UDPData, """")
    sRet = sRet & "-" & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1))
    
    GetStatsGeneric = "stats generic " & sRet
    
    End Function

Private Function GetEntered(ByVal UDPData As String) As String

    '"Spanyardie<2><STEAM_ID_LAN><>" entered the game
    
    'return the data thus - 'Spanyardie(STEAM_ID)-playing'
    
    'name is to first '<'
    'steam id is after second '<'
    
    Dim lPos As Long, lPosEnd As Long
    Dim sRet As String
    
    'name
    lPos = InStr(UDPData, "<")
    sRet = Mid$(UDPData, 2, lPos - 2) & "("
    
    'steam id
    lPos = InStr(lPos + 1, UDPData, "<")
    lPosEnd = InStr(lPos + 1, UDPData, ">")
    sRet = sRet & Mid$(UDPData, lPos + 1, lPosEnd - (lPos + 1)) & ")-playing"
    
    
    'that's it!
    GetEntered = sRet

End Function
