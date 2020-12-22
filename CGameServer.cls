VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGameServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"CUDPPackets"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Game Server Class
'------------------------------------------------------------------------------------
Option Explicit

Private mvarPackets As CUDPPackets
Private mvarRemoteHostName As String
Private mvarRemoteHostIP As String
Private mvarRemotePort As Long
Private mbIgnorePackets As Boolean

Public Sub SaveToFile(ByVal FileNameAndPath As String)
End Sub

Public Property Let RemotePort(ByVal vData As Long)
    
    mvarRemotePort = vData
    
End Property

Public Property Get RemotePort() As Long
    
    RemotePort = mvarRemotePort
    
End Property

Public Property Let RemoteHostIP(ByVal vData As String)
    
    mvarRemoteHostIP = vData
    
End Property


Public Property Get RemoteHostIP() As String
    
    RemoteHostIP = mvarRemoteHostIP
    
End Property

Public Property Let RemoteHostName(ByVal vData As String)
    
    mvarRemoteHostName = vData
    
End Property

Public Property Get RemoteHostName() As String
    
    RemoteHostName = mvarRemoteHostName
    
End Property

Public Property Get Packets() As CUDPPackets
    
    If mvarPackets Is Nothing Then
        Set mvarPackets = New CUDPPackets
    End If

    Set Packets = mvarPackets
    
End Property

Private Sub Class_Initialize()

    Set mvarPackets = New CUDPPackets
    
    mbIgnorePackets = False
    
End Sub

Private Sub Class_Terminate()

    Set mvarPackets = Nothing
    
End Sub

Public Property Get IgnorePackets() As Boolean

    IgnorePackets = mbIgnorePackets
    
End Property

Public Property Let IgnorePackets(ByVal bNew As Boolean)

    mbIgnorePackets = bNew
    
End Property
