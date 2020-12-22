VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupportedProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Supported Protocol Class
'------------------------------------------------------------------------------------
Option Explicit

Private mvarProtocolName As String
Private mvarPacketLength As Long
Private mvarLocalHost As String
Private mvarListenPort As Long

Public Property Let ListenPort(ByVal vData As Long)
    
    mvarListenPort = vData
    
End Property

Public Property Get ListenPort() As Long
    
    ListenPort = mvarListenPort
    
End Property

Public Property Let LocalHost(ByVal vData As String)
    
    mvarLocalHost = vData
    
End Property

Public Property Get LocalHost() As String
    
    LocalHost = mvarLocalHost
    
End Property

Public Property Let PacketLength(ByVal vData As Long)
    
    mvarPacketLength = vData
    
End Property

Public Property Get PacketLength() As Long
    
    PacketLength = mvarPacketLength
    
End Property

Public Property Let ProtocolName(ByVal vData As String)
    
    mvarProtocolName = vData
    
End Property

Public Property Get ProtocolName() As String
    
    ProtocolName = mvarProtocolName
    
End Property

