VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUDPPacketData"
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
'-          UDP Packet Data Class
'------------------------------------------------------------------------------------
Option Explicit

Private mvarRemoteIP As String
Private mvarRemotePort As Long
Private mvarRawData As String
Private mvarDateRcv As Date
Private mvarTimeRcv As Date
Private mvarPacketID As Long
Private mbCommandEvent As Boolean
Private mbCommandValue As Boolean

Public Property Let PacketID(ByVal vData As Long)
    
    mvarPacketID = vData
    
End Property

Public Property Get PacketID() As Long
    
    PacketID = mvarPacketID
    
End Property

Public Property Let TimeRcv(ByVal vData As Date)
    
    mvarTimeRcv = vData
    
End Property

Public Property Get TimeRcv() As Date
    
    TimeRcv = mvarTimeRcv
    
End Property

Public Property Let DateRcv(ByVal vData As Date)
    
    mvarDateRcv = vData
    
End Property

Public Property Get DateRcv() As Date
    
    DateRcv = mvarDateRcv
    
End Property

Public Property Let RawData(ByVal vData As String)
    
    mvarRawData = vData
    
End Property

Public Property Get RawData() As String
    
    RawData = mvarRawData
    
End Property

Public Property Let RemotePort(ByVal vData As Long)
    
    mvarRemotePort = vData
    
End Property

Public Property Get RemotePort() As Long
    
    RemotePort = mvarRemotePort
    
End Property

Public Property Let RemoteIP(ByVal vData As String)
    
    mvarRemoteIP = vData
    
End Property

Public Property Get RemoteIP() As String
    
    RemoteIP = mvarRemoteIP
    
End Property


Public Property Get CommandEvent() As Boolean

    CommandEvent = mbCommandEvent
    
End Property

Public Property Let CommandEvent(ByVal bNew As Boolean)

    mbCommandEvent = bNew
    
End Property

Public Property Get CommandValue() As Boolean

    CommandValue = mbCommandValue
    
End Property

Public Property Let CommandValue(ByVal bNew As Boolean)

    mbCommandValue = bNew
    
End Property
