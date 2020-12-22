VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUDPPackets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CUDPPacketData"
Attribute VB_Ext_KEY = "Member0" ,"CUDPPacketData"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Packet collection class
'------------------------------------------------------------------------------------
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mlPacketID As Long

Public Function Add(RemoteIP As String, RemotePort As Long, RawData As String, DateRcv As Date, TimeRcv As Date, PacketID As Long, ByVal CommandEvent As Boolean, ByVal CommandValue As Boolean, Optional sKey As String) As CUDPPacketData
    
    Dim oNew As CUDPPacketData
    Dim lPacketID As Long
    
    On Error GoTo Add_Error
    
    'create a new packet
    Set oNew = New CUDPPacketData

    lPacketID = UDPPacketID
    
    'set the data
    With oNew
        .CommandEvent = CommandEvent
        .CommandValue = CommandValue
        .RemoteIP = RemoteIP
        .RemotePort = RemotePort
        .RawData = RawData
        .DateRcv = DateRcv
        .TimeRcv = TimeRcv
        .PacketID = lPacketID
        'create the new key
        sKey = .RemoteIP & .RemotePort & lPacketID
        If Len(sKey) = 0 Then
            mCol.Add oNew
        Else
            mCol.Add oNew, sKey
        End If
    End With

    'return the object created
    Set Add = oNew
    
Exit_Properly:
    Set oNew = Nothing
    Exit Function
    
Add_Error:
    Set Add = Nothing
    GoTo Exit_Properly
    
End Function

Public Property Get Item(vntIndexKey As Variant) As CUDPPacketData
Attribute Item.VB_UserMemId = 0
  
  On Error GoTo Item_Get_Error
  
  Set Item = mCol(vntIndexKey)
  
Exit_Properly:
    Exit Property
    
Item_Get_Error:
    Set Item = Nothing
    GoTo Exit_Properly
    
End Property

Public Property Get Count() As Long
    
    Count = mCol.Count
    
End Property


Public Sub Remove(vntIndexKey As Variant)
    
    mCol.Remove vntIndexKey
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    
    Set NewEnum = mCol.[_NewEnum]
    
End Property

Private Sub Class_Initialize()
    
    Set mCol = New Collection
    
    mlPacketID = 1
    
End Sub

Private Sub Class_Terminate()
    
    Set mCol = Nothing
    
End Sub

Public Property Get UDPPacketID() As Long

    UDPPacketID = mlPacketID
    
    mlPacketID = mlPacketID + 1
    
End Property

Public Sub Clear()

    Dim lIndex As Long
    
    For lIndex = 1 To mCol.Count
        mCol.Remove lIndex
    Next lIndex
    
End Sub
