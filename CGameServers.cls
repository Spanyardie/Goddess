VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGameServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CGameServer"
Attribute VB_Ext_KEY = "Member0" ,"CGameServer"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Game Servers Collection Class
'------------------------------------------------------------------------------------
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(RemoteHostName As String, RemoteHostIP As String, RemotePort As Long, Optional sKey As String) As CGameServer

    'create a new object
    Dim oNew As CGameServer
    
    On Error GoTo Add_Error
    
    Set oNew = New CGameServer


    With oNew
        .RemoteHostName = RemoteHostName
        .RemoteHostIP = RemoteHostIP
        .RemotePort = RemotePort
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

Public Property Get Item(vntIndexKey As Variant) As CGameServer
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
    
End Sub

Private Sub Class_Terminate()
    
    Set mCol = Nothing
    
End Sub

Public Function IsGameServerConnected(ByVal RemoteIP As String, ByVal RemotePort As Long) As Boolean

    Dim oGS As CGameServer
    
    On Error GoTo IsGameServerConnected_Error
    
    IsGameServerConnected = False
    
    Set oGS = mCol(RemoteIP & RemotePort)
    
    If Not oGS Is Nothing Then
        IsGameServerConnected = True
    End If
    
Exit_Properly:
    Set oGS = Nothing
    Exit Function
    
IsGameServerConnected_Error:
    IsGameServerConnected = False
    GoTo Exit_Properly
    
End Function

Public Sub Clear()

    Dim lIndex As Long
    
    For lIndex = 1 To mCol.Count
        mCol.Remove 1
    Next lIndex
    
End Sub
