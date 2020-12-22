VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CFilter"
Attribute VB_Ext_KEY = "Member0" ,"CFilter"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Filters collection class
'------------------------------------------------------------------------------------
Option Explicit

Private mCol As Collection

Private Const MODULE As String = "CFilters:"

Public Function Add(FilterType As String, FilterEvent As String, FilterValue As Variant, CommandEvent As Boolean, CommandValue As Boolean, Optional sKey As String) As CFilter

    'create a new object
    Dim oNew As CFilter
    
    On Error GoTo Add_Error
    
    Set oNew = New CFilter

    With oNew
        .FilterType = FilterType
        .FilterEvent = FilterEvent
        If IsObject(FilterValue) Then
            Set .FilterValue = FilterValue
        Else
            .FilterValue = FilterValue
        End If
        .CommandEvent = CommandEvent
        .CommandValue = CommandValue
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

Public Property Get Item(vntIndexKey As Variant) As CFilter
Attribute Item.VB_UserMemId = 0

    Set Item = mCol(vntIndexKey)
  
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

