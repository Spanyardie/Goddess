VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMessageFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CMessageFile"
Attribute VB_Ext_KEY = "Member0" ,"CMessageFile"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Message Files Collection Class
'------------------------------------------------------------------------------------
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(FileType As String, FilePath As String, Optional sKey As String) As CMessageFile

    Dim oNew As CMessageFile
    
    Set oNew = New CMessageFile

    On Error GoTo Add_Error
    
    With oNew
        .FileType = FileType
        .FilePath = FilePath
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

Public Property Get Item(vntIndexKey As Variant) As CMessageFile
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

