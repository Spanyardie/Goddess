VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilter"
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
'-          Filter Class
'------------------------------------------------------------------------------------
Option Explicit

Private mvarFilterType As String
Private mvarFilterEvent As String
Private mvarFilterValue As Variant
Private mbCommandEvent As Boolean
Private mbCommandValue As Boolean

Public Property Let FilterValue(ByVal vData As Variant)

    mvarFilterValue = vData
    
End Property

Public Property Set FilterValue(ByVal vData As Variant)

    Set mvarFilterValue = vData
    
End Property

Public Property Get FilterValue() As Variant

    If IsObject(mvarFilterValue) Then
        Set FilterValue = mvarFilterValue
    Else
        FilterValue = mvarFilterValue
    End If
    
End Property

Public Property Let FilterEvent(ByVal vData As String)

    mvarFilterEvent = vData
    
End Property

Public Property Get FilterEvent() As String
    
    FilterEvent = mvarFilterEvent
    
End Property

Public Property Let FilterType(ByVal vData As String)
    
    mvarFilterType = vData
    
End Property

Public Property Get FilterType() As String
    
    FilterType = mvarFilterType
    
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

