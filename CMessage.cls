VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMessage"
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
'-          Message Class
'------------------------------------------------------------------------------------
Option Explicit

Private mvarMessageName As String
Private mbCommandEvent As Boolean 'this is a command without a value
Private mbCommandValue As Boolean

Public Property Let MessageName(ByVal vData As String)

    mvarMessageName = vData
    
End Property

Public Property Get MessageName() As String

    MessageName = mvarMessageName
    
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

