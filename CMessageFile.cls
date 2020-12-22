VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMessageFile"
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
'-          Message File Class
'------------------------------------------------------------------------------------
Option Explicit

Private mvarFileType As String
Private mvarFilePath As String

Public Property Let FilePath(ByVal vData As String)
    
    mvarFilePath = vData
    
End Property

Public Property Get FilePath() As String
    
    FilePath = mvarFilePath
    
End Property

Public Property Let FileType(ByVal vData As String)
    
    mvarFileType = vData
    
End Property

Public Property Get FileType() As String
    
    FileType = mvarFileType
    
End Property

