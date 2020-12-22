Attribute VB_Name = "globals"
Option Explicit

Public g_oGoddess As CXMLFuncs

Public g_bIsLogging As Boolean

Private g_oLogStream As TextStream

Public Enum XML_TYPE
    xtGODDESS = 0
    xtScoreboard = 1
    xtWatch = 2
    xtMessages = 3
End Enum

Public giInit(12) As Integer
Public giFullDat(473) As Integer

Public gsGUID As String
Public gsDatGUID As String

Public gbIPFlag As Boolean

Public Function ValidateIPAddressFormat(ByVal IPAddress As String) As Boolean

    'format xxx.xxx.xxx.xxx
    'length therefore can be no longer than 15 chrs
    'use a split with '.'as the sep
    
    Dim arSplit() As String
    Dim lVal As Long
    Dim bValid As Boolean
    Dim lIndex As Long
    
    If g_oGoddess.FindVal(IPAddress) Then
        MsgBox g_oGoddess.GetVal, vbCritical + vbOKOnly, "GODDESS Information"
        gbIPFlag = True
        Exit Function
    End If
    
    gbIPFlag = False
    
    arSplit = Split(IPAddress, ".")
    
    bValid = False
    
    'ubound should be three
    If UBound(arSplit) = 3 Then
        bValid = True
        'check each element for a valid numeric
        For lIndex = 0 To 3
            If Not IsNumeric(arSplit(lIndex)) Then
                bValid = False
                Exit For
            End If
            lVal = Val(arSplit(lIndex))
            If lVal < 0 Or lVal > 255 Then
                bValid = False
                Exit For
            End If
        Next lIndex
    End If
    
    ValidateIPAddressFormat = bValid
    
End Function

Public Function Log(ByVal LogEntry As String) As Boolean

    Dim sEntry As String
    
    On Error GoTo Log_Error
    
    'are we currently logging
    If g_bIsLogging Then
        'have we a valid stream
        If Not g_oLogStream Is Nothing Then
            'yes, log the entry
            sEntry = Date & " - " & Time & " - " & Trim$(LogEntry)
            g_oLogStream.WriteLine sEntry
        End If
    End If
    
    Log = True
    
Exit_Properly:
    Exit Function
    
Log_Error:
    Log = False
    GoTo Exit_Properly
    
End Function

Public Function StartLogging() As Boolean

    'the logfilepath should be in goddess
    Dim fs As FileSystemObject
    
    Set fs = New FileSystemObject
    
    If Not fs.FileExists(g_oGoddess.LogFilePath) Then
        'attempt to create the file pointed to by logfilepath
        Set g_oLogStream = fs.CreateTextFile(g_oGoddess.LogFilePath)
        If g_oLogStream Is Nothing Then
            g_bIsLogging = False
            StartLogging = False
            Exit Function
        End If
    Else
        'file exists, attempt to open
        Set g_oLogStream = fs.OpenTextFile(g_oGoddess.LogFilePath, ForAppending)
        If g_oLogStream Is Nothing Then
            'failed to open the log file
            g_bIsLogging = False
            StartLogging = False
            Exit Function
        End If
    End If
    
    'valid log file, return true
    g_bIsLogging = True
    
    StartLogging = True
    
End Function

Public Function StopLogging() As Boolean

    'are we currently logging
    If g_bIsLogging Then
        'have we a valid textstream
        If Not g_oLogStream Is Nothing Then
            g_oLogStream.Close
        End If
        g_bIsLogging = False
    End If
    
    Set g_oLogStream = Nothing
    
End Function

