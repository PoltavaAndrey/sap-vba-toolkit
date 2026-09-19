VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Interface:   ILogger
' Purpose:     Abstraction for a logger. Lets the concrete implementation
'              (file, Immediate Window, database, etc.) be swapped without
'              changing the code that performs the logging.
'
' Usage:
'   Dim logger As ILogger
'   Set logger = New clsLogger
'   logger.Log llInfo, "Message"
'==============================================================================

Public Enum LogLevel
    llDebug = 0
    llInfo = 1
    llWarning = 2
    llError = 3
End Enum

' Write a single message to the log at the given severity level.
Public Sub Log(ByVal Level As LogLevel, ByVal Message As String)
End Sub

' Explicitly close the log (release the file handle, etc.).
Public Sub CloseLog()
End Sub
