VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ILogger

'==============================================================================
' Class:       clsLogger
' Implements:  ILogger
' Purpose:     Simple file-based logger with daily rotation and optional
'              mirroring of messages to the Immediate Window (for debugging
'              in the VBA editor).
'
' Example:
'   Dim logger As ILogger
'   Set logger = New clsLogger
'   logger.Log llInfo, "Processing started"
'   ...
'   logger.CloseLog
'
' Log line format:
'   2026-01-15 10:23:45 | INFO    | Message text
'==============================================================================

Private Const MODULE_NAME As String = "clsLogger"

Private mLogFilePath As String
Private mFileNumber As Integer
Private mIsOpen As Boolean
Private mMinLevel As LogLevel
Private mEchoToImmediate As Boolean

' -----------------------------------------------------------------------------
' Initializes the logger.
' LogFolder      - folder for log files (created automatically if missing)
' MinLevel       - minimum severity level that is actually written
' EchoImmediate  - whether to also mirror messages to Debug.Print
' -----------------------------------------------------------------------------
Public Sub Init(Optional ByVal LogFolder As String = vbNullString, _
                 Optional ByVal MinLevel As LogLevel = llDebug, _
                 Optional ByVal EchoImmediate As Boolean = True)

    Dim folder As String
    folder = LogFolder
    If Len(folder) = 0 Then folder = ThisWorkbook.Path & "\Logs"

    If Len(Dir(folder, vbDirectory)) = 0 Then
        MkDir folder
    End If

    mLogFilePath = folder & "\" & Format(Now, "yyyy-mm-dd") & "_macro.log"
    mMinLevel = MinLevel
    mEchoToImmediate = EchoImmediate

    mFileNumber = FreeFile
    Open mLogFilePath For Append As #mFileNumber
    mIsOpen = True

    Me.Log llInfo, "=== Logger initialized. File: " & mLogFilePath & " ==="
End Sub

Private Sub Class_Terminate()
    Me.CloseLog
End Sub

' --- ILogger implementation --------------------------------------------------

Private Sub ILogger_Log(ByVal Level As LogLevel, ByVal Message As String)
    Me.Log Level, Message
End Sub

Private Sub ILogger_CloseLog()
    Me.CloseLog
End Sub

' --- Public methods (for direct use of clsLogger) ---------------------------

Public Sub Log(ByVal Level As LogLevel, ByVal Message As String)
    Dim line As String

    If Not mIsOpen Then Me.Init ' auto-initialize with default parameters
    If Level < mMinLevel Then Exit Sub

    line = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
           PadLevel(Level) & " | " & Message

    Print #mFileNumber, line

    If mEchoToImmediate Then
        Debug.Print line
    End If
End Sub

Public Sub CloseLog()
    If mIsOpen Then
        Close #mFileNumber
        mIsOpen = False
    End If
End Sub

' --- Helpers ------------------------------------------------------------------

Private Function PadLevel(ByVal Level As LogLevel) As String
    Select Case Level
        Case llDebug:   PadLevel = "DEBUG  "
        Case llInfo:    PadLevel = "INFO   "
        Case llWarning: PadLevel = "WARNING"
        Case llError:   PadLevel = "ERROR  "
        Case Else:      PadLevel = "?????? "
    End Select
End Function

Public Property Get LogFilePath() As String
    LogFilePath = mLogFilePath
End Property
