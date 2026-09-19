VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSapConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class:       clsSapConnector
' Purpose:     Connects to an already running SAP GUI instance via SAP GUI
'              Scripting. This class is only responsible for establishing the
'              connection and obtaining a GuiSession; working with individual
'              controls is delegated to clsSapSession.
'
' Prerequisites:
'   1. Scripting must be enabled in SAP Logon:
'      SAP Logon -> Options -> Accessibility & Scripting -> "Enable scripting"
'      (plus the server-side profile parameter
'      sapgui/user_scripting = TRUE, set via transaction RZ11)
'   2. The user must already be logged on in one of the open SAP GUI windows.
'
' Example:
'   Dim connector As clsSapConnector
'   Set connector = New clsSapConnector
'   connector.Init logger
'
'   If connector.Connect() Then
'       Dim sapSession As clsSapSession
'       Set sapSession = connector.GetSession(0) ' first open session
'   End If
'==============================================================================

Private Const MODULE_NAME As String = "clsSapConnector"

Private mLogger As ILogger
Private mSapGuiAuto As Object   ' SAPGUI (GuiApplication) COM object
Private mConnection As Object   ' GuiConnection
Private mIsConnected As Boolean

' -----------------------------------------------------------------------------
' Injects dependencies (a simple form of DI, since VBA has no parameterized
' constructors).
' -----------------------------------------------------------------------------
Public Sub Init(ByVal Logger As ILogger)
    Set mLogger = Logger
End Sub

' -----------------------------------------------------------------------------
' Connects to the first available running SAP GUI window.
' Returns True on success.
' -----------------------------------------------------------------------------
Public Function Connect() As Boolean
    On Error GoTo ErrHandler

    mLogger.Log llInfo, MODULE_NAME & ".Connect: attempting to connect to SAP GUI"

    If Not IsSapGuiRunning() Then
        mLogger.Log llError, MODULE_NAME & _
            ".Connect: SAP Logon is not running or scripting is unavailable"
        Connect = False
        Exit Function
    End If

    Dim sapGuiAuto As Object
    Set sapGuiAuto = GetObject("SAPGUI")
    Set mSapGuiAuto = sapGuiAuto.GetScriptingEngine

    If mSapGuiAuto.Children.Count = 0 Then
        mLogger.Log llError, MODULE_NAME & _
            ".Connect: no open SAP connections (no window is logged on)"
        Connect = False
        Exit Function
    End If

    Set mConnection = mSapGuiAuto.Children(0)
    mIsConnected = True

    mLogger.Log llInfo, MODULE_NAME & ".Connect: connected successfully. " & _
        "Connection: " & mConnection.Description
    Connect = True
    Exit Function

ErrHandler:
    mLogger.Log llError, MODULE_NAME & ".Connect: error " & Err.Number & _
        " - " & Err.Description
    Connect = False
End Function

' -----------------------------------------------------------------------------
' Returns a clsSapSession wrapper around a GuiSession by index.
' Defaults to index 0 - the first (primary) session of the current connection.
' -----------------------------------------------------------------------------
Public Function GetSession(Optional ByVal SessionIndex As Long = 0) As clsSapSession
    On Error GoTo ErrHandler

    If Not mIsConnected Then
        mLogger.Log llError, MODULE_NAME & _
            ".GetSession: no active connection. Call Connect() first"
        Set GetSession = Nothing
        Exit Function
    End If

    Dim rawSession As Object
    Set rawSession = mConnection.Children(SessionIndex)

    Dim wrapper As New clsSapSession
    wrapper.Init rawSession, mLogger

    Set GetSession = wrapper
    mLogger.Log llInfo, MODULE_NAME & ".GetSession: obtained session #" & SessionIndex
    Exit Function

ErrHandler:
    mLogger.Log llError, MODULE_NAME & ".GetSession: error " & Err.Number & _
        " - " & Err.Description
    Set GetSession = Nothing
End Function

Public Property Get IsConnected() As Boolean
    IsConnected = mIsConnected
End Property

' -----------------------------------------------------------------------------
' Checks whether the SAP Logon process (saplogon.exe) is running.
' This does not guarantee that scripting is enabled, but it rules out the
' obvious case where SAP is not open at all, giving a clearer error message.
' -----------------------------------------------------------------------------
Private Function IsSapGuiRunning() As Boolean
    On Error Resume Next
    Dim testObj As Object
    Set testObj = GetObject("SAPGUI")
    IsSapGuiRunning = (Not testObj Is Nothing)
    On Error GoTo 0
End Function
