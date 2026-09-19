VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSapSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class:       clsSapSession
' Purpose:     Wraps a GuiSession object and exposes convenient high-level
'              methods for working with SAP GUI controls (input fields,
'              buttons, checkboxes, tables, tabs, etc.) without having to
'              write FindById and cast types every time.
'
' Note:        All control IDs are copied from SAP using "Script Recording
'              and Playback" (F6 / the icon in the status bar) or by hovering
'              a field with the built-in SAP scripting helper.
'
' Example:
'   sapSession.StartTransaction "ME21N"
'   sapSession.SetText "wnd[0]/usr/ctxtEKKO-BSART", "NB"
'   sapSession.ClickButton "wnd[0]/tbar[0]/btn[11]" ' Save
'==============================================================================

Private Const MODULE_NAME As String = "clsSapSession"

Private mSession As Object      ' GuiSession
Private mLogger As ILogger

' -----------------------------------------------------------------------------
Public Sub Init(ByVal RawSession As Object, ByVal Logger As ILogger)
    Set mSession = RawSession
    Set mLogger = Logger
End Sub

' -----------------------------------------------------------------------------
' Starts a transaction by code (equivalent to typing the code into the
' command field).
' -----------------------------------------------------------------------------
Public Sub StartTransaction(ByVal TCode As String)
    LogAction "StartTransaction(" & TCode & ")"
    mSession.StartTransaction TCode
End Sub

Public Sub EndTransaction()
    LogAction "EndTransaction"
    mSession.EndTransaction
End Sub

' -----------------------------------------------------------------------------
' Sets the text value of an input field (GuiTextField, GuiCTextField, etc.)
' -----------------------------------------------------------------------------
Public Function SetText(ByVal ControlId As String, ByVal Value As String) As Boolean
    On Error GoTo ErrHandler
    LogAction "SetText(" & ControlId & ", " & Value & ")"

    mSession.FindById(ControlId).Text = Value
    SetText = True
    Exit Function

ErrHandler:
    LogError "SetText", ControlId
    SetText = False
End Function

' -----------------------------------------------------------------------------
' Reads the text value of a field.
' -----------------------------------------------------------------------------
Public Function GetText(ByVal ControlId As String) As String
    On Error GoTo ErrHandler
    GetText = mSession.FindById(ControlId).Text
    Exit Function

ErrHandler:
    LogError "GetText", ControlId
    GetText = vbNullString
End Function

' -----------------------------------------------------------------------------
' Presses a button (GuiButton) or any "pressable" control with a Press method.
' -----------------------------------------------------------------------------
Public Function ClickButton(ByVal ControlId As String) As Boolean
    On Error GoTo ErrHandler
    LogAction "ClickButton(" & ControlId & ")"

    mSession.FindById(ControlId).Press
    ClickButton = True
    Exit Function

ErrHandler:
    LogError "ClickButton", ControlId
    ClickButton = False
End Function

' -----------------------------------------------------------------------------
' Selects a tab (GuiTab).
' -----------------------------------------------------------------------------
Public Function SelectTab(ByVal ControlId As String) As Boolean
    On Error GoTo ErrHandler
    LogAction "SelectTab(" & ControlId & ")"

    mSession.FindById(ControlId).Select
    SelectTab = True
    Exit Function

ErrHandler:
    LogError "SelectTab", ControlId
    SelectTab = False
End Function

' -----------------------------------------------------------------------------
' Sets the state of a checkbox (GuiCheckBox).
' -----------------------------------------------------------------------------
Public Function SetCheckBox(ByVal ControlId As String, ByVal Checked As Boolean) As Boolean
    On Error GoTo ErrHandler
    LogAction "SetCheckBox(" & ControlId & ", " & Checked & ")"

    mSession.FindById(ControlId).Selected = Checked
    SetCheckBox = True
    Exit Function

ErrHandler:
    LogError "SetCheckBox", ControlId
    SetCheckBox = False
End Function

' -----------------------------------------------------------------------------
' Selects a value in a dropdown (GuiComboBox) by key.
' -----------------------------------------------------------------------------
Public Function SelectComboBoxItem(ByVal ControlId As String, ByVal Key As String) As Boolean
    On Error GoTo ErrHandler
    LogAction "SelectComboBoxItem(" & ControlId & ", " & Key & ")"

    mSession.FindById(ControlId).Key = Key
    SelectComboBoxItem = True
    Exit Function

ErrHandler:
    LogError "SelectComboBoxItem", ControlId
    SelectComboBoxItem = False
End Function

' -----------------------------------------------------------------------------
' Sends a function key (0 = Enter, 3 = F3/Back, 11 = Save, 15 = F12/Cancel,
' etc. - see the SAP VKey code table in the documentation).
' -----------------------------------------------------------------------------
Public Sub SendVKey(ByVal KeyCode As Integer)
    LogAction "SendVKey(" & KeyCode & ")"
    mSession.FindById("wnd[0]").SendVKey KeyCode
End Sub

' -----------------------------------------------------------------------------
' Reads the status bar text (an error/success message shown after an action).
' -----------------------------------------------------------------------------
Public Function GetStatusBarText() As String
    On Error GoTo ErrHandler
    GetStatusBarText = mSession.FindById("wnd[0]/sbar").Text
    Exit Function

ErrHandler:
    GetStatusBarText = vbNullString
End Function

' -----------------------------------------------------------------------------
' Checks whether a control exists without raising an exception (useful before
' touching optional fields/popup windows).
' -----------------------------------------------------------------------------
Public Function ControlExists(ByVal ControlId As String) As Boolean
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = mSession.FindById(ControlId)
    ControlExists = (Not ctrl Is Nothing)
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' Returns a clsSapTable wrapper for a classic dynpro table control
' (GuiTableControl), e.g. the item table in ME21N/ME22N/ME23N/VA01, etc.
' Not applicable to ALV Grid controls (GuiGridView) - those are exposed via
' a different scripting API and are out of scope for this toolkit.
' -----------------------------------------------------------------------------
Public Function GetTable(ByVal ControlId As String) As clsSapTable
    On Error GoTo ErrHandler

    Dim rawTable As Object
    Set rawTable = mSession.FindById(ControlId)

    Dim wrapper As New clsSapTable
    wrapper.Init rawTable, mLogger

    Set GetTable = wrapper
    LogAction "GetTable(" & ControlId & ")"
    Exit Function

ErrHandler:
    LogError "GetTable", ControlId
    Set GetTable = Nothing
End Function

' -----------------------------------------------------------------------------
' Direct access to the "raw" GuiSession for non-standard cases not covered
' by the wrappers above.
' -----------------------------------------------------------------------------
Public Property Get RawSession() As Object
    Set RawSession = mSession
End Property

' --- Internal helpers ---------------------------------------------------------

Private Sub LogAction(ByVal Description As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llDebug, MODULE_NAME & ": " & Description
    End If
End Sub

Private Sub LogError(ByVal MethodName As String, ByVal ControlId As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llError, MODULE_NAME & "." & MethodName & _
            ": control '" & ControlId & "' not found or unavailable" & _
            " (" & Err.Number & " - " & Err.Description & ")"
    End If
End Sub
