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
' Назначение:  Подключение к уже запущенному SAP GUI через SAP GUI Scripting
'              (SAP GUI Scripting API / "wdyauto"). Класс отвечает только
'              за установление соединения и получение GuiSession; работа
'              с элементами управления вынесена в clsSapSession.
'
' Предварительные условия:
'   1. В SAP Logon включен и разрешен скриптинг:
'      SAP Logon -> Параметры -> Скриптинг -> "Включить скриптинг"
'      (плюс серверная опция RZ11 sapgui/user_scripting = TRUE)
'   2. Пользователь уже вошел в систему в одном из открытых окон SAP GUI.
'
' Пример использования:
'   Dim connector As clsSapConnector
'   Set connector = New clsSapConnector
'   connector.Init logger
'
'   If connector.Connect() Then
'       Dim sapSession As clsSapSession
'       Set sapSession = connector.GetSession(0) ' первая открытая сессия
'   End If
'==============================================================================

Private Const MODULE_NAME As String = "clsSapConnector"

Private mLogger As ILogger
Private mSapGuiAuto As Object   ' SAPGUI (GuiApplication) COM-объект
Private mConnection As Object   ' GuiConnection
Private mIsConnected As Boolean

' -----------------------------------------------------------------------------
' Инициализация зависимостей (простая форма DI, т.к. в VBA нет конструкторов
' с параметрами).
' -----------------------------------------------------------------------------
Public Sub Init(ByVal Logger As ILogger)
    Set mLogger = Logger
End Sub

' -----------------------------------------------------------------------------
' Подключается к первому доступному запущенному окну SAP GUI.
' Возвращает True при успехе.
' -----------------------------------------------------------------------------
Public Function Connect() As Boolean
    On Error GoTo ErrHandler

    mLogger.Log llInfo, MODULE_NAME & ".Connect: попытка подключения к SAP GUI"

    If Not IsSapGuiRunning() Then
        mLogger.Log llError, MODULE_NAME & _
            ".Connect: SAP Logon не запущен или скриптинг недоступен"
        Connect = False
        Exit Function
    End If

    Dim sapGuiAuto As Object
    Set sapGuiAuto = GetObject("SAPGUI")
    Set mSapGuiAuto = sapGuiAuto.GetScriptingEngine

    If mSapGuiAuto.Children.Count = 0 Then
        mLogger.Log llError, MODULE_NAME & _
            ".Connect: нет открытых соединений SAP (ни одно окно не залогинено)"
        Connect = False
        Exit Function
    End If

    Set mConnection = mSapGuiAuto.Children(0)
    mIsConnected = True

    mLogger.Log llInfo, MODULE_NAME & ".Connect: успешно подключено. " & _
        "Соединение: " & mConnection.Description
    Connect = True
    Exit Function

ErrHandler:
    mLogger.Log llError, MODULE_NAME & ".Connect: ошибка " & Err.Number & _
        " - " & Err.Description
    Connect = False
End Function

' -----------------------------------------------------------------------------
' Возвращает обертку clsSapSession над GuiSession по индексу.
' По умолчанию - первая (основная) сессия текущего соединения.
' -----------------------------------------------------------------------------
Public Function GetSession(Optional ByVal SessionIndex As Long = 0) As clsSapSession
    On Error GoTo ErrHandler

    If Not mIsConnected Then
        mLogger.Log llError, MODULE_NAME & _
            ".GetSession: нет активного соединения. Сначала вызовите Connect()"
        Set GetSession = Nothing
        Exit Function
    End If

    Dim rawSession As Object
    Set rawSession = mConnection.Children(SessionIndex)

    Dim wrapper As New clsSapSession
    wrapper.Init rawSession, mLogger

    Set GetSession = wrapper
    mLogger.Log llInfo, MODULE_NAME & ".GetSession: получена сессия #" & SessionIndex
    Exit Function

ErrHandler:
    mLogger.Log llError, MODULE_NAME & ".GetSession: ошибка " & Err.Number & _
        " - " & Err.Description
    Set GetSession = Nothing
End Function

Public Property Get IsConnected() As Boolean
    IsConnected = mIsConnected
End Property

' -----------------------------------------------------------------------------
' Проверка, что процесс SAP Logon (saplogon.exe) запущен.
' Это не гарантирует включенный скриптинг, но отсекает очевидный случай,
' когда SAP вообще не открыт - и дает более понятное сообщение об ошибке.
' -----------------------------------------------------------------------------
Private Function IsSapGuiRunning() As Boolean
    On Error Resume Next
    Dim testObj As Object
    Set testObj = GetObject("SAPGUI")
    IsSapGuiRunning = (Not testObj Is Nothing)
    On Error GoTo 0
End Function
