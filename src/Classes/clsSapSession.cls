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
' Назначение:  Обертка над объектом GuiSession, предоставляющая удобные
'              высокоуровневые методы для работы с элементами управления
'              SAP GUI (поля ввода, кнопки, чекбоксы, таблицы, вкладки и т.д.)
'              без необходимости каждый раз писать FindById и приводить типы.
'
' Замечание:   Все Id элементов копируются из SAP через
'              "Помощь по сценарию" (F6 / кнопка со значком в статусной строке)
'              либо через встроенный SAP Scripting Recorder.
'
' Пример использования:
'   sapSession.StartTransaction "ME21N"
'   sapSession.SetText "wnd[0]/usr/ctxtEKKO-BSART", "NB"
'   sapSession.ClickButton "wnd[0]/tbar[0]/btn[11]" ' Сохранить
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
' Запускает транзакцию по коду (аналог ввода кода в поле команд).
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
' Установить текстовое значение в поле ввода (GuiTextField, GuiCTextField и т.п.)
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
' Прочитать текстовое значение поля.
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
' Нажать кнопку (GuiButton) или "нажимаемый" элемент со свойством Press.
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
' Выбрать вкладку (GuiTab).
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
' Установить состояние чекбокса (GuiCheckBox).
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
' Выбрать значение в выпадающем списке (GuiComboBox) по ключу.
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
' Отправить функциональную клавишу (0 = Enter, 3 = F3/Назад, 11 = Сохранить,
' 15 = F12/Отмена и т.д. - см. таблицу VKey в документации SAP).
' -----------------------------------------------------------------------------
Public Sub SendVKey(ByVal KeyCode As Integer)
    LogAction "SendVKey(" & KeyCode & ")"
    mSession.FindById("wnd[0]").SendVKey KeyCode
End Sub

' -----------------------------------------------------------------------------
' Прочитать текст строки статуса (сообщение об ошибке/успехе после действия).
' -----------------------------------------------------------------------------
Public Function GetStatusBarText() As String
    On Error GoTo ErrHandler
    GetStatusBarText = mSession.FindById("wnd[0]/sbar").Text
    Exit Function

ErrHandler:
    GetStatusBarText = vbNullString
End Function

' -----------------------------------------------------------------------------
' Проверка наличия элемента без выброса исключения (полезно перед обращением
' к опциональным полям/всплывающим окнам).
' -----------------------------------------------------------------------------
Public Function ControlExists(ByVal ControlId As String) As Boolean
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = mSession.FindById(ControlId)
    ControlExists = (Not ctrl Is Nothing)
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' Прямой доступ к "сырому" GuiSession для нестандартных случаев,
' не покрытых обертками выше.
' -----------------------------------------------------------------------------
Public Property Get RawSession() As Object
    Set RawSession = mSession
End Property

' --- Внутренние помощники ---------------------------------------------------

Private Sub LogAction(ByVal Description As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llDebug, MODULE_NAME & ": " & Description
    End If
End Sub

Private Sub LogError(ByVal MethodName As String, ByVal ControlId As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llError, MODULE_NAME & "." & MethodName & _
            ": не найден или недоступен элемент '" & ControlId & _
            "' (" & Err.Number & " - " & Err.Description & ")"
    End If
End Sub
