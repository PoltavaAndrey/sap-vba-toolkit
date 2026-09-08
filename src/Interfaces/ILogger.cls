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
' Назначение:  Абстракция логгера. Позволяет подменять реализацию
'              (файл, Immediate Window, БД и т.п.) без изменения кода,
'              который логированием пользуется.
'
' Использование:
'   Dim logger As ILogger
'   Set logger = New clsLogger
'   logger.Log llInfo, "Сообщение"
'==============================================================================

Public Enum LogLevel
    llDebug = 0
    llInfo = 1
    llWarning = 2
    llError = 3
End Enum

' Записать одно сообщение в лог с указанным уровнем важности.
Public Sub Log(ByVal Level As LogLevel, ByVal Message As String)
End Sub

' Явно закрыть лог (освободить файловый дескриптор и т.п.)
Public Sub CloseLog()
End Sub
