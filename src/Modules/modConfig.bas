Attribute VB_Name = "modConfig"
Option Explicit

'==============================================================================
' Module:      modConfig
' Назначение:  Централизованные константы и настройки проекта.
'              Меняйте значения здесь, а не по всему коду.
'==============================================================================

Public Const APP_NAME As String = "SAP VBA Toolkit"
Public Const APP_VERSION As String = "0.1.0"

' Пути по умолчанию (используются, если пользователь не указал свои)
Public Const DEFAULT_LOG_SUBFOLDER As String = "Logs"

' Пример путей к файлам - настройте под свою задачу или выносите в UI/аргументы
Public Const DEFAULT_INPUT_FILE As String = "C:\Data\input.xlsx"
Public Const DEFAULT_OUTPUT_FILE As String = "C:\Data\output.xlsx"
