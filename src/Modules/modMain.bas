Attribute VB_Name = "modMain"
Option Explicit

'==============================================================================
' Module:      modMain
' Назначение:  Точка входа макроса. Здесь код связывается с интерфейсом
'              пользователя (кнопка на листе, пункт меню Application.Run и т.п.)
'==============================================================================

' Запуск сценария с путями по умолчанию из modConfig.
Public Sub RunDefault()
    RunScenario modConfig.DEFAULT_INPUT_FILE, modConfig.DEFAULT_OUTPUT_FILE
End Sub

' Запуск сценария с явно заданными путями к входному и выходному файлам.
' Можно вызывать из кнопки на листе или из Application.Run из другой книги.
Public Sub RunScenario(ByVal InputFilePath As String, ByVal OutputFilePath As String)
    Dim app As New clsAppController
    app.Init
    app.Run InputFilePath, OutputFilePath
End Sub

' Быстрая проверка, что подключение к SAP GUI Scripting в принципе работает,
' без полного сценария - удобно для первичной диагностики окружения.
Public Sub TestSapConnection()
    Dim logger As New clsLogger
    logger.Init MinLevel:=llDebug

    Dim connector As New clsSapConnector
    connector.Init logger

    If connector.Connect() Then
        MsgBox "Подключение к SAP GUI успешно установлено.", vbInformation
    Else
        MsgBox "Не удалось подключиться к SAP GUI. Подробности - в лог-файле: " & _
               logger.LogFilePath, vbExclamation
    End If

    logger.CloseLog
End Sub
