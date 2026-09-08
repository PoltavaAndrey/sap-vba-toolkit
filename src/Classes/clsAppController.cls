VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class:       clsAppController
' Назначение:  Фасад приложения. Собирает вместе логгер, коннектор SAP и
'              классы работы с Excel, и предоставляет один понятный метод
'              Run() для запуска сценария. Это единственное место, где
'              создаются конкретные реализации (New clsLogger, New
'              clsSapConnector и т.д.) - остальной код общается только
'              с интерфейсами/обертками.
'
' Пример использования (см. также modMain.bas):
'   Dim app As New clsAppController
'   app.Init
'   app.Run "C:\data\input.xlsx", "C:\data\output.xlsx"
'==============================================================================

Private Const MODULE_NAME As String = "clsAppController"

Private mLogger As ILogger
Private mConnector As clsSapConnector

Public Sub Init()
    Dim logger As New clsLogger
    logger.Init MinLevel:=llDebug, EchoImmediate:=True
    Set mLogger = logger

    Set mConnector = New clsSapConnector
    mConnector.Init mLogger

    mLogger.Log llInfo, MODULE_NAME & ".Init: приложение инициализировано"
End Sub

' -----------------------------------------------------------------------------
' Пример сквозного сценария:
'   1. Читаем входной Excel-файл со списком данных для обработки в SAP.
'   2. Подключаемся к SAP GUI и для каждой строки выполняем транзакцию.
'   3. Пишем результат (успех/ошибка) в новый Excel-отчет.
'
' Метод намеренно оставлен как шаблон - замените TCode, ID полей SAP и
' структуру входного файла на реальные из вашего процесса.
' -----------------------------------------------------------------------------
Public Sub Run(ByVal InputFilePath As String, ByVal OutputFilePath As String)
    On Error GoTo ErrHandler

    mLogger.Log llInfo, "=== Запуск сценария ==="

    ' 1. Чтение входных данных ------------------------------------------------
    Dim reader As New clsExcelReader
    reader.Init mLogger

    If Not reader.OpenFile(InputFilePath) Then
        mLogger.Log llError, "Не удалось открыть входной файл. Остановка."
        Exit Sub
    End If

    Dim records As Collection
    Set records = reader.ReadSheetAsRecords("Лист1")
    reader.CloseFile

    mLogger.Log llInfo, "Прочитано строк: " & records.Count

    ' 2. Подключение к SAP и обработка строк ----------------------------------
    If Not mConnector.Connect() Then
        mLogger.Log llError, "Не удалось подключиться к SAP GUI. Остановка."
        Exit Sub
    End If

    Dim sapSession As clsSapSession
    Set sapSession = mConnector.GetSession(0)

    Dim writer As New clsExcelWriter
    writer.Init mLogger
    writer.CreateNewWorkbook
    writer.WriteHeaders "Результат", Array("Материал", "Статус", "Сообщение SAP")

    Dim i As Long
    Dim rec As Object
    i = 1
    For Each rec In records
        i = i + 1
        ProcessSingleRecord sapSession, rec, writer, i
    Next rec

    ' 3. Сохранение отчета -----------------------------------------------------
    writer.AutoFitColumns "Результат"
    writer.SaveAs OutputFilePath
    writer.CloseFile

    mLogger.Log llInfo, "=== Сценарий завершен. Отчет: " & OutputFilePath & " ==="
    Exit Sub

ErrHandler:
    mLogger.Log llError, MODULE_NAME & ".Run: непредвиденная ошибка " & _
        Err.Number & " - " & Err.Description
End Sub

' -----------------------------------------------------------------------------
' Обработка одной строки входных данных в SAP. Пример для транзакции ME21N
' (создание заказа на закупку) - замените под свою транзакцию и поля.
' -----------------------------------------------------------------------------
Private Sub ProcessSingleRecord(ByVal sapSession As clsSapSession, _
                                 ByVal rec As Object, _
                                 ByVal writer As clsExcelWriter, _
                                 ByVal OutputRow As Long)

    Dim material As String
    material = CStr(rec("Материал"))

    mLogger.Log llInfo, "Обработка материала: " & material

    sapSession.StartTransaction "ME21N"
    sapSession.SetText "wnd[0]/usr/ctxtEKKO-BSART", "NB"
    sapSession.SetText "wnd[0]/usr/tabsHEADER_DETAIL/.../ctxtEKPO-MATNR", material
    sapSession.SendVKey 0 ' Enter

    Dim statusText As String
    statusText = sapSession.GetStatusBarText()

    Dim isOk As Boolean
    isOk = (InStr(1, LCase(statusText), "error") = 0) And _
           (InStr(1, statusText, "ошибк") = 0)

    writer.WriteRow "Результат", OutputRow, _
        Array(material, IIf(isOk, "OK", "ОШИБКА"), statusText)

    sapSession.EndTransaction
End Sub
