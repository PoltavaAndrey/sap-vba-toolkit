VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class:       clsExcelWriter
' Назначение:  Создание новых книг Excel и запись в них данных (например,
'              итогового отчета по результатам обработки в SAP).
'              Работает в отдельном (по умолчанию скрытом) экземпляре Excel.
'
' Пример использования:
'   Dim writer As New clsExcelWriter
'   writer.Init logger
'   writer.CreateNewWorkbook
'   writer.WriteHeaders "Отчет", Array("Дата", "Материал", "Статус")
'   writer.WriteRow "Отчет", 2, Array(Date, "1000123", "OK")
'   writer.SaveAs "C:\output\report.xlsx"
'   writer.CloseFile
'==============================================================================

Private Const MODULE_NAME As String = "clsExcelWriter"

Private mLogger As ILogger
Private mExcelApp As Object
Private mWorkbook As Object
Private mVisible As Boolean

Public Sub Init(ByVal Logger As ILogger, Optional ByVal Visible As Boolean = False)
    Set mLogger = Logger
    mVisible = Visible
End Sub

' -----------------------------------------------------------------------------
' Создает новую пустую книгу в отдельном экземпляре Excel.
' -----------------------------------------------------------------------------
Public Function CreateNewWorkbook() As Boolean
    On Error GoTo ErrHandler

    Set mExcelApp = CreateObject("Excel.Application")
    mExcelApp.Visible = mVisible
    mExcelApp.DisplayAlerts = False

    Set mWorkbook = mExcelApp.Workbooks.Add

    LogInfo "CreateNewWorkbook: создана новая книга"
    CreateNewWorkbook = True
    Exit Function

ErrHandler:
    LogErr "CreateNewWorkbook", Err.Number & " - " & Err.Description
    CreateNewWorkbook = False
End Function

' -----------------------------------------------------------------------------
' Возвращает лист по имени, создавая его при отсутствии.
' -----------------------------------------------------------------------------
Public Function GetOrAddSheet(ByVal SheetName As String) As Object
    On Error Resume Next
    Dim ws As Object
    Set ws = mWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = mWorkbook.Worksheets.Add
        ws.Name = SheetName
    End If

    Set GetOrAddSheet = ws
End Function

' -----------------------------------------------------------------------------
' Записывает заголовки в первую строку листа (создает лист при необходимости).
' -----------------------------------------------------------------------------
Public Sub WriteHeaders(ByVal SheetName As String, ByVal Headers As Variant)
    On Error GoTo ErrHandler

    Dim ws As Object
    Set ws = GetOrAddSheet(SheetName)

    Dim i As Long
    For i = LBound(Headers) To UBound(Headers)
        ws.Cells(1, i - LBound(Headers) + 1).Value = Headers(i)
    Next i

    ws.Rows(1).Font.Bold = True
    LogInfo "WriteHeaders: " & SheetName
    Exit Sub

ErrHandler:
    LogErr "WriteHeaders", Err.Number & " - " & Err.Description
End Sub

' -----------------------------------------------------------------------------
' Записывает одну строку данных (массив значений) в указанную строку листа.
' -----------------------------------------------------------------------------
Public Sub WriteRow(ByVal SheetName As String, ByVal RowIndex As Long, ByVal Values As Variant)
    On Error GoTo ErrHandler

    Dim ws As Object
    Set ws = GetOrAddSheet(SheetName)

    Dim i As Long
    For i = LBound(Values) To UBound(Values)
        ws.Cells(RowIndex, i - LBound(Values) + 1).Value = Values(i)
    Next i
    Exit Sub

ErrHandler:
    LogErr "WriteRow", Err.Number & " - " & Err.Description
End Sub

' -----------------------------------------------------------------------------
' Записывает двумерный массив целиком, начиная с левой верхней ячейки.
' Удобно для выгрузки данных, прочитанных clsExcelReader.ReadRange/ReadSheet.
' -----------------------------------------------------------------------------
Public Sub WriteArray(ByVal SheetName As String, ByVal TopLeftCell As String, ByVal Data As Variant)
    On Error GoTo ErrHandler

    Dim ws As Object
    Set ws = GetOrAddSheet(SheetName)

    Dim rows As Long, cols As Long
    rows = UBound(Data, 1) - LBound(Data, 1) + 1
    cols = UBound(Data, 2) - LBound(Data, 2) + 1

    Dim topLeft As Object
    Set topLeft = ws.Range(TopLeftCell)

    ws.Range(topLeft, topLeft.Offset(rows - 1, cols - 1)).Value = Data

    LogInfo "WriteArray: " & SheetName & " (" & rows & "x" & cols & ")"
    Exit Sub

ErrHandler:
    LogErr "WriteArray", Err.Number & " - " & Err.Description
End Sub

' -----------------------------------------------------------------------------
' Автоподбор ширины столбцов на листе.
' -----------------------------------------------------------------------------
Public Sub AutoFitColumns(ByVal SheetName As String)
    On Error Resume Next
    mWorkbook.Worksheets(SheetName).Columns.AutoFit
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' Сохраняет книгу по указанному пути. Формат определяется по расширению
' (.xlsx = 51, .xlsm = 52, .xls = 56).
' -----------------------------------------------------------------------------
Public Function SaveAs(ByVal FilePath As String) As Boolean
    On Error GoTo ErrHandler

    Dim fmt As Long
    fmt = FileFormatByExtension(FilePath)

    mWorkbook.SaveAs FileName:=FilePath, FileFormat:=fmt
    LogInfo "SaveAs: " & FilePath
    SaveAs = True
    Exit Function

ErrHandler:
    LogErr "SaveAs", Err.Number & " - " & Err.Description
    SaveAs = False
End Function

Public Sub CloseFile(Optional ByVal SaveChanges As Boolean = False)
    On Error Resume Next

    If Not mWorkbook Is Nothing Then
        mWorkbook.Close SaveChanges:=SaveChanges
        Set mWorkbook = Nothing
    End If

    If Not mExcelApp Is Nothing Then
        mExcelApp.Quit
        Set mExcelApp = Nothing
    End If

    LogInfo "CloseFile"
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Me.CloseFile
End Sub

' --- Внутренние помощники ---------------------------------------------------

Private Function FileFormatByExtension(ByVal FilePath As String) As Long
    Dim ext As String
    ext = LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, ".")))

    Select Case ext
        Case "xlsx": FileFormatByExtension = 51
        Case "xlsm": FileFormatByExtension = 52
        Case "xls":  FileFormatByExtension = 56
        Case "csv":  FileFormatByExtension = 6
        Case Else:   FileFormatByExtension = 51
    End Select
End Function

Private Sub LogInfo(ByVal Message As String)
    If Not mLogger Is Nothing Then mLogger.Log llInfo, MODULE_NAME & ": " & Message
End Sub

Private Sub LogErr(ByVal MethodName As String, ByVal Message As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llError, MODULE_NAME & "." & MethodName & ": " & Message
    End If
End Sub
