VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class:       clsExcelReader
' Назначение:  Открытие сторонних (внешних) книг Excel и чтение данных из них
'              без изменения активной книги пользователя. Открывает файлы
'              в фоновом режиме (без экрана) и гарантированно закрывает их.
'
' Пример использования:
'   Dim reader As New clsExcelReader
'   reader.Init logger
'   If reader.OpenFile("C:\data\input.xlsx") Then
'       Dim data As Variant
'       data = reader.ReadRange("Лист1", "A1:C100")
'       reader.CloseFile
'   End If
'==============================================================================

Private Const MODULE_NAME As String = "clsExcelReader"

Private mLogger As ILogger
Private mExcelApp As Object      ' отдельный экземпляр Excel.Application
Private mWorkbook As Object      ' Excel.Workbook
Private mOwnsAppInstance As Boolean
Private mFilePath As String

Public Sub Init(ByVal Logger As ILogger)
    Set mLogger = Logger
End Sub

' -----------------------------------------------------------------------------
' Открывает файл Excel только для чтения, в отдельном скрытом экземпляре
' приложения (чтобы не мешать книге, из которой запущен макрос).
' -----------------------------------------------------------------------------
Public Function OpenFile(ByVal FilePath As String) As Boolean
    On Error GoTo ErrHandler

    If Len(Dir(FilePath)) = 0 Then
        LogErr "OpenFile", "Файл не найден: " & FilePath
        OpenFile = False
        Exit Function
    End If

    Set mExcelApp = CreateObject("Excel.Application")
    mExcelApp.Visible = False
    mExcelApp.DisplayAlerts = False
    mOwnsAppInstance = True

    Set mWorkbook = mExcelApp.Workbooks.Open(FilePath, ReadOnly:=True)
    mFilePath = FilePath

    LogInfo "OpenFile: открыт файл " & FilePath
    OpenFile = True
    Exit Function

ErrHandler:
    LogErr "OpenFile", Err.Number & " - " & Err.Description
    OpenFile = False
End Function

' -----------------------------------------------------------------------------
' Считывает диапазон в двумерный массив Variant (индексация с 1).
' -----------------------------------------------------------------------------
Public Function ReadRange(ByVal SheetName As String, ByVal RangeAddress As String) As Variant
    On Error GoTo ErrHandler

    Dim ws As Object
    Set ws = mWorkbook.Worksheets(SheetName)

    ReadRange = ws.Range(RangeAddress).Value
    LogInfo "ReadRange: " & SheetName & "!" & RangeAddress
    Exit Function

ErrHandler:
    LogErr "ReadRange", Err.Number & " - " & Err.Description
    ReadRange = Empty
End Function

' -----------------------------------------------------------------------------
' Считывает лист целиком (по используемому диапазону) как массив.
' -----------------------------------------------------------------------------
Public Function ReadSheet(ByVal SheetName As String) As Variant
    On Error GoTo ErrHandler

    Dim ws As Object
    Set ws = mWorkbook.Worksheets(SheetName)

    ReadSheet = ws.UsedRange.Value
    LogInfo "ReadSheet: " & SheetName
    Exit Function

ErrHandler:
    LogErr "ReadSheet", Err.Number & " - " & Err.Description
    ReadSheet = Empty
End Function

' -----------------------------------------------------------------------------
' Читает лист как список Collection словарей (заголовки из первой строки).
' Удобно для построчной обработки без работы с индексами массива.
' Требует Microsoft Scripting Runtime (или используйте Collection/Object).
' -----------------------------------------------------------------------------
Public Function ReadSheetAsRecords(ByVal SheetName As String) As Collection
    Dim result As New Collection
    Dim raw As Variant
    raw = ReadSheet(SheetName)

    If IsEmpty(raw) Then
        Set ReadSheetAsRecords = result
        Exit Function
    End If

    Dim headers() As String
    Dim colCount As Long, rowCount As Long
    Dim i As Long, j As Long

    rowCount = UBound(raw, 1)
    colCount = UBound(raw, 2)
    ReDim headers(1 To colCount)

    For j = 1 To colCount
        headers(j) = CStr(raw(1, j))
    Next j

    For i = 2 To rowCount
        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        For j = 1 To colCount
            rec(headers(j)) = raw(i, j)
        Next j
        result.Add rec
    Next i

    Set ReadSheetAsRecords = result
End Function

' -----------------------------------------------------------------------------
' Получить значение одной ячейки.
' -----------------------------------------------------------------------------
Public Function GetCellValue(ByVal SheetName As String, ByVal CellAddress As String) As Variant
    On Error GoTo ErrHandler
    GetCellValue = mWorkbook.Worksheets(SheetName).Range(CellAddress).Value
    Exit Function

ErrHandler:
    LogErr "GetCellValue", Err.Number & " - " & Err.Description
    GetCellValue = Empty
End Function

' -----------------------------------------------------------------------------
' Закрывает книгу и завершает временный экземпляр Excel.
' -----------------------------------------------------------------------------
Public Sub CloseFile()
    On Error Resume Next

    If Not mWorkbook Is Nothing Then
        mWorkbook.Close SaveChanges:=False
        Set mWorkbook = Nothing
    End If

    If mOwnsAppInstance And Not mExcelApp Is Nothing Then
        mExcelApp.Quit
        Set mExcelApp = Nothing
    End If

    LogInfo "CloseFile: файл закрыт (" & mFilePath & ")"
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Me.CloseFile
End Sub

' --- Внутренние помощники ---------------------------------------------------

Private Sub LogInfo(ByVal Message As String)
    If Not mLogger Is Nothing Then mLogger.Log llInfo, MODULE_NAME & ": " & Message
End Sub

Private Sub LogErr(ByVal MethodName As String, ByVal Message As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llError, MODULE_NAME & "." & MethodName & ": " & Message
    End If
End Sub
