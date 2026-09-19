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
' Purpose:     Creates new Excel workbooks and writes data into them (e.g. a
'              summary report of SAP processing results). Works in a separate
'              (hidden by default) Excel instance.
'
' Example:
'   Dim writer As New clsExcelWriter
'   writer.Init logger
'   writer.CreateNewWorkbook
'   writer.WriteHeaders "Report", Array("Date", "Material", "Status")
'   writer.WriteRow "Report", 2, Array(Date, "1000123", "OK")
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
' Creates a new blank workbook in a separate Excel instance.
' -----------------------------------------------------------------------------
Public Function CreateNewWorkbook() As Boolean
    On Error GoTo ErrHandler

    Set mExcelApp = CreateObject("Excel.Application")
    mExcelApp.Visible = mVisible
    mExcelApp.DisplayAlerts = False

    Set mWorkbook = mExcelApp.Workbooks.Add

    LogInfo "CreateNewWorkbook: new workbook created"
    CreateNewWorkbook = True
    Exit Function

ErrHandler:
    LogErr "CreateNewWorkbook", Err.Number & " - " & Err.Description
    CreateNewWorkbook = False
End Function

' -----------------------------------------------------------------------------
' Returns a sheet by name, creating it if it does not exist yet.
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
' Writes headers into the first row of a sheet (creating the sheet if needed).
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
' Writes a single row of data (an array of values) into the given row of a
' sheet.
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
' Writes a whole 2D array starting at the given top-left cell. Convenient for
' dumping data produced by clsExcelReader.ReadRange/ReadSheet or by
' clsStockReservationMatcher.FindMatches.
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
' Auto-fits column widths on a sheet.
' -----------------------------------------------------------------------------
Public Sub AutoFitColumns(ByVal SheetName As String)
    On Error Resume Next
    mWorkbook.Worksheets(SheetName).Columns.AutoFit
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' Saves the workbook to the given path. Format is derived from the file
' extension (.xlsx = 51, .xlsm = 52, .xls = 56, .csv = 6).
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

' --- Internal helpers ---------------------------------------------------------

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
