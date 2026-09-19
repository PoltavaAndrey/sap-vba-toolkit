VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSapTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class:       clsSapTable
' Purpose:     Wraps a classic dynpro table control (GuiTableControl) and
'              provides convenient methods for reading and writing cells,
'              scrolling, row selection, and bulk-extracting the whole table
'              into a 2D Variant array (which can then be fed into
'              clsExcelWriter or clsStockReservationMatcher).
'
' Scope:       This class targets GuiTableControl (the classic table control
'              used in transactions like ME21N/ME22N/ME23N item overview,
'              VA01/VA02, etc.). ALV Grid controls (GuiGridView), used in
'              many list/report transactions, expose a different scripting
'              API and are NOT covered here - contributions welcome.
'
' Example:
'   Dim table As clsSapTable
'   Set table = sapSession.GetTable("wnd[0]/usr/tblSAPLV45TCTRL_ITEM_OV")
'
'   ' Read a single cell
'   Dim material As String
'   material = table.GetCellText(0, "MATNR")
'
'   ' Write a single cell (row 2, column "MENGE")
'   table.SetCellText 2, "MENGE", "10"
'
'   ' Extract the whole table into an array for further processing/export
'   Dim data As Variant
'   data = table.ReadAllRows(Array("MATNR", "MENGE", "MEINS"))
'==============================================================================

Private Const MODULE_NAME As String = "clsSapTable"

Private mTable As Object        ' GuiTableControl
Private mLogger As ILogger

' -----------------------------------------------------------------------------
Public Sub Init(ByVal RawTableControl As Object, ByVal Logger As ILogger)
    Set mTable = RawTableControl
    Set mLogger = Logger
End Sub

' -----------------------------------------------------------------------------
' Total number of rows in the table (including rows not currently scrolled
' into view).
' -----------------------------------------------------------------------------
Public Property Get RowCount() As Long
    On Error Resume Next
    RowCount = mTable.RowCount
End Property

' -----------------------------------------------------------------------------
' Number of rows visible on screen at once (the scrollable "viewport" size).
' -----------------------------------------------------------------------------
Public Property Get VisibleRowCount() As Long
    On Error Resume Next
    VisibleRowCount = mTable.VisibleRowCount
End Property

' -----------------------------------------------------------------------------
' Number of columns in the table.
' -----------------------------------------------------------------------------
Public Property Get ColumnCount() As Long
    On Error Resume Next
    ColumnCount = mTable.Columns.Count
End Property

' -----------------------------------------------------------------------------
' Returns the technical column IDs (e.g. "MATNR", "MENGE") in display order.
' Useful for discovering column names before hard-coding them.
' -----------------------------------------------------------------------------
Public Function GetColumnIds() As Variant
    On Error GoTo ErrHandler

    Dim n As Long
    n = mTable.Columns.Count

    Dim ids() As String
    ReDim ids(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        ids(i) = mTable.Columns.ElementAt(i).Name
    Next i

    GetColumnIds = ids
    Exit Function

ErrHandler:
    LogError "GetColumnIds", "n/a"
    GetColumnIds = Empty
End Function

' -----------------------------------------------------------------------------
' Reads the text of a cell in a currently VISIBLE row (0-based, relative to
' the current scroll position - the same indexing GuiTableControl itself
' uses). To read an arbitrary row regardless of scroll position, use
' GetCellTextAbsolute instead.
' -----------------------------------------------------------------------------
Public Function GetCellText(ByVal VisibleRowIndex As Long, ByVal ColumnId As String) As String
    On Error GoTo ErrHandler
    GetCellText = mTable.GetCell(VisibleRowIndex, ColumnId).Text
    Exit Function

ErrHandler:
    LogError "GetCellText", ColumnId & " (row " & VisibleRowIndex & ")"
    GetCellText = vbNullString
End Function

' -----------------------------------------------------------------------------
' Writes the text of a cell in a currently visible, editable row/column.
' -----------------------------------------------------------------------------
Public Function SetCellText(ByVal VisibleRowIndex As Long, ByVal ColumnId As String, _
                             ByVal Value As String) As Boolean
    On Error GoTo ErrHandler
    LogAction "SetCellText(row=" & VisibleRowIndex & ", col=" & ColumnId & ", value=" & Value & ")"

    mTable.GetCell(VisibleRowIndex, ColumnId).Text = Value
    SetCellText = True
    Exit Function

ErrHandler:
    LogError "SetCellText", ColumnId & " (row " & VisibleRowIndex & ")"
    SetCellText = False
End Function

' -----------------------------------------------------------------------------
' Scrolls the table so that AbsoluteRowIndex (0-based, across the whole
' table) becomes visible as row 0.
' -----------------------------------------------------------------------------
Public Sub ScrollToRow(ByVal AbsoluteRowIndex As Long)
    On Error Resume Next
    mTable.VerticalScrollbar.Position = AbsoluteRowIndex
End Sub

' -----------------------------------------------------------------------------
' Reads the text of a cell by ABSOLUTE row index (0-based, across the whole
' table), scrolling the table as needed. Slower than GetCellText when reading
' many cells in a loop - prefer ReadAllRows for bulk extraction.
' -----------------------------------------------------------------------------
Public Function GetCellTextAbsolute(ByVal AbsoluteRowIndex As Long, ByVal ColumnId As String) As String
    On Error GoTo ErrHandler
    ScrollToRow AbsoluteRowIndex
    GetCellTextAbsolute = mTable.GetCell(0, ColumnId).Text
    Exit Function

ErrHandler:
    LogError "GetCellTextAbsolute", ColumnId & " (row " & AbsoluteRowIndex & ")"
    GetCellTextAbsolute = vbNullString
End Function

' -----------------------------------------------------------------------------
' Selects (highlights) a row by absolute index, e.g. before pressing a
' toolbar button that acts on the selected row.
' -----------------------------------------------------------------------------
Public Function SelectRow(ByVal AbsoluteRowIndex As Long) As Boolean
    On Error GoTo ErrHandler
    LogAction "SelectRow(" & AbsoluteRowIndex & ")"

    ScrollToRow AbsoluteRowIndex
    mTable.GetAbsoluteRow(AbsoluteRowIndex).Selected = True
    SelectRow = True
    Exit Function

ErrHandler:
    LogError "SelectRow", CStr(AbsoluteRowIndex)
    SelectRow = False
End Function

' -----------------------------------------------------------------------------
' Extracts the requested columns for ALL rows of the table into a 2D Variant
' array (1-based: rows 1..RowCount, columns 1..UBound(ColumnIds)+1), scrolling
' through the table one row at a time as needed.
'
' ColumnIds - a 1D array (or Array(...)) of technical column names, e.g.
'             Array("WERKS", "LGORT", "MATNR", "PSPNR", "MENGE")
'
' Note: this reads one absolute row per scroll step, which is simple and
' reliable but not the fastest possible approach; for very large tables
' consider reading VisibleRowCount rows per scroll step instead.
' -----------------------------------------------------------------------------
Public Function ReadAllRows(ByVal ColumnIds As Variant) As Variant
    On Error GoTo ErrHandler

    Dim totalRows As Long
    totalRows = Me.RowCount

    Dim colCount As Long
    colCount = UBound(ColumnIds) - LBound(ColumnIds) + 1

    If totalRows = 0 Then
        ReadAllRows = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To totalRows, 1 To colCount)

    Dim r As Long, c As Long
    For r = 0 To totalRows - 1
        ScrollToRow r
        For c = LBound(ColumnIds) To UBound(ColumnIds)
            result(r + 1, c - LBound(ColumnIds) + 1) = mTable.GetCell(0, ColumnIds(c)).Text
        Next c
    Next r

    LogAction "ReadAllRows: extracted " & totalRows & " row(s), " & colCount & " column(s)"
    ReadAllRows = result
    Exit Function

ErrHandler:
    LogError "ReadAllRows", "n/a"
    ReadAllRows = Empty
End Function

' -----------------------------------------------------------------------------
' Direct access to the "raw" GuiTableControl for non-standard cases not
' covered by the wrappers above.
' -----------------------------------------------------------------------------
Public Property Get RawTable() As Object
    Set RawTable = mTable
End Property

' --- Internal helpers ---------------------------------------------------------

Private Sub LogAction(ByVal Description As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llDebug, MODULE_NAME & ": " & Description
    End If
End Sub

Private Sub LogError(ByVal MethodName As String, ByVal Context As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llError, MODULE_NAME & "." & MethodName & _
            ": failed for '" & Context & "' (" & Err.Number & " - " & Err.Description & ")"
    End If
End Sub
