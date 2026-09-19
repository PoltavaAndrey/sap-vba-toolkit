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
' Purpose:     Application facade. Wires together the logger, the SAP
'              connector, and the Excel classes, and exposes one clear Run()
'              method to kick off a scenario. This is the only place where
'              concrete implementations are created (New clsLogger, New
'              clsSapConnector, etc.) - the rest of the code talks only to
'              interfaces/wrappers.
'
' Example (see also modMain.bas):
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

    mLogger.Log llInfo, MODULE_NAME & ".Init: application initialized"
End Sub

' -----------------------------------------------------------------------------
' Example end-to-end scenario:
'   1. Read an input Excel file with a list of items to process in SAP.
'   2. Connect to SAP GUI and run a transaction for every row.
'   3. Write the outcome (success/error) into a new Excel report.
'
' This method is intentionally left as a template - replace the transaction
' code, SAP field IDs and input file structure with the real ones from your
' process.
' -----------------------------------------------------------------------------
Public Sub Run(ByVal InputFilePath As String, ByVal OutputFilePath As String)
    On Error GoTo ErrHandler

    mLogger.Log llInfo, "=== Scenario started ==="

    ' 1. Read input data -------------------------------------------------------
    Dim reader As New clsExcelReader
    reader.Init mLogger

    If Not reader.OpenFile(InputFilePath) Then
        mLogger.Log llError, "Could not open the input file. Stopping."
        Exit Sub
    End If

    Dim records As Collection
    Set records = reader.ReadSheetAsRecords("Sheet1")
    reader.CloseFile

    mLogger.Log llInfo, "Rows read: " & records.Count

    ' 2. Connect to SAP and process rows ---------------------------------------
    If Not mConnector.Connect() Then
        mLogger.Log llError, "Could not connect to SAP GUI. Stopping."
        Exit Sub
    End If

    Dim sapSession As clsSapSession
    Set sapSession = mConnector.GetSession(0)

    Dim writer As New clsExcelWriter
    writer.Init mLogger
    writer.CreateNewWorkbook
    writer.WriteHeaders "Result", Array("Material", "Status", "SAP message")

    Dim i As Long
    Dim rec As Object
    i = 1
    For Each rec In records
        i = i + 1
        ProcessSingleRecord sapSession, rec, writer, i
    Next rec

    ' 3. Save the report --------------------------------------------------------
    writer.AutoFitColumns "Result"
    writer.SaveAs OutputFilePath
    writer.CloseFile

    mLogger.Log llInfo, "=== Scenario finished. Report: " & OutputFilePath & " ==="
    Exit Sub

ErrHandler:
    mLogger.Log llError, MODULE_NAME & ".Run: unexpected error " & _
        Err.Number & " - " & Err.Description
End Sub

' -----------------------------------------------------------------------------
' Processes a single input row in SAP. Shown here for transaction ME21N
' (create purchase order) - replace with your own transaction and fields.
' -----------------------------------------------------------------------------
Private Sub ProcessSingleRecord(ByVal sapSession As clsSapSession, _
                                 ByVal rec As Object, _
                                 ByVal writer As clsExcelWriter, _
                                 ByVal OutputRow As Long)

    Dim material As String
    material = CStr(rec("Material"))

    mLogger.Log llInfo, "Processing material: " & material

    sapSession.StartTransaction "ME21N"
    sapSession.SetText "wnd[0]/usr/ctxtEKKO-BSART", "NB"
    sapSession.SetText "wnd[0]/usr/tabsHEADER_DETAIL/.../ctxtEKPO-MATNR", material
    sapSession.SendVKey 0 ' Enter

    Dim statusText As String
    statusText = sapSession.GetStatusBarText()

    Dim isOk As Boolean
    isOk = (InStr(1, LCase(statusText), "error") = 0)

    writer.WriteRow "Result", OutputRow, _
        Array(material, IIf(isOk, "OK", "ERROR"), statusText)

    sapSession.EndTransaction
End Sub

' -----------------------------------------------------------------------------
' Example scenario built around SAP table controls and the array matcher:
'   1. Open a SAP transaction that shows a table control (e.g. a reservation
'      list, or a stock overview).
'   2. Extract the on-screen table into a Variant array via clsSapTable.
'   3. Repeat for the second table (or read it from an Excel file instead).
'   4. Run clsStockReservationMatcher.FindMatches to see which reservations
'      can be closed with the available stock, and export the result.
'
' This is a template showing how the pieces fit together - the actual
' transaction, table control ID and column IDs must be adjusted to match
' your SAP system.
' -----------------------------------------------------------------------------
Public Sub RunStockMatchingFromSapTables(ByVal ReservationsTableId As String, _
                                          ByVal StockTableId As String, _
                                          ByVal OutputFilePath As String)
    On Error GoTo ErrHandler

    mLogger.Log llInfo, "=== Stock matching scenario started ==="

    If Not mConnector.Connect() Then
        mLogger.Log llError, "Could not connect to SAP GUI. Stopping."
        Exit Sub
    End If

    Dim sapSession As clsSapSession
    Set sapSession = mConnector.GetSession(0)

    ' Adjust the column IDs to match the technical names of your table
    ' controls (found via SAP's "Script Recording and Playback" helper).
    Dim reservationColumns As Variant
    reservationColumns = Array("WERKS", "LGORT", "MATNR", "PSPNR", "MENGE")

    Dim stockColumns As Variant
    stockColumns = Array("WERKS", "LGORT", "MATNR", "PSPNR", "MENGE")

    Dim reservationsTable As clsSapTable
    Set reservationsTable = sapSession.GetTable(ReservationsTableId)
    Dim reservations As Variant
    reservations = reservationsTable.ReadAllRows(reservationColumns)

    Dim stockTable As clsSapTable
    Set stockTable = sapSession.GetTable(StockTableId)
    Dim stock As Variant
    stock = stockTable.ReadAllRows(stockColumns)

    Dim matcher As New clsStockReservationMatcher
    matcher.Init mLogger

    Dim leftoverReservations As Variant, leftoverStock As Variant
    Dim matches As Variant
    matches = matcher.FindMatches(reservations, stock, leftoverReservations, leftoverStock)

    Dim writer As New clsExcelWriter
    writer.Init mLogger
    writer.CreateNewWorkbook

    writer.WriteHeaders "Matches", Array("Plant", "StorageLocation", "Material", _
        "Project", "QtyOnStock", "QtyReserved", "ReservationProject", "ReservationWarehouse")
    If Not IsEmpty(matches) Then writer.WriteArray "Matches", "A2", matches

    If Not IsEmpty(leftoverReservations) Then
        writer.WriteHeaders "UnclosedReservations", Array("Plant", "Warehouse", "Material", _
            "Project", "OriginalQty", "RemainingQty")
        writer.WriteArray "UnclosedReservations", "A2", leftoverReservations
    End If

    writer.AutoFitColumns "Matches"
    writer.SaveAs OutputFilePath
    writer.CloseFile

    mLogger.Log llInfo, "=== Stock matching scenario finished. Report: " & OutputFilePath & " ==="
    Exit Sub

ErrHandler:
    mLogger.Log llError, MODULE_NAME & ".RunStockMatchingFromSapTables: unexpected error " & _
        Err.Number & " - " & Err.Description
End Sub
