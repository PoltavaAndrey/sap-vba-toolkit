Attribute VB_Name = "modMain"
Option Explicit

'==============================================================================
' Module:      modMain
' Purpose:     Macro entry point. This is where the code connects to the
'              user interface (a button on a sheet, an Application.Run
'              call, a menu item, etc.).
'==============================================================================

' Runs the SAP + Excel scenario using the default paths from modConfig.
Public Sub RunDefault()
    RunScenario modConfig.DEFAULT_INPUT_FILE, modConfig.DEFAULT_OUTPUT_FILE
End Sub

' Runs the scenario with explicitly given input/output file paths. Can be
' called from a button on a sheet or via Application.Run from another
' workbook.
Public Sub RunScenario(ByVal InputFilePath As String, ByVal OutputFilePath As String)
    Dim app As New clsAppController
    app.Init
    app.Run InputFilePath, OutputFilePath
End Sub

' Runs the reservation/stock matching scenario, reading both data sets
' straight from SAP table controls. Adjust the two control IDs to match the
' transaction you are automating.
Public Sub RunStockMatching()
    Dim app As New clsAppController
    app.Init
    app.RunStockMatchingFromSapTables _
        "wnd[0]/usr/tblRESERVATIONS_TABLE", _
        "wnd[0]/usr/tblSTOCK_TABLE", _
        "C:\Data\stock_matching_report.xlsx"
End Sub

' Quick sanity check that the SAP GUI Scripting connection works at all,
' without running the full scenario - handy for first-time environment
' diagnostics.
Public Sub TestSapConnection()
    Dim logger As New clsLogger
    logger.Init MinLevel:=llDebug

    Dim connector As New clsSapConnector
    connector.Init logger

    If connector.Connect() Then
        MsgBox "Connected to SAP GUI successfully.", vbInformation
    Else
        MsgBox "Could not connect to SAP GUI. See the log file for details: " & _
               logger.LogFilePath, vbExclamation
    End If

    logger.CloseLog
End Sub
