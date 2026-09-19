Attribute VB_Name = "modDemo"
Option Explicit

'==============================================================================
' Module:      modDemo
' Purpose:     Stand-alone demonstration of clsStockReservationMatcher using
'              hard-coded sample data - no SAP connection required. Useful
'              for quickly understanding/testing the matching algorithm, or
'              as a starting point for unit tests.
'
' Run:         modDemo.RunStockMatchingDemo (F5 in the VBA editor), then
'              check the Immediate Window (Ctrl+G) for the output.
'==============================================================================

Public Sub RunStockMatchingDemo()
    Dim logger As New clsLogger
    logger.Init MinLevel:=llDebug, EchoImmediate:=True

    ' Reservations: [Plant, Warehouse, Material, Project, Quantity]
    Dim reservations As Variant
    reservations = Array( _
        Array("1000", "WH-A", "MAT-001", "PRJ-100", 50), _
        Array("1000", "WH-B", "MAT-001", "PRJ-200", 30), _
        Array("1000", "WH-A", "MAT-002", "PRJ-100", 10), _
        Array("2000", "WH-C", "MAT-001", "PRJ-300", 20) _
    )

    ' Stock: [Plant, StorageLocation, Material, Project, Quantity]
    Dim stock As Variant
    stock = Array( _
        Array("1000", "SL-01", "MAT-001", "PRJ-900", 40), _
        Array("1000", "SL-02", "MAT-001", "PRJ-910", 60), _
        Array("1000", "SL-01", "MAT-002", "PRJ-900", 5), _
        Array("2000", "SL-05", "MAT-999", "PRJ-920", 100) _
    )

    Dim matcher As New clsStockReservationMatcher
    matcher.Init logger

    Dim leftoverReservations As Variant, leftoverStock As Variant
    Dim matches As Variant
    matches = matcher.FindMatches( _
        NestedArrayTo2D(reservations, 5), _
        NestedArrayTo2D(stock, 5), _
        leftoverReservations, leftoverStock)

    Debug.Print "--- Matches ---"
    PrintArray matches

    Debug.Print "--- Unclosed reservations (Plant, Warehouse, Material, Project, OriginalQty, RemainingQty) ---"
    PrintArray leftoverReservations

    Debug.Print "--- Remaining stock (Plant, StorageLocation, Material, Project, OriginalQty, RemainingQty) ---"
    PrintArray leftoverStock

    logger.CloseLog
End Sub

' --- Helpers for this demo only ----------------------------------------------

' Converts an array of 1D row-arrays (as written above with Array(...)) into
' a proper 1-based 2D Variant array, the shape expected by
' clsStockReservationMatcher.FindMatches.
Private Function NestedArrayTo2D(ByVal Rows As Variant, ByVal ColCount As Long) As Variant
    Dim n As Long
    n = UBound(Rows) - LBound(Rows) + 1

    Dim result() As Variant
    ReDim result(1 To n, 1 To ColCount)

    Dim i As Long, c As Long
    For i = 1 To n
        For c = 1 To ColCount
            result(i, c) = Rows(i - 1 + LBound(Rows))(c - 1)
        Next c
    Next i

    NestedArrayTo2D = result
End Function

Private Sub PrintArray(ByVal Data As Variant)
    If IsEmpty(Data) Then
        Debug.Print "(empty)"
        Exit Sub
    End If

    Dim r As Long, c As Long
    Dim line As String
    For r = LBound(Data, 1) To UBound(Data, 1)
        line = vbNullString
        For c = LBound(Data, 2) To UBound(Data, 2)
            line = line & CStr(Data(r, c)) & vbTab
        Next c
        Debug.Print line
    Next r
End Sub
