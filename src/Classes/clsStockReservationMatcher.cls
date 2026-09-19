VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStockReservationMatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Class:       clsStockReservationMatcher
' Purpose:     Compares two in-memory data arrays -
'                 1) Reservations : [Plant, Warehouse,       Material, Project, Quantity]
'                 2) Stock        : [Plant, StorageLocation,  Material, Project, Quantity]
'              - and finds reservation lines that can be fully or partially
'              closed using available stock, producing a third array of
'              matches:
'                 [StockPlant, StorageLocation, Material, Project,
'                  QtyOnStock, QtyReserved, ReservationProject, ReservationWarehouse]
'
' Matching key: Plant + Material only. Project and Warehouse/StorageLocation
'              are NOT required to match - they are carried through into the
'              result purely for reference, so the person reviewing the
'              report can see which project/warehouse the reservation came
'              from versus which project/location the stock sits under.
'              If your business rules require matching by Project as well,
'              see the "Customization" note below.
'
' Allocation logic: for every reservation row (in input order) the matcher
'              looks for stock rows with the same Plant+Material and unused
'              quantity, and allocates min(remaining reservation qty,
'              remaining stock qty) to it. One output row is produced per
'              non-zero allocation. A single reservation may be split across
'              several stock rows (multiple output rows), and a single stock
'              row may be used to close several reservations, until either
'              side runs out of quantity. This greedy, first-come-first-served
'              strategy is a reasonable default for a template - swap in
'              FIFO-by-date, priority-by-project, etc. as needed.
'
' Input arrays: 1-based 2D Variant arrays with exactly 5 columns and NO
'              header row (strip headers before calling FindMatches, e.g.
'              when the data comes from clsExcelReader.ReadRange, which
'              includes headers if you read them as part of the range).
'
' Example:
'   Dim matcher As New clsStockReservationMatcher
'   matcher.Init logger
'
'   Dim matches As Variant
'   Dim leftoverReservations As Variant, leftoverStock As Variant
'   matches = matcher.FindMatches(reservations, stock, leftoverReservations, leftoverStock)
'
'   writer.WriteHeaders "Matches", Array("Plant", "StorageLocation", "Material", _
'       "Project", "QtyOnStock", "QtyReserved", "ReservationProject", "ReservationWarehouse")
'   writer.WriteArray "Matches", "A2", matches
'==============================================================================

Private Const MODULE_NAME As String = "clsStockReservationMatcher"

' Column layout of the RESERVATIONS input array (1-based).
Public Enum ReservationColumn
    rcPlant = 1
    rcWarehouse = 2
    rcMaterial = 3
    rcProject = 4
    rcQuantity = 5
End Enum

' Column layout of the STOCK input array (1-based).
Public Enum StockColumn
    scPlant = 1
    scStorageLocation = 2
    scMaterial = 3
    scProject = 4
    scQuantity = 5
End Enum

' Column layout of the RESULT array produced by FindMatches (1-based).
Public Enum ResultColumn
    resPlant = 1                  ' Plant the stock belongs to
    resStorageLocation = 2        ' Stock storage location
    resMaterial = 3                ' Material
    resProject = 4                 ' Project on the STOCK side
    resQtyOnStock = 5               ' Quantity taken from this stock line for this match
    resQtyReserved = 6              ' Quantity of the reservation closed by this match
    resReservationProject = 7      ' Project on the RESERVATION side
    resReservationWarehouse = 8    ' Warehouse on the RESERVATION side
End Enum

Private mLogger As ILogger

Public Sub Init(ByVal Logger As ILogger)
    Set mLogger = Logger
End Sub

' -----------------------------------------------------------------------------
' Finds reservation/stock matches as described in the class header.
'
' Reservations, Stock  - 1-based 2D arrays, 5 columns each (see the Column
'                        enums above), no header row.
' RemainingReservations (ByRef, optional) - returns the reservation rows
'                        that could NOT be fully closed, as a 2D array with
'                        one extra trailing column: remaining quantity.
'                        Empty if every reservation was fully closed.
' RemainingStock (ByRef, optional) - returns the stock rows that still have
'                        quantity left after all allocations, with one extra
'                        trailing column: remaining quantity.
'
' Returns: a 2D array per ResultColumn, or Empty if there are no matches.
' -----------------------------------------------------------------------------
Public Function FindMatches(ByVal Reservations As Variant, ByVal Stock As Variant, _
                             Optional ByRef RemainingReservations As Variant, _
                             Optional ByRef RemainingStock As Variant) As Variant
    On Error GoTo ErrHandler

    RemainingReservations = Empty
    RemainingStock = Empty

    If IsEmpty(Reservations) Or IsEmpty(Stock) Then
        LogInfo "FindMatches: one of the input arrays is empty, nothing to do"
        FindMatches = Empty
        Exit Function
    End If

    Dim resRows As Long, stockRows As Long
    resRows = UBound(Reservations, 1) - LBound(Reservations, 1) + 1
    stockRows = UBound(Stock, 1) - LBound(Stock, 1) + 1

    Dim resLo As Long, stockLo As Long
    resLo = LBound(Reservations, 1)
    stockLo = LBound(Stock, 1)

    Dim resRemaining() As Double
    Dim stockRemaining() As Double
    ReDim resRemaining(1 To resRows)
    ReDim stockRemaining(1 To stockRows)

    Dim r As Long, s As Long
    For r = 1 To resRows
        resRemaining(r) = SafeQty(Reservations(r + resLo - 1, rcQuantity))
    Next r
    For s = 1 To stockRows
        stockRemaining(s) = SafeQty(Stock(s + stockLo - 1, scQuantity))
    Next s

    Dim results As New Collection

    For r = 1 To resRows
        If resRemaining(r) > 0 Then
            Dim resPlantKey As String, resMaterialKey As String
            resPlantKey = NormalizeKey(Reservations(r + resLo - 1, rcPlant))
            resMaterialKey = NormalizeKey(Reservations(r + resLo - 1, rcMaterial))

            For s = 1 To stockRows
                If resRemaining(r) <= 0 Then Exit For

                If stockRemaining(s) > 0 Then
                    If NormalizeKey(Stock(s + stockLo - 1, scPlant)) = resPlantKey And _
                       NormalizeKey(Stock(s + stockLo - 1, scMaterial)) = resMaterialKey Then

                        Dim allocated As Double
                        allocated = MinOf(resRemaining(r), stockRemaining(s))

                        If allocated > 0 Then
                            Dim row(1 To 8) As Variant
                            row(resPlant) = Stock(s + stockLo - 1, scPlant)
                            row(resStorageLocation) = Stock(s + stockLo - 1, scStorageLocation)
                            row(resMaterial) = Stock(s + stockLo - 1, scMaterial)
                            row(resProject) = Stock(s + stockLo - 1, scProject)
                            row(resQtyOnStock) = allocated
                            row(resQtyReserved) = allocated
                            row(resReservationProject) = Reservations(r + resLo - 1, rcProject)
                            row(resReservationWarehouse) = Reservations(r + resLo - 1, rcWarehouse)
                            results.Add row

                            resRemaining(r) = resRemaining(r) - allocated
                            stockRemaining(s) = stockRemaining(s) - allocated
                        End If
                    End If
                End If
            Next s
        End If
    Next r

    RemainingReservations = BuildRemainingArray(Reservations, resRemaining, resLo, 5)
    RemainingStock = BuildRemainingArray(Stock, stockRemaining, stockLo, 5)

    LogInfo "FindMatches: " & Reservations2Text(resRows) & " reservation row(s), " & _
        stockRows & " stock row(s) -> " & results.Count & " match(es)"

    FindMatches = CollectionToArray(results, 8)
    Exit Function

ErrHandler:
    LogErr "FindMatches", Err.Number & " - " & Err.Description
    FindMatches = Empty
End Function

' --- Internal helpers ---------------------------------------------------------

Private Function Reservations2Text(ByVal n As Long) As String
    Reservations2Text = CStr(n)
End Function

' Builds a "leftover rows" array: original columns + 1 trailing "remaining
' quantity" column, keeping only rows where remaining quantity is still > 0.
Private Function BuildRemainingArray(ByVal Source As Variant, ByRef Remaining() As Double, _
                                      ByVal SourceLo As Long, ByVal SourceCols As Long) As Variant
    Dim n As Long
    n = UBound(Remaining)

    Dim leftoverCount As Long
    leftoverCount = 0

    Dim i As Long
    For i = 1 To n
        If Remaining(i) > 0 Then leftoverCount = leftoverCount + 1
    Next i

    If leftoverCount = 0 Then
        BuildRemainingArray = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To leftoverCount, 1 To SourceCols + 1)

    Dim outRow As Long, c As Long
    outRow = 0
    For i = 1 To n
        If Remaining(i) > 0 Then
            outRow = outRow + 1
            For c = 1 To SourceCols
                result(outRow, c) = Source(i + SourceLo - 1, c)
            Next c
            result(outRow, SourceCols + 1) = Remaining(i)
        End If
    Next i

    BuildRemainingArray = result
End Function

' Converts a Collection of 1-row Variant arrays (1..ColCount) into a proper
' 2D array (1..Collection.Count, 1..ColCount).
Private Function CollectionToArray(ByVal Source As Collection, ByVal ColCount As Long) As Variant
    If Source.Count = 0 Then
        CollectionToArray = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To Source.Count, 1 To ColCount)

    Dim i As Long, c As Long
    Dim oneRow As Variant
    For i = 1 To Source.Count
        oneRow = Source(i)
        For c = 1 To ColCount
            result(i, c) = oneRow(c)
        Next c
    Next i

    CollectionToArray = result
End Function

' Normalizes a value for key comparison: trims whitespace and compares as
' text, so numeric vs. text representations of the same code (e.g. plant
' "1000" stored as a Double vs. a String) still match. Does NOT pad/strip
' leading zeros - keep Plant/Material formatting consistent between the two
' source arrays (e.g. both as text with leading zeros, or both as numbers).
Private Function NormalizeKey(ByVal Value As Variant) As String
    NormalizeKey = Trim$(CStr(Value))
End Function

Private Function SafeQty(ByVal Value As Variant) As Double
    On Error Resume Next
    SafeQty = CDbl(Value)
End Function

Private Function MinOf(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        MinOf = a
    Else
        MinOf = b
    End If
End Function

Private Sub LogInfo(ByVal Message As String)
    If Not mLogger Is Nothing Then mLogger.Log llInfo, MODULE_NAME & ": " & Message
End Sub

Private Sub LogErr(ByVal MethodName As String, ByVal Message As String)
    If Not mLogger Is Nothing Then
        mLogger.Log llError, MODULE_NAME & "." & MethodName & ": " & Message
    End If
End Sub
