Attribute VB_Name = "modUtils"
Option Explicit

'==============================================================================
' Module:      modUtils
' Purpose:     General-purpose helper functions not tied to a specific class
'              (file checks, string helpers, etc.).
'==============================================================================

' Checks whether a file exists on disk.
Public Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = (Len(Dir(FilePath)) > 0)
End Function

' Checks whether a folder exists; creates it (including any missing parent
' levels) if it does not.
Public Sub EnsureFolderExists(ByVal FolderPath As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long

    parts = Split(FolderPath, "\")
    currentPath = parts(0)

    For i = 1 To UBound(parts)
        currentPath = currentPath & "\" & parts(i)
        If Len(Dir(currentPath, vbDirectory)) = 0 Then
            MkDir currentPath
        End If
    Next i
End Sub

' Returns the current timestamp in a format convenient for file names.
Public Function TimestampForFileName() As String
    TimestampForFileName = Format(Now, "yyyymmdd_hhnnss")
End Function

' Safely casts a value to String (Null/Empty -> "").
Public Function SafeStr(ByVal Value As Variant) As String
    If IsNull(Value) Or IsEmpty(Value) Then
        SafeStr = vbNullString
    Else
        SafeStr = CStr(Value)
    End If
End Function
