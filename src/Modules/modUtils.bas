Attribute VB_Name = "modUtils"
Option Explicit

'==============================================================================
' Module:      modUtils
' Назначение:  Общие вспомогательные функции, не привязанные к конкретному
'              классу (проверка файлов, работа со строками и т.п.).
'==============================================================================

' Проверяет существование файла на диске.
Public Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = (Len(Dir(FilePath)) > 0)
End Function

' Проверяет существование папки; при отсутствии создает ее (в т.ч. вложенные
' уровни, которых еще нет).
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

' Возвращает текущую метку времени в формате, удобном для имен файлов.
Public Function TimestampForFileName() As String
    TimestampForFileName = Format(Now, "yyyymmdd_hhnnss")
End Function

' Безопасное приведение к строке (Null/Empty -> "").
Public Function SafeStr(ByVal Value As Variant) As String
    If IsNull(Value) Or IsEmpty(Value) Then
        SafeStr = vbNullString
    Else
        SafeStr = CStr(Value)
    End If
End Function
