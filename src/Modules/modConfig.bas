Attribute VB_Name = "modConfig"
Option Explicit

'==============================================================================
' Module:      modConfig
' Purpose:     Centralized constants and project settings. Change values
'              here rather than scattering them across the codebase.
'==============================================================================

Public Const APP_NAME As String = "SAP VBA Toolkit"
Public Const APP_VERSION As String = "0.2.0"

' Default paths (used if the caller does not supply its own).
Public Const DEFAULT_LOG_SUBFOLDER As String = "Logs"

' Example file paths - adjust to your task, or move to a UI/parameters.
Public Const DEFAULT_INPUT_FILE As String = "C:\Data\input.xlsx"
Public Const DEFAULT_OUTPUT_FILE As String = "C:\Data\output.xlsx"
