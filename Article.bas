Attribute VB_Name = "Module1"
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const MAX_PATH = 260
Public Const BLOCK_SIZE = 10000

Global Cmd As String 'Data Command
Global Tbl As String 'Data Table
Global sCmdSource As String
Global sPath As String

Public Function TemporaryFileName() As String
Dim temp_path As String
Dim temp_file As String
Dim length As Long

    ' Get the temporary file path
    temp_path = VBA.Space$(MAX_PATH)
    length = GetTempPath(MAX_PATH, temp_path)
    temp_path = Left$(temp_path, length)

    ' Get the file name
    temp_file = VBA.Space$(MAX_PATH)
    GetTempFileName temp_path, "per", 0, temp_file
    TemporaryFileName = Left$(temp_file, InStr(temp_file, VBA.Chr$(0)) - 1)
End Function



