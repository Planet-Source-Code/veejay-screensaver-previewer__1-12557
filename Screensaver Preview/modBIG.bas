Attribute VB_Name = "modBIG"
Option Explicit

Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH As Long = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Const PROCESS_QUERY_INFORMATION = &H400

Public Sub FindFiles(inPath As String, inRecursive As Boolean)
' Finds all files with an extension of "*.scr" and adds
' them to a listbox, and a combobox
Dim filename As String, DirName As String
Dim i As Long, Continue As Integer, SearchStr As String
Dim hSearch As Long, wfd As WIN32_FIND_DATA
Dim List() As String, O As Long
SearchStr = "*.scr"
    If Right(inPath, 1) <> "\" Then inPath = inPath & "\"
    Continue = True
hSearch = FindFirstFile(inPath & SearchStr, wfd)
Continue = True
If hSearch <> INVALID_HANDLE_VALUE Then
    While Continue
        DoEvents
        filename = StripNulls(wfd.cFileName)
        If (filename <> ".") And (filename <> "..") Then
            ReDim Preserve List(O)
            List(O) = inPath & filename
            O = O + 1
        End If
        Continue = FindNextFile(hSearch, wfd)
    Wend
    Continue = FindClose(hSearch)
End If
SelectionSort List, 0, UBound(List)
For O = 0 To UBound(List)
    frmMain.lstPath.AddItem List(O)
    frmMain.cmbNames.AddItem GetFileName(List(O), False)
Next O
End Sub

Public Function GetSysDir() As String
    ' This function will get the path to the SystemDirectory
    Dim sSave As String, ret As Long
    sSave = Space(255)
    ret = GetSystemDirectory(sSave, 255)
    sSave = Left$(sSave, ret)
    GetSysDir = sSave
End Function

Public Function GetWinDir() As String
    ' This function will get the path to the Windows Directory
    Dim Gwdvar As String, Gwdvar_Length As Integer
    Gwdvar = Space(255)
    Gwdvar_Length = GetWindowsDirectory(Gwdvar, 255)
    GetWinDir = Left(Gwdvar, Gwdvar_Length)
End Function
