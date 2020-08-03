Attribute VB_Name = "Module1"
Public Const MAX_PATH = 260
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

Declare Function FindFirstFile Lib "kernel32" _
Alias "FindFirstFileA" (ByVal lpFileName As String, _
lpFindFileData As WIN32_FIND_DATA) As Long
  
Declare Function FindNextFile Lib "kernel32" _
Alias "FindNextFileA" (ByVal hFindFile As Long, _
lpFindFileData As WIN32_FIND_DATA) As Long
  
Declare Function FindClose Lib "kernel32" _
(ByVal hFindFile As Long) As Long


