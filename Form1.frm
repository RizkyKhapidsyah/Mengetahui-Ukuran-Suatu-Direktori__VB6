VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mengetahui Ukuran Suatu Direktori"
   ClientHeight    =   3090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Properties"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function SizeOf(ByVal DirPath As String) As Double
Dim hFind As Long
Dim fdata As WIN32_FIND_DATA
Dim dblSize As Double
Dim sName As String
Dim x As Long
On Error Resume Next
  x = GetAttr(DirPath)
  If Err Then SizeOf = 0: Exit Function
  If (x And vbDirectory) = vbDirectory Then
     dblSize = 0
     Err.Clear
     sName = Dir$(EndSlash(DirPath) & "*.*", vbSystem _
            Or vbHidden Or vbDirectory)
     If Err.Number = 0 Then
        hFind = FindFirstFile(EndSlash(DirPath) & _
                "*.*", fdata)
       If hFind = 0 Then Exit Function
        Do
        If (fdata.dwFileAttributes And vbDirectory) = _
           vbDirectory Then
           sName = Left$(fdata.cFileName, _
                InStr(fdata.cFileName, vbNullChar) - 1)
           If sName <> "." And sName <> ".." Then
              dblSize = dblSize + _
                      SizeOf(EndSlash(DirPath) & sName)
            End If
          Else
            dblSize = dblSize + fdata.nFileSizeHigh * _
                      65536 + fdata.nFileSizeLow
          End If
          DoEvents
        Loop While FindNextFile(hFind, fdata) <> 0
        hFind = FindClose(hFind)
     End If
  Else
     On Error Resume Next
     dblSize = FileLen(DirPath)
  End If
  SizeOf = dblSize
End Function

Private Function EndSlash(ByVal PathIn As String) As String
  If Right$(PathIn, 1) = "\" Then
     EndSlash = PathIn
  Else
     EndSlash = PathIn & "\"
  End If
End Function

Private Sub Command1_Click()
  'Ganti 'C:\Windows' di bawah dengan nama direktori
  'yang ingin Anda ketahui ukurannya.
  MsgBox "Ukuran direktori C:\Windows = " & Format(SizeOf("C:\Windows"), "#,#") & " bytes", vbInformation, "Ukuran Direktori"
End Sub


