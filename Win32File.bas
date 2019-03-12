Attribute VB_Name = "Win32File"
'
' Win32File
' Copyright (c) 2016-2017 Timothy Baxendale (pcluddite@outlook.com)
' 
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
' 
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'
Option Explicit
'
' Last update: 10/3/17
' Requires: N/A
' Description: A set of functions based on the Windows API to manage the file system
'

Private Type FILETIME
    dwLowDate  As Long
    dwHighDate As Long
End Type

Private Type SYSTEMTIME
    wYear      As Integer
    wMonth     As Integer
    wDayOfWeek As Integer
    wDay       As Integer
    wHour      As Integer
    wMinute    As Integer
    wSecond    As Integer
    wMillisecs As Integer
End Type

Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternateFileName As String * 14
End Type

Private EMPTY_FILETIME As FILETIME

Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
Private Const INVALID_HANDLE_VALUE = -1&
Private Const GENERIC_READ = &H80000000

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
   (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function LocalFileTimeToFileTime Lib "kernel32" _
   (lpLocalFileTime As FILETIME, _
    lpFileTime As FILETIME) As Long
    
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" _
   (lpFileTime As FILETIME, _
    lpLocalFileTime As FILETIME) As Long

Private Declare PtrSafe Function SetFileTime Lib "kernel32" _
   (ByVal hFile As LongPtr, _
    ByVal lpCreationTime As LongPtr, _
    ByVal lpLastAccessTime As LongPtr, _
    ByVal lpLastWriteTime As LongPtr) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME, _
    lpFileTime As FILETIME) As Long
    
Private Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, _
    lpSystemTime As SYSTEMTIME) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
   (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Boolean

Private Declare PtrSafe Function GetFileTime Lib "kernel32" _
    (ByVal hFile As LongPtr, _
     ByVal lpCreationTime As LongPtr, _
     ByVal lpLastAccessTime As LongPtr, _
     ByVal lpLastWriteTime As LongPtr) As Boolean

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" _
    (ByVal lpFileName As String, _
     ByVal nBufferLength As Long, _
     ByVal lpBuffer As String, _
     ByVal lpFilePart As String) As Long
     
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" _
    (ByVal lpExistingFileName As String, _
     ByVal lpNewFileName As String) As Long

Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" _
   (lpTimeZone As LongPtr, _
    lpUniversalTime As SYSTEMTIME, _
    lpLocalTime As SYSTEMTIME) As Long

'
' Gets the file extension of a path
'
Property Get FileExt(ByVal Path As String) As String
    Dim nIdx As Long
    Let nIdx = InStrRev(Path, ".")
    If nIdx <> Empty Then
        FileExt = Right(Path, Len(Path) - nIdx + 1)
    End If
End Property

'
' Gets only the file name from a path (excluding the extension)
'
Property Get FileBaseName(ByVal Path As String) As String
    Dim Ext As String, Start As Long
    Ext = FileExt(Path)
    If Ext <> Empty Then
        Start = InStrRev(Path, "\") + 1
        If Start < 1 Then Start = 1
        FileBaseName = Mid(Path, Start, Len(Path) - Start - Len(Ext) + 1)
    End If
End Property

'
' Lists all files in a directory
'
Function ListFiles(ByVal Directory As String, Optional ByVal Attributes As VbFileAttribute = vbNormal) As String()
    Dim Files As String, File As String
    File = Dir(Directory, Attributes)
    Files = File
    Do While True
        File = Dir
        If File = "" Then Exit Do
        Files = Files & ";" & File
    Loop
    ListFiles = Split(Files, ";")
End Function

'
' Determines if a file exists
'
Function FileExists(ByVal Path As String) As Boolean
    Dim FindFileData As WIN32_FIND_DATA
    Dim hFile As Long

    hFile = FindFirstFile(Path, FindFileData)
    FileExists = hFile <> INVALID_HANDLE_VALUE

    If FileExists Then FindClose hFile
End Function

'
' Determines if a file exists
'
Function DirExists(ByVal Path As String) As Boolean
    If Right(Path, 1) <> "\" Then
        DirExists = Dir(Path & "\", vbNormal) <> ""
    Else
        DirExists = Dir(Path, vbNormal) <> ""
    End If
End Function

'
' Gets the path of the directory in which a file or folder resides
'
Property Get FilePath(ByVal Path As String) As String
    FilePath = Mid(Path, 1, InStrRev(Path, "\") - 1)
End Property

'
' Builds a path from two strings
'
Function PathCombine(ByVal Path1 As String, ByVal Path2 As String)
    Dim SlashEnds As Boolean, SlashStarts As Boolean
    SlashEnds = (Right(Path1, 1) = "\")
    SlashStarts = (Left(Path2, 1) = "\")
    
    If SlashEnds Xor SlashStarts Then
        PathCombine = Path1 & Path2
    ElseIf SlashEnds And SlashStarts Then
        PathCombine = Path1 & Mid(Path2, 2)
    Else
        PathCombine = Path1 & "\" & Path2
    End If
End Function

'
' Set the time that a file was last modified
' FilePath - the path to the file
' dtDate - the last write time
'
Property Let FileLastWriteTime(ByVal FilePath As String, ByVal dtDate As Date)
    Dim lpFileTime As FILETIME
    lpFileTime = DateToFileTime(dtDate)
    SetFileTimeVBA FilePath, 0, 0, VarPtr(lpFileTime)
End Property

'
' Set the time that a file was created
' FilePath - the path to the file
' dtDate - the creation time
'
Property Let FileCreationTime(ByVal FilePath As String, ByVal dtDate As Date)
    Dim lpFileTime As FILETIME
    lpFileTime = DateToFileTime(dtDate)
    SetFileTimeVBA FilePath, VarPtr(lpFileTime), 0, 0
End Property

'
' Set the time that a file was created
' FilePath - the path to the file
' dtDate - the creation time
'
Property Let FileLastAccessTime(ByVal FilePath As String, ByVal dtDate As Date)
    Dim lpFileTime As FILETIME
    lpFileTime = DateToFileTime(dtDate)
    SetFileTimeVBA FilePath, 0, VarPtr(lpFileTime), 0
End Property

'
' Get the time that a file was last modified
'
Property Get FileLastWriteTime(ByVal FilePath As String) As Date
    Dim lpFileTime As FILETIME
    GetFileTimeVBA FilePath, 0, 0, VarPtr(lpFileTime)
    FileLastWriteTime = FileTimeToDate(lpFileTime)
End Property

'
' Get the time that a file was last accessed
'
Property Get FileLastAccessTime(ByVal FilePath As String) As Date
    Dim lpFileTime As FILETIME
    GetFileTimeVBA FilePath, 0, VarPtr(lpFileTime), 0
    FileLastAccessTime = FileTimeToDate(lpFileTime)
End Property

'
' Get the time that a file was created
'
Property Get FileCreationTime(ByVal FilePath As String) As Date
    Dim lpFileTime As FILETIME
    GetFileTimeVBA FilePath, VarPtr(lpFileTime), 0, 0
    FileCreationTime = FileTimeToDate(lpFileTime)
End Property

Private Function DateToFileTime(ByVal dtDate As Date) As FILETIME
    Dim lRet As Long
    Dim lpFileTime As FILETIME
    Dim lpLocalTime As FILETIME
    Dim lpSystemTime As SYSTEMTIME
    
    With lpSystemTime
        .wYear = Year(dtDate)
        .wMonth = Month(dtDate)
        .wDay = Day(dtDate)
        .wDayOfWeek = Weekday(dtDate) - 1
        .wHour = Hour(dtDate)
        .wMinute = Minute(dtDate)
        .wSecond = Second(dtDate)
    End With
    
    lRet = SystemTimeToFileTime(lpSystemTime, lpLocalTime)
    lRet = LocalFileTimeToFileTime(lpLocalTime, DateToFileTime)
End Function

Private Function FileTimeToDate(ByRef lpFileTime As FILETIME) As Date
    Dim lRet As Long
    Dim lpLocal As FILETIME
    Dim lpSysTime As SYSTEMTIME
    
    lRet = FileTimeToLocalFileTime(lpFileTime, lpLocal)
    lRet = FileTimeToSystemTime(lpLocal, lpSysTime)
    'lRet = SystemTimeToTzSpecificLocalTime(0, lpSysTime, lpLocal)
    
    With lpSysTime
        FileTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Private Function SetFileTimeVBA(ByVal FilePath As String, ByVal lpCreateTime As LongPtr, ByVal lpAccessTime As LongPtr, ByVal lpLastWriteTime As LongPtr) As Boolean
    Dim lRet As Long, hFile As Long
    
    hFile = CreateFile(FilePath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    lRet = SetFileTime(hFile, lpCreateTime, lpAccessTime, lpLastWriteTime)
    CloseHandle hFile
    
    SetFileTimeVBA = lRet <> 0
End Function

Private Function GetFileTimeVBA(ByVal FilePath As String, ByRef lpCreateTime As LongPtr, ByRef lpAccessTime As LongPtr, ByRef lpLastWriteTime As LongPtr) As Boolean
    Dim lRet As Long, hFile As Long
    
    hFile = CreateFile(FilePath, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    lRet = GetFileTime(hFile, lpCreateTime, lpAccessTime, lpLastWriteTime)
    CloseHandle hFile
    
    GetFileTimeVBA = lRet <> 0
End Function

'
' Get the full path name of a file from a relative path
'
Property Get FileFullName(ByVal RelativePath As String) As String
    Dim nRet As Long
    nRet = GetFullPathNameVBA(RelativePath, FileFullName)
End Property

'
' Get the full path name of a file from a relative path
'
Property Get FileName(ByVal Path As String) As String
    Dim nRet As Long
    nRet = GetFullPathNameVBA(Path, lpFilePart:=FileName)
End Property

Private Function GetFullPathNameVBA(ByVal RelativePath As String, Optional ByRef lpBuff As String, Optional ByRef lpFilePart As String) As Long
    Dim nLen As Long
    
    nLen = MAX_PATH
    lpBuff = String(nLen, 0)
    lpFilePart = String(nLen, 0)
    GetFullPathNameVBA = GetFullPathName(RelativePath, nLen, lpBuff, lpFilePart)
    
    If GetFullPathNameVBA > nLen Then
        nLen = GetFullPathNameVBA + 10
        lpBuff = String(nLen, 0)
        lpFilePart = String(nLen, 0)
        GetFullPathNameVBA = GetFullPathName(RelativePath, nLen, lpBuff, lpFilePart)
    End If
End Function

'
' Move a file from one location to another. Use this function to rename.
'
Function FileMove(ByVal Source As String, ByVal Destination As String) As Boolean
    FileMove = MoveFile(Source, Destination)
End Function
