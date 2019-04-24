Attribute VB_Name = "basShell"
'*    w.bloggar
'*    Copyright (C) 2001-2019 Marcelo Lv Cabral <https://lvcabral.com>
'*
'*    This program is free software; you can redistribute it and/or modify
'*    it under the terms of the GNU General Public License as published by
'*    the Free Software Foundation; either version 2 of the License, or
'*    (at your option) any later version.
'*
'*    This program is distributed in the hope that it will be useful,
'*    but WITHOUT ANY WARRANTY; without even the implied warranty of
'*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'*    GNU General Public License for more details.
'*
'*    You should have received a copy of the GNU General Public License along
'*    with this program; if not, write to the Free Software Foundation, Inc.,
'*    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
Option Explicit

Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_PRINTHOOD = &H1B

Private Type SHITEMID
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long

Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pszPath As String)

Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
    
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4

Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_SILENT = &H4
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_WANTMAPPINGHANDLE = &H20

Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_NOERRORUI = &H400
Private Const SHARD_PATH = &H2&

' GetDriveType return values
Private Const DRIVE_NO_ROOT_DIR = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const SHFMT_OPT_FULL = &H1
Private Const SHFMT_OPT_SYSONLY = &H2

Public Sub AddToRecentDocs(strFileName As String)
  ' Comments  : Adds a file to the 'Documents' submenu on the
  '             Windows Start menu
  ' Parameters: strFileName - full path to the document. The file must
  '             have a registered extension
  ' Returns   : Nothing
  
  '
  On Error GoTo PROC_ERR
  
  SHAddToRecentDocs SHARD_PATH, strFileName

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "AddToRecentDocs"
  Resume PROC_EXIT
  
End Sub

Public Sub ClearRecentDocs()
  ' Comments  : Clears the list of recently-opened documents
  '             from the Windows Start menu
  ' Parameters: None
  ' Returns   : Nothing
  
  '
  On Error GoTo PROC_ERR

  SHAddToRecentDocs SHARD_PATH, vbNullString

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ClearRecentDocs"
  Resume PROC_EXIT

End Sub

Public Function GetShellAppdataLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Appdata" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Appdata folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_APPDATA, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellAppdataLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellAppdataLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonAppdataLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the All User's
  '             "Appdata" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Appdata folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_APPDATA, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonAppdataLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonAppdataLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonDesktopDirectoryLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonDesktopDirectory" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonDesktopDirectory folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_DESKTOPDIRECTORY, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonDesktopDirectoryLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonDesktopDirectoryLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonProgramsLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonPrograms" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonPrograms folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_PROGRAMS, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonProgramsLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonProgramsLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonStartMenuLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonStartMenu" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonStartMenu folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_STARTMENU, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonStartMenuLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonStartMenuLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellCommonStartupLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "CommonStartup" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's CommonStartup folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_COMMON_STARTUP, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellCommonStartupLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellCommonStartupLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellDesktopDirectoryLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "DesktopDirectory" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's DesktopDirectory folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_DESKTOPDIRECTORY, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellDesktopDirectoryLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellDesktopDirectoryLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellDesktopLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Desktop" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Desktop folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_DESKTOP, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellDesktopLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellDesktopLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellFavoritesLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Favorites" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Favorites folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_FAVORITES, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
        
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellFavoritesLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellFavoritesLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellFontsLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Fonts" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Fonts folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_FONTS, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellFontsLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellFontsLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellPersonalLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Personal" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Personal folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_PERSONAL, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellPersonalLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellPersonalLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellProgramsLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Programs" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Programs folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_PROGRAMS, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellProgramsLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellProgramsLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellRecentLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Recent" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Recent folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_RECENT, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
        
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellRecentLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellRecentLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellSendToLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "SendTo" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's SendTo folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_SENDTO, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellSendToLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellSendToLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellStartMenuLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "StartMenu" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's StartMenu folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_STARTMENU, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellStartMenuLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellStartMenuLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellStartupLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Startup" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Startup folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_STARTUP, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
          
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellStartupLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellStartupLocation"
  Resume PROC_EXIT
  
End Function

Public Function GetShellTemplatesLocation( _
  lnghWnd As Long) _
  As String
  ' Comments  : Returns the path of the user's
  '             "Templates" folder
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  ' Returns   : Path of the user's Templates folder
  
  '
  Dim lngResult As Long
  Dim strPath As String
  Dim idlist As ITEMIDLIST
  
  On Error GoTo PROC_ERR
    
  ' populate an ITEMIDLIST struct with the specified folder information
  lngResult = SHGetSpecialFolderLocation( _
    lnghWnd, CSIDL_TEMPLATES, idlist)
    
  If lngResult = 0 Then
    
    ' if the information is present, get the path information
    strPath = Space$(260)
    lngResult = SHGetPathFromIDList( _
        ByVal idlist.mkid.cb, _
        ByVal strPath)
      
    ' free memory allocated by shell
    CoTaskMemFree idlist.mkid.cb
    
    'if a path was found, trim off trailing null char
    strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    GetShellTemplatesLocation = strPath
    
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetShellTemplatesLocation"
  Resume PROC_EXIT
  
End Function

Public Sub ShellCopyFile( _
  lnghWnd As Long, _
  ByVal strSource As String, _
  ByVal strDestination As String, _
  Optional ByVal fSilent As Boolean = False, _
  Optional strTitle As String = "")
  ' Comments  : Copies a file or files to a single destination
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  '             strSource - file spec for files to copy
  '             strDestination - destination file name or directory
  '             fSilent - if true, no warnings are displayed
  '             strTitle - title of the progress dialog
  ' Returns   : Nothing
  
  ' Update    : Code Service Pack 3
  '
  Dim foCopy As SHFILEOPSTRUCT
  Dim lngFlags As Long
  Dim lngResult As Long
  Dim lngStructLen As Long
  Dim abytBuf() As Byte
    
  On Error GoTo PROC_ERR
  
  ' check to be sure file exists
  If Dir$(strSource) <> "" Then
    
    ' set flags for no prompting
    If fSilent Then
      lngFlags = FOF_NOCONFIRMMKDIR Or FOF_NOCONFIRMATION Or FOF_SILENT
    End If
    
    lngStructLen = LenB(foCopy)
    ReDim abytBuf(1 To lngStructLen)
  
    ' set shell file operations settings
    With foCopy
      .hwnd = lnghWnd
      .pFrom = strSource & vbNullChar & vbNullChar
      .pTo = strDestination & vbNullChar & vbNullChar
      .fFlags = lngFlags
      .lpszProgressTitle = strTitle & vbNullChar & vbNullChar
      .wFunc = FO_COPY
      
      If strTitle <> "" Then
        .fFlags = .fFlags Or FOF_SIMPLEPROGRESS
        
        ' Adjust alignment by copying to byte array
        Call CopyMemory(abytBuf(1), foCopy, lngStructLen)
        Call CopyMemory(abytBuf(19), abytBuf(21), 12)
      
        lngResult = SHFileOperation(abytBuf(1))
      Else
        lngResult = SHFileOperation(foCopy)
      End If
    
    End With
    
  End If

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellCopyFile"
  Resume PROC_EXIT

End Sub

Public Sub ShellMoveFile( _
  lnghWnd As Long, _
  ByVal strSource As String, _
  ByVal strDestination As String, _
  Optional ByVal fSilent As Boolean = False, _
  Optional strTitle As String = "")
  ' Comments  : Copies a file or files to a single destination
  ' Parameters: lnghWnd - handle to window to serve as
  '             the parent for the dialog. Use a form's
  '             hWnd property for example
  '             strSource - file spec for files to copy
  '             strDestination - destination file name or directory
  '             fSilent - if true, no warnings are displayed
  '             strTitle - title of the progress dialog
  ' Returns   : Nothing
  
  ' Update    : Code Service Pack 3
  '
  Dim foMove As SHFILEOPSTRUCT
  Dim lngFlags As Long
  Dim lngResult As Long
  Dim lngStructLen As Long
  Dim abytBuf() As Byte
    
  On Error GoTo PROC_ERR
  
  ' check to be sure file exists
  If Dir$(strSource) <> "" Then
    
    ' set flags for no prompting
    If fSilent Then
      lngFlags = FOF_NOCONFIRMMKDIR Or FOF_NOCONFIRMATION Or FOF_SILENT
    End If
    
    lngStructLen = LenB(foMove)
    ReDim abytBuf(1 To lngStructLen)
  
    ' set shell file operations settings
    With foMove
      .hwnd = lnghWnd
      .pFrom = strSource & vbNullChar & vbNullChar
      .pTo = strDestination & vbNullChar & vbNullChar
      .fFlags = lngFlags
      .lpszProgressTitle = strTitle & vbNullChar & vbNullChar
      .wFunc = FO_MOVE
      
      If strTitle <> "" Then
        .fFlags = .fFlags Or FOF_SIMPLEPROGRESS
        
        ' Adjust alignment by copying to byte array
        Call CopyMemory(abytBuf(1), foMove, lngStructLen)
        Call CopyMemory(abytBuf(19), abytBuf(21), 12)
      
        lngResult = SHFileOperation(abytBuf(1))
      Else
        lngResult = SHFileOperation(foMove)
      End If
    
    End With
    
  End If

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellMoveFile"
  Resume PROC_EXIT

End Sub

Public Sub ShellRecycleFile( _
  lnghWnd As Long, _
  ByVal strFileSpec As String, _
  Optional fUndoable As Boolean = True, _
  Optional strTitle As String = "")
  ' Comments  : Sends the specified file or files
  '             to the Windows 95/NT recycle bin
  ' Parameters: lnghWnd - handle to window to serve as the parent for the
  '             dialog. Use a form's hWnd property for example
  '             strFileSpec - full path to the file(s) todelete. May include
  '             wildcard characters
  '             fUndoable - If true, the files are permanently deleted
  '             strTitle - title of the progress dialog
  ' Returns   : Nothing
  
  ' Update    : Code Service Pack 3
  '
  Dim foDelete As SHFILEOPSTRUCT
  Dim lngResult As Long
  Dim lngFlags As Long
  Dim lngStructLen As Long
  Dim abytBuf() As Byte
      
  On Error GoTo PROC_ERR
      
  ' skip empty file specs
  If Not strFileSpec = vbNullString Then
  
    lngStructLen = LenB(foDelete)
    ReDim abytBuf(1 To lngStructLen)
  
    ' set optional flag to permanently delete the files
    If fUndoable = True Then
      lngFlags = FOF_ALLOWUNDO
    End If
    
    With foDelete
      .hwnd = lnghWnd
      .wFunc = FO_DELETE
      .pFrom = strFileSpec & vbNullChar & vbNullChar
      .fFlags = lngFlags
      .lpszProgressTitle = strTitle & vbNullChar & vbNullChar
    
      If strTitle <> "" Then
        .fFlags = .fFlags Or FOF_SIMPLEPROGRESS
        
        ' Adjust alignment by copying to byte array
        Call CopyMemory(abytBuf(1), foDelete, lngStructLen)
        Call CopyMemory(abytBuf(19), abytBuf(21), 12)
      
        lngResult = SHFileOperation(abytBuf(1))
      Else
        lngResult = SHFileOperation(foDelete)
      End If

    End With

  End If
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellRecycleFile"
  Resume PROC_EXIT
  
End Sub

Public Function FolderExists(Pasta As String) As Boolean
On Error GoTo ErrorHandler
  FolderExists = Len(Dir$(Pasta & "\.", vbDirectory)) > 0
  Exit Function

ErrorHandler:
  FolderExists = False
End Function

Public Function CreatePath(ByVal Path As String) As Boolean
Dim NewLen As Integer
Dim DirLen As Integer
Dim MaxLen As Integer
    NewLen = 4
    MaxLen = Len(Path)
    If Right$(Path, 1) <> "\" Then
        Path = Path + "\"
        MaxLen = MaxLen + 1
    End If
    On Error Resume Next
    Do While True
        DirLen = InStr(NewLen, Path, "\")
        Err = 0
        MkDir Left$(Path, DirLen - 1)
        CreatePath = (Err = 0)
        NewLen = DirLen + 1
        If NewLen >= MaxLen Then Exit Do
    Loop
End Function

Public Function GetTempFolder() As String
    Dim sTmpPath As String * 512
    Dim nRet As Long

    nRet = GetTempPath(512, sTmpPath)
    If (nRet > 0 And nRet < 512) Then
        GetTempFolder = Left$(sTmpPath, InStr(sTmpPath, vbNullChar) - 1)
    End If
End Function


