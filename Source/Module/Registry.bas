Attribute VB_Name = "basRegistry"
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
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const KEY_QUERY_VALUE = &H1
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, _
   ByVal lpData As Long, ByVal lpcbData As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
   (ByVal hKey As Long, ByVal lpClass As String, _
   lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
   lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
   lpftLastWriteTime As Any) As Long
   
'To update windows Icon Cache
Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

' A file type association has changed.
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Public Sub SaveKey(hKey As Long, strPath As String)
    Dim lngKeyHand As Long
    Call RegCreateKey(hKey, strPath, lngKeyHand)
    Call RegCloseKey(lngKeyHand)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim lngKeyHand As Long
    Dim lngValueType As Long
    Dim lngResult As Long
    Dim strBuf As String
    Dim lngDataBufSize As Long
    Dim intZeroPos As Integer
    Call RegOpenKey(hKey, strPath, lngKeyHand)
    lngResult = RegQueryValueEx(lngKeyHand, strValue, 0&, lngValueType, ByVal 0&, lngDataBufSize)
    If lngValueType = REG_SZ Then
        strBuf = String(lngDataBufSize, " ")
        lngResult = RegQueryValueEx(lngKeyHand, strValue, 0&, 0&, ByVal strBuf, lngDataBufSize)
        If lngResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
    Call RegCloseKey(lngKeyHand)
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim lngKeyHand As Long
    Call RegCreateKey(hKey, strPath, lngKeyHand)
    Call RegSetValueEx(lngKeyHand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    Call RegCloseKey(lngKeyHand)
End Sub

Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lngResult As Long
    Dim lngValueType As Long
    Dim lngBuffer As Long
    Dim lngDataBufSize As Long
    Dim lngKeyHand As Long
    Call RegOpenKey(hKey, strPath, lngKeyHand)
    lngDataBufSize = 4
    lngResult = RegQueryValueEx(lngKeyHand, strValueName, 0&, lngValueType, lngBuffer, lngDataBufSize)
    If lngResult = ERROR_SUCCESS Then
        If lngValueType = REG_DWORD Then
            GetDWord = lngBuffer
        End If
    End If
    Call RegCloseKey(lngKeyHand)
End Function

Function SaveDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lngResult As Long
    Dim lngKeyHand As Long
    Call RegCreateKey(hKey, strPath, lngKeyHand)
    lngResult = RegSetValueEx(lngKeyHand, strValueName, 0&, REG_DWORD, lData, 4)
    Call RegCloseKey(lngKeyHand)
End Function

Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    Call RegDeleteKey(hKey, strKey)
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim lngKeyHand As Long
    Call RegOpenKey(hKey, strPath, lngKeyHand)
    Call RegDeleteValue(lngKeyHand, strValue)
    Call RegCloseKey(lngKeyHand)
End Function

Public Function EnumKeys(ByVal hKey As Long, ByVal strPath As String) As String()
Dim lngKeyHand As Long
Dim aKeys() As String
Dim strBuff As String
Dim lngBuff As Long
Dim k As Long
    Call RegOpenKey(hKey, strPath, lngKeyHand)
    For k = 0 To 199
        lngBuff = 255
        strBuff = Space(lngBuff)
        Call RegEnumKey(lngKeyHand, k, strBuff, lngBuff)
        strBuff = Trim(Replace(strBuff, Chr(0), " "))
        If strBuff = "" Then Exit For
        ReDim Preserve aKeys(k)
        aKeys(k) = strBuff
    Next
    Call RegCloseKey(lngKeyHand)
    EnumKeys = aKeys
End Function

Public Function EnumValues(ByVal hBaseKey As Long, ByVal strPath As String) As String()
    Dim lResult                 As Long
    Dim hKey                    As Long
    Dim sName                   As String
    Dim lNameSize               As Long
    Dim lIndex                  As Long
    Dim cJunk                   As Long
    Dim cNameMax                As Long
    Dim ft                      As Currency
    Dim iKeyCount               As Integer
    Dim sKeyNames()             As String
   ' Log "EnterEnumValues"

   On Error GoTo EnumValuesError
   lIndex = 0
   lResult = RegOpenKeyEx(hBaseKey, strPath, 0, KEY_QUERY_VALUE, hKey)
   If (lResult = ERROR_SUCCESS) Then
      ' Log "OpenedKey:" & m_hClassKey & "," & m_sSectionKey
      lResult = RegQueryInfoKey(hKey, "", cJunk, 0, _
                               cJunk, cJunk, cJunk, cJunk, _
                               cNameMax, cJunk, cJunk, ft)
       Do While lResult = ERROR_SUCCESS
   
           'Set buffer space
           lNameSize = cNameMax + 1
           sName = String$(lNameSize, 0)
           If (lNameSize = 0) Then lNameSize = 1
           
           ' Log "Requesting Next Value"
         
           'Get value name:
           lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, _
                                  0&, 0&, 0&, 0&)
           ' Log "RegEnumValue returned:" & lResult
           If (lResult = ERROR_SUCCESS) Then
       
                ' Although in theory you can also retrieve the actual
                ' value and type here, I found it always (ultimately) resulted in
                ' a GPF, on Win95 and NT.  Why?  Can anyone help?
       
               sName = Left$(sName, lNameSize)
               ' Log "Enumerated value:" & sName
                 
               ReDim Preserve sKeyNames(0 To iKeyCount) As String
               sKeyNames(iKeyCount) = sName
               iKeyCount = iKeyCount + 1
           End If
           lIndex = lIndex + 1
       Loop
   End If
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If

   ' Log "Exit Enumerate Values"
   EnumValues = sKeyNames
   Exit Function
   
EnumValuesError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   Err.Raise vbObjectError + 1048 + 26003, App.EXEName & ".cRegistry", Err.Description
   Exit Function
End Function

Public Function IsAssociated() As Boolean
    IsAssociated = (GetString(HKEY_CLASSES_ROOT, ".post", "Content Type") <> "")
End Function

Public Sub Associate(ByVal Action As Boolean)
On Error Resume Next
    If Action Then
        'ignore if already associated
        If IsAssociated Then Exit Sub
        'create an entry in the class key
        Call SaveString(HKEY_CLASSES_ROOT, ".post", "", "postfile")
        'content type
        Call SaveString(HKEY_CLASSES_ROOT, ".post", "Content Type", "text/plain")
        'name
        Call SaveString(HKEY_CLASSES_ROOT, "postfile", "", "Blog Post File")
        'edit flags
        Call SaveDWord(HKEY_CLASSES_ROOT, "postfile", "EditFlags", "0000")
        'file's icon (can be an icon file, or an icon located within a dll file)
        'in this example, I am using a resource icon in this exe.
        Call SaveString(HKEY_CLASSES_ROOT, "postfile\DefaultIcon", "", App.Path & "\" & App.EXEName & ".exe,1")
        'Shell
        Call SaveString(HKEY_CLASSES_ROOT, "postfile\Shell", "", "")
        'Shell Open
        Call SaveString(HKEY_CLASSES_ROOT, "postfile\Shell\Open", "", "")
        'Shell open command
        Call SaveString(HKEY_CLASSES_ROOT, "postfile\Shell\Open\Command", "", App.Path & "\" & App.EXEName & ".exe %1")
        'Update the Windows Icon Cache to see our icon right away:
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    Else
        'ignore if not associated
        If Not IsAssociated Then Exit Sub
        'delete the registry keys
        Call DeleteKey(HKEY_CLASSES_ROOT, ".post")
        Call DeleteKey(HKEY_CLASSES_ROOT, "postfile\DefaultIcon")
        Call DeleteKey(HKEY_CLASSES_ROOT, "postfile\Shell\Open\Command")
        Call DeleteKey(HKEY_CLASSES_ROOT, "postfile\Shell\Open")
        Call DeleteKey(HKEY_CLASSES_ROOT, "postfile\Shell")
        Call DeleteKey(HKEY_CLASSES_ROOT, "postfile")
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    End If
End Sub

Function MediaPlayerInfo(ByVal sKey As String) As String
    Dim lResult As Long, hKey As Long, dwType As Long
    Dim szBuffer As String, lBuffSize As Long
    
    ' Open the key.
    lResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\MediaPlayer\CurrentMetadata", 0, 1, hKey)
    
     If lResult <> ERROR_SUCCESS Then GoTo NoNews
    ' Query the value.
    
    ' Determine how large the buffer needs to be
    lResult = RegQueryValueEx(hKey, sKey, 0&, dwType, ByVal szBuffer, lBuffSize)
    If lResult = ERROR_SUCCESS Then
       ' Build buffer and get data
       If lBuffSize > 0 Then
          szBuffer = Space(lBuffSize)
          lResult = RegQueryValueEx(hKey, sKey, 0&, dwType, ByVal szBuffer, Len(szBuffer))
          If lResult = ERROR_SUCCESS Then
             ' Trim NULL and return successful query!
             MediaPlayerInfo = Left(szBuffer, lBuffSize - 1)
          End If
       End If
    End If
    RegCloseKey (hKey)

NoNews:

End Function

