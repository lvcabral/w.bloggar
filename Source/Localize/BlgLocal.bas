Attribute VB_Name = "basBlgLocal"
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
Public Const LOCALE_SLANGUAGE       As Long = &H2  'localized name of language
Public Const LOCALE_SABBREVLANGNAME As Long = &H3  'abbreviated language name
Public Const LCID_INSTALLED         As Long = &H1  'installed locale ids
Public Const LCID_SUPPORTED         As Long = &H2  'supported locale ids
Public Const LCID_ALTERNATE_SORTS   As Long = &H4  'alternate sort locale ids

Public LCID As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long

Public Declare Function EnumSystemLocales Lib "kernel32" _
  Alias "EnumSystemLocalesA" _
  (ByVal lpLocaleEnumProc As Long, _
   ByVal dwFlags As Long) As Long

Public Sub Main()
    If Not FileExists(App.Path + "\wbloggar.len") Then
        MsgBox "The file wbloggar.len is missing, please reinstall de translator tool!", vbExclamation
        End
    End If
    frmBlgLocal.Show
End Sub

Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim nSize As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If nSize Then
    
     'pad a buffer with spaces
      sReturn = Space$(nSize)
       
     'and call again passing the buffer
      nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (nSize > 0)
      If nSize Then
      
        'nSize holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, nSize - 1)
      
      End If
   
   End If
    
End Function

Public Function EnumSystemLocalesProc(lpLocaleString As Long) As Long

  'application-defined callback function for EnumSystemLocales

   Dim pos As Integer
   Dim dwLocaleDec As Long
   Dim dwLocaleHex As String
   Dim sLocaleName As String
     
  'pad a string to hold the format
   dwLocaleHex = Space$(32)
   
  'copy the string pointed to by the return value
   CopyMemory ByVal dwLocaleHex, lpLocaleString, ByVal Len(dwLocaleHex)
   
  'locate the terminating null
   pos = InStr(dwLocaleHex, Chr$(0))
   
   If pos Then
     'strip the null
      dwLocaleHex = Left$(dwLocaleHex, pos - 1)
      
     'we need the last 4 chrs - this
     'is the locale ID in hex
      dwLocaleHex = (Right$(dwLocaleHex, 4))
      
     'convert the string to a long
      dwLocaleDec = CLng("&H" & dwLocaleHex)
      
     'get the language and abbreviation for that locale
      sLocaleName = GetUserLocaleInfo(dwLocaleDec, LOCALE_SLANGUAGE)
   End If
   
   'add the data to the list
   If FileExists(App.Path & "\" & Format(dwLocaleDec) & ".lng") Then
      frmBlgLocal.cboLCID.AddItem Chr(149) + " " + sLocaleName & " - " & dwLocaleDec
   Else
      frmBlgLocal.cboLCID.AddItem sLocaleName & " - " & dwLocaleDec
   End If
  frmBlgLocal.cboLCID.ItemData(frmBlgLocal.cboLCID.NewIndex) = dwLocaleDec
   
  'and return 1 to continue enumeration
   EnumSystemLocalesProc = 1
   
End Function

Public Function SearchItemData(ctrComboBox As Object, ByVal ItemData As Long) As Boolean
On Error GoTo ErrorHandler
Dim Index As Long
    For Index = 0 To ctrComboBox.ListCount - 1
        If ctrComboBox.ItemData(Index) = ItemData Then
            ctrComboBox.ListIndex = Index
            SearchItemData = True
            Exit For
        End If
    Next
    Exit Function
ErrorHandler:
    SearchItemData = False
End Function

Public Function FileExists(ByVal Arquivo As String) As Boolean
    Dim nArq%
    nArq% = FreeFile
    On Error Resume Next
    Open Arquivo$ For Input As #nArq%
    If Err = 0 Then
       FileExists = True
    Else
       FileExists = False
    End If
    Close #nArq%
End Function

