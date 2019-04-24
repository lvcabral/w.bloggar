Attribute VB_Name = "UniFunctions"
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
  Public m_bIsNt     As Boolean
    
  Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
  Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
  Public Const CP_UTF8 = 65001
   
  'Purpose:Convert   Utf8   to   Unicode
  Public Function UTF8_Decode(ByVal sUTF8 As String) As String
    
        Dim lngUtf8Size As Long
        Dim strBuffer As String
        Dim lngBufferSize As Long
        Dim lngResult As Long
        Dim bytUtf8() As Byte
        Dim n As Long
    
        If LenB(sUTF8) = 0 Then Exit Function
        m_bIsNt = True
        If m_bIsNt Then
              On Error GoTo EndFunction
              bytUtf8 = StrConv(sUTF8, vbFromUnicode)
              lngUtf8Size = UBound(bytUtf8) + 1
              On Error GoTo 0
              lngBufferSize = lngUtf8Size * 2
              strBuffer = String$(lngBufferSize, vbNullChar)
              lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), _
                    lngUtf8Size, StrPtr(strBuffer), lngBufferSize)
              If lngResult Then
                    UTF8_Decode = Left$(strBuffer, lngResult)
              End If
        Else
              Dim i As Long
              Dim TopIndex As Long
              Dim TwoBytes(1) As Byte
              Dim ThreeBytes(2) As Byte
              Dim AByte As Byte
              Dim TStr As String
              Dim BArray() As Byte

              On Error Resume Next
              TopIndex = Len(sUTF8)
              If TopIndex = 0 Then Exit Function
              BArray = StrConv(sUTF8, vbFromUnicode)
              i = 0
              TopIndex = TopIndex - 1
              Do While i <= TopIndex
                    AByte = BArray(i)
                    If AByte < &H80 Then
                          TStr = TStr & Chr$(AByte): i = i + 1
                    ElseIf AByte >= &HE0 Then
                          ThreeBytes(0) = BArray(i): i = i + 1
                          ThreeBytes(1) = BArray(i): i = i + 1
                          ThreeBytes(2) = BArray(i): i = i + 1
                          TStr = TStr & ChrW$((ThreeBytes(0) And &HF) * &H1000 + (ThreeBytes(1) And &H3F) * &H40 + (ThreeBytes(2) And &H3F))
                    ElseIf (AByte >= &HC2) And (AByte <= &HDB) Then
                          TwoBytes(0) = BArray(i): i = i + 1
                          TwoBytes(1) = BArray(i): i = i + 1
                          TStr = TStr & ChrW$((TwoBytes(0) And &H1F) * &H40 + (TwoBytes(1) And &H3F))
                    Else
                          TStr = TStr & Chr$(AByte): i = i + 1
                    End If
              Loop
              UTF8_Decode = TStr
              Erase BArray
        End If
    
EndFunction:
    
  End Function
  
  Public Function UTF8_Encode(ByVal strUnicode As String, Optional ByVal bHTML As Boolean) As String
        Dim i As Long
        Dim TLen As Long
        Dim lPtr As Long
        Dim UTF16 As Long
        Dim UTF8_EncodeLong As String
    
        TLen = Len(strUnicode)
        If TLen = 0 Then Exit Function
    
        If m_bIsNt Then
              Dim lngBufferSize As Long
              Dim lngResult As Long
              Dim bytUtf8() As Byte
              lngBufferSize = TLen * 3 + 1
              ReDim bytUtf8(lngBufferSize - 1)
              lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), _
                    TLen, bytUtf8(0), lngBufferSize, vbNullString, 0)
              If lngResult Then
                    lngResult = lngResult - 1
                    ReDim Preserve bytUtf8(lngResult)
                    UTF8_Encode = StrConv(bytUtf8, vbUnicode)
              End If
        Else
              For i = 1 To TLen
                    lPtr = StrPtr(strUnicode) + ((i - 1) * 2)
                    CopyMemory UTF16, ByVal lPtr, 2
                    If UTF16 < &H80 Then
                          UTF8_EncodeLong = Chr$(UTF16)
                    ElseIf UTF16 < &H800 Then
                          UTF8_EncodeLong = Chr$(&H80 + (UTF16 And &H3F))
                          UTF16 = UTF16 \ &H40
                          UTF8_EncodeLong = Chr$(&HC0 + (UTF16 And &H1F)) & UTF8_EncodeLong
                    Else
                          UTF8_EncodeLong = Chr$(&H80 + (UTF16 And &H3F))
                          UTF16 = UTF16 \ &H40
                          UTF8_EncodeLong = Chr$(&H80 + (UTF16 And &H3F)) & UTF8_EncodeLong
                          UTF16 = UTF16 \ &H40
                          UTF8_EncodeLong = Chr$(&HE0 + (UTF16 And &HF)) & UTF8_EncodeLong
                    End If
                    UTF8_Encode = UTF8_Encode & UTF8_EncodeLong
              Next
        End If
        If bHTML Then
              UTF8_Encode = Replace$(UTF8_Encode, vbCrLf, "<br/>")
        End If
  End Function
