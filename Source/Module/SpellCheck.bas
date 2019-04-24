Attribute VB_Name = "basSpellCheck"
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
Public WordTree As clsWordTree
Public WordsToAdd() As String
Public bFileNotFound As Boolean

Function Soundex(sWord As String) As String
    Dim Num As String * 4 ' Holds the generated code
    Dim i As Long
    Dim J As Long
    Dim sChar As String
    Dim sLastChar As String
    On Error GoTo ErrHandler
    sLastChar = UCase$(Mid$(sWord, 1, 1)) ' Get the first letter
    LSet Num = sLastChar & "000"
        
    ' Create the code starting at the second letter.
    J = 2
    For i = 2 To Len(sWord)
        sChar = UCase$(Mid$(sWord, i, 1))
        ' If two letters that are the same are next to each other
        ' only count one of them
        If Not (StrComp(sLastChar, sChar) = 0) Then
            If InStr("BFPV", sChar) Then
               Mid$(Num, J, 1) = "1"
               J = J + 1
            ElseIf InStr("CGJKQSXZ", sChar) Then
               Mid$(Num, J, 1) = "2"
               J = J + 1
            ElseIf InStr("DT", sChar) Then
               Mid$(Num, J, 1) = "3"
               J = J + 1
            ElseIf StrComp(sChar, "L") = 0 Then
               Mid$(Num, J, 1) = "4"
               J = J + 1
            ElseIf InStr("MN", sChar) Then
               Mid$(Num, J, 1) = "5"
               J = J + 1
            ElseIf StrComp(sChar, "R") = 0 Then
               Mid$(Num, J, 1) = "6"
               J = J + 1
            End If
            sLastChar = sChar
        End If
        ' Don't try to build a SoundEx > 4 chars long
        If J = 5 Then Exit For
    Next
    
    ' Return 4 char SoundEx code:
    Soundex = Num
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "vbwSpellCheck", Err.Description
End Function
