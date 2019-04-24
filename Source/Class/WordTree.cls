VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWordTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Public Root As Collection
Public Sub Add(strCode As String, strWord As String)
    Dim WordCollection As Collection
    Set WordCollection = New Collection
    Dim i As Long
    On Error GoTo e
     ' Add collection to hold words with give
     ' soundex code.
     ' If this code is already in the collection, there will be
     ' an error so skip to e:
    Root.Add WordCollection, strCode
e:
    ' Add the new word to the collection for the given code
    Set WordCollection = Root.Item(strCode)
    WordCollection.Add strWord
End Sub
' Returns a collection of words that have the same soundex code
' as the given code.
' If the code isn't in the Root collection, nothing is returned
Public Property Get GetCodeNode(strCode As String) As Collection
On Error Resume Next
    Set GetCodeNode = Root(strCode)
End Property

Private Sub Class_Initialize()
    Set Root = New Collection
End Sub

Public Property Get Count() As Long
   Count = Root.Count
End Property
