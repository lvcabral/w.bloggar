VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PostData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Public AccountID As Integer
Public BlogID As String
Public PostID As String
Public Title As String
Public Text As String
Public More As String
Public Excerpt As String
Public Keywords As String
Public DateTime As Variant
Public Author As String
Public Categories As String
Public AllowComments As Integer
Public AllowPings As Integer
Public TextFilter As String
Public TrackBack As String

Private Sub Class_Initialize()
    AccountID = -1
    DateTime = CDate(0)
    AllowComments = -1
    AllowPings = -1
End Sub

Public Sub LoadData(ByVal strFilePath As String)
Dim vValue As Variant
Dim DOMDocument As DOMDocument
Dim Param As IXMLDOMNode
Dim Value As IXMLDOMNode
Dim objPost As xmlStruct
Dim vResponse
On Error Resume Next
    If FileExists(strFilePath) Then
        Set DOMDocument = New DOMDocument
        DOMDocument.Load strFilePath
        '
        ' Convert the response into a variant array
        '
        Set Param = DOMDocument.selectSingleNode(".//param")
        If Param Is Nothing Then Exit Sub
        Set Value = Param.selectSingleNode(".//value")
        If Value Is Nothing Then Exit Sub
        ParseValue Value, vResponse
        Set objPost = vResponse
        Me.AccountID = objPost.Member("accountID").Value
        Me.BlogID = objPost.Member("blogID").Value
        Me.PostID = objPost.Member("postID").Value
        Me.Title = ConvertHTMLEntities(objPost.Member("textTitle").Value, False)
        Me.Text = ConvertHTMLEntities(objPost.Member("textBody").Value, False)
        Me.DateTime = objPost.Member("dateCreated").Value
        Me.More = ConvertHTMLEntities(objPost.Member("textMore").Value, False)
        Me.Excerpt = ConvertHTMLEntities(objPost.Member("excerpt").Value, False)
        Me.Keywords = ConvertHTMLEntities(objPost.Member("keywords").Value, False)
        Me.AllowComments = objPost.Member("allowComments").Value
        Me.AllowPings = objPost.Member("allowPings").Value
        Me.TextFilter = objPost.Member("textFilter").Value
        Me.TrackBack = objPost.Member("trackBack").Value
        Me.Categories = ConvertHTMLEntities(objPost.Member("categories").Value, False)
    End If
End Sub

Public Function SaveData(ByVal strFilePath As String) As Boolean
On Error Resume Next
Dim objPost As New xmlStruct
Dim sXML As String
    If Me.PostID <> "" Then
        objPost.Add "accountID", Me.AccountID
        objPost.Add "blogID", Me.BlogID
        objPost.Add "postID", Me.PostID
    Else
        objPost.Add "accountID", -1
        objPost.Add "blogID", ""
        objPost.Add "postID", ""
    End If
    objPost.Add "textTitle", ConvertHTMLEntities(Me.Title, True)
    objPost.Add "textBody", ConvertHTMLEntities(Me.Text, True)
    objPost.Add "dateCreated", Me.DateTime
    objPost.Add "textMore", ConvertHTMLEntities(Me.More, True)
    objPost.Add "excerpt", ConvertHTMLEntities(Me.Excerpt, True)
    objPost.Add "keywords", ConvertHTMLEntities(Me.Keywords, True)
    objPost.Add "allowComments", Me.AllowComments
    objPost.Add "allowPings", Me.AllowPings
    objPost.Add "textFilter", Me.TextFilter
    objPost.Add "trackBack", Me.TrackBack
    objPost.Add "categories", ConvertHTMLEntities(Me.Categories, True)
    '
    ' Setup the xml header
    '
    sXML = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & "?>" & vbCrLf
    sXML = sXML & "<postData>" & vbCrLf
    sXML = sXML & "<params>" & vbCrLf
    sXML = sXML & "<param>" & vbCrLf
    '
    ' Save Data
    '
    sXML = sXML & FormatParameter(objPost) & vbCrLf
    '
    ' Setup the xml footer
    '
    sXML = sXML & "</param>" & vbCrLf
    sXML = sXML & "</params>" & vbCrLf
    sXML = sXML & "</postData>"
    '
    ' Save File
    '
    SaveData = SaveBinaryFile(strFilePath, sXML)
End Function
