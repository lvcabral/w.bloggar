VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpell"
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
Public Event BeforeChange(bCancel As Boolean)
Public Event BeforeAdd(bCancel As Boolean)
Public Event BeforeSave(bCancel As Boolean)
Public Event BeforeIgnore(bCancel As Boolean) ' not called
Public Event CompleteSpellCheck()
' Properties
Private m_TextBox As Object
Private m_Dictionary As String
Private m_ShowOnTop As Boolean
Private WithEvents frmSpellCheck As frmCheck
Attribute frmSpellCheck.VB_VarHelpID = -1
Private bLoadedWords As Boolean

' returns the handle of the Spell Checker window
Public Property Let WindowhWnd(ByVal hwnd As Long)
    hwnd = frmCheck.hwnd
End Property
' Returns a reference to the actual form
' has to be declared as object - vb does not allow the use of form in this case
Public Property Get SpellCheckerForm() As Object
    Set SpellCheckerForm = frmSpellCheck
End Property
' Dictionary path
Public Property Let Dictionary(ByVal vData As String)
    If m_Dictionary <> vData And m_Dictionary <> "" Then
        SaveDictIfNeeded
        bLoadedWords = False
    End If
    m_Dictionary = vData
End Property
Public Property Get Dictionary() As String
    Dictionary = m_Dictionary
End Property
' Allow adding to dictionary?
Public Property Let AllowAdd(ByVal vData As Boolean)
    frmSpellCheck.cmdAdd.Enabled = vData
End Property
Public Property Get AllowAdd() As Boolean
    AllowAdd = frmSpellCheck.cmdAdd.Enabled
End Property
' Set the caption of the window
Public Property Let SpellCaption(ByVal vData As String)
    frmSpellCheck.Caption = vData
End Property
Public Property Get SpellCaption() As String
    SpellCaption = frmSpellCheck.Caption
End Property
' set the TextBox/RichTextBox to check
Public Property Set TextBox(ByVal vData As Object)
    ' useless
    On Error Resume Next
    IIf vData.Text = "", True, False
    If Err Then
        Err.Raise vbObjectError + 5000, "Spell Checker", "Invalid TextBox object."
    Else
        Set m_TextBox = vData
    End If
End Property
Public Property Get TextBox() As Object
    Set TextBox = m_TextBox
End Property
' Events
Private Sub frmSpellCheck_BeforeAdd(bCancel As Boolean)
    RaiseEvent BeforeAdd(bCancel)
End Sub
Private Sub frmSpellCheck_BeforeChange(bCancel As Boolean)
    RaiseEvent BeforeChange(bCancel)
End Sub

Private Sub frmSpellCheck_BeforeSave()
    SaveDictIfNeeded
End Sub

Private Sub frmSpellCheck_CompleteSpellCheck()
    RaiseEvent CompleteSpellCheck
End Sub
' Procedures
Public Sub SpellCheck(ByVal bFromStart As Boolean)
    If Not bLoadedWords Then
        'LoadWords Dictionary
        frmProgress.RunJob "LoadWords", Dictionary
        bLoadedWords = True
    End If
    If bFromStart Then
        TextBox.SelStart = 0
    End If
    frmSpellCheck.SpellCheck TextBox
End Sub
Public Sub AddCustomItem(ByVal sItem As String)
    If frmSpellCheck.AddItem(LCase$(sItem), False) = False Then
        Err.Raise vbObjectError + 5001, "Spell Checker", "That item already exists in the dictionary"
    End If
End Sub
Public Sub LoadDictionary()
    ' LoadWords Dictionary
    frmProgress.RunJob "LoadWords", Dictionary
End Sub
Public Sub SaveDictionary()
    SaveWords Dictionary
End Sub
' Class
Private Sub Class_Initialize()
    Set frmSpellCheck = New frmCheck
    ' set the default dictionary
    Dictionary = App.Path & "\dictionary.txt"
    ' initialise the WordsToAdd array
    ReDim WordsToAdd(0)
    ' Set default properties
    m_ShowOnTop = True
End Sub

Private Sub Class_Terminate()
    Unload frmSpellCheck
End Sub

Private Sub SaveDictIfNeeded()
    Dim bCancel As Boolean
    If Forms.Count = 0 Then Exit Sub
    If frmSpellCheck.bDictChanged Then
        RaiseEvent BeforeSave(bCancel)
        If Not bCancel Then SaveWords Dictionary
    End If
End Sub

Private Sub SaveWords(strDictionaryFile As String)
    If bFileNotFound Then Exit Sub
    Dim iFreeFile As Integer
    Dim i As Long
    On Error Resume Next
    iFreeFile = FreeFile
    Open strDictionaryFile For Append As iFreeFile
    For i = 1 To UBound(WordsToAdd)
        Print #iFreeFile, WordsToAdd(i)
    Next
    Close iFreeFile
End Sub

