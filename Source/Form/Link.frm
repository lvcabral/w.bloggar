VERSION 5.00
Begin VB.Form frmLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Link"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Link.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtText 
      Height          =   315
      Left            =   1020
      MaxLength       =   255
      TabIndex        =   5
      Top             =   945
      Width           =   5580
   End
   Begin VB.ComboBox cboTarget 
      Height          =   315
      ItemData        =   "Link.frx":000C
      Left            =   4905
      List            =   "Link.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1365
      Width           =   1710
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   5460
      TabIndex        =   11
      Top             =   1845
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   10
      Top             =   1845
      Width           =   1155
   End
   Begin VB.CheckBox chkTarget 
      Appearance      =   0  'Flat
      Caption         =   "Open Link on"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3405
      TabIndex        =   8
      Top             =   1395
      Width           =   1755
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   1020
      MaxLength       =   255
      TabIndex        =   3
      Top             =   540
      Width           =   5580
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      Left            =   1020
      TabIndex        =   7
      Top             =   1365
      Width           =   2175
   End
   Begin VB.ComboBox cboURL 
      Height          =   315
      Left            =   1020
      TabIndex        =   1
      Text            =   "http://"
      Top             =   120
      Width           =   5580
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "T&ext:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   4
      Top             =   1005
      Width           =   390
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Cl&ass:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   6
      Top             =   1425
      Width           =   435
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Title:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   600
      Width           =   360
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&URL:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   345
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Public HtmlTag As String

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboClass, KeyAscii
End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboURL, KeyAscii
End Sub

Private Sub chkTarget_Click()
    cboTarget.Enabled = chkTarget.Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim i As Integer
    If Index = 0 Then
        If Trim(cboURL.Text) = "" Or Trim(cboURL.Text) = "http://" Then
            MsgBox GetMsg(msgEnterURL), vbInformation
            cboURL.SetFocus
            cboURL.SelStart = Len(cboURL.Text)
            Exit Sub
        End If
        HtmlTag = "<a href=""" & Trim(cboURL.Text) & """ "
        If Trim(cboClass.Text) <> "" Then
            HtmlTag = HtmlTag & "class=""" & cboClass.Text & """ "
        End If
        If Trim(txtTitle.Text) <> "" Then
            HtmlTag = HtmlTag & "title=""" & txtTitle.Text & """ "
        End If
        If chkTarget.Value = vbChecked Then
            Select Case cboTarget.ListIndex
            Case 0
                HtmlTag = HtmlTag & "target=""_blank"""
            Case 1
                HtmlTag = HtmlTag & "target=""_self"""
            Case 2
                HtmlTag = HtmlTag & "target=""_top"""
            End Select
        End If
        HtmlTag = Trim(HtmlTag) & ">" & txtText.Text
        Set objXMLReg = New XMLRegistry
        objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
        SaveCombo cboURL, "Links", "URL", 3
        If Trim(cboClass.Text) <> "" Then
            SaveCombo cboClass, "Classes", "Class"
        End If
        objXMLReg.SaveSetting App.Title, "Links", "UseTarget", Format(chkTarget.Value)
        objXMLReg.SaveSetting App.Title, "Links", "TheTarget", Format(cboTarget.ListIndex)
        Set objXMLReg = Nothing
    Else
        HtmlTag = ""
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
Dim i As Integer, strAux As String
On Error Resume Next
    LocalizeForm
    HtmlTag = ""
    'Load Links
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    cboURL.AddItem "http://"
    cboURL.AddItem "mailto:"
    For i = 0 To 30
        strAux = objXMLReg.GetSetting(App.Title, "Links", "URL" & Format(i, "00"), "**")
        If strAux = "**" Then
            Exit For
        Else
            cboURL.AddItem strAux
        End If
    Next
    'Load Classes
    For i = 0 To 30
        strAux = objXMLReg.GetSetting(App.Title, "Classes", "Class" & Format(i, "00"), "**")
        If strAux = "**" Then
            Exit For
        Else
            cboClass.AddItem strAux
        End If
    Next
    chkTarget.Value = Val(objXMLReg.GetSetting(App.Title, "Links", "UseTarget", "0"))
    cboTarget.ListIndex = Val(objXMLReg.GetSetting(App.Title, "Links", "TheTarget", "0"))
    cboTarget.Enabled = chkTarget.Value
    Set objXMLReg = Nothing
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblInsertLink)
    lblField(0).Caption = GetLbl(lblURL) & ":"
    lblField(1).Caption = GetLbl(lblTitle) & ":"
    lblField(2).Caption = GetLbl(lblClass) & ":"
    lblField(3).Caption = GetLbl(lblText) & ":"
    chkTarget.Caption = GetLbl(lblTarget)
    cboTarget.AddItem GetLbl(lblNewWindow)
    cboTarget.AddItem GetLbl(lblSameFrame)
    cboTarget.AddItem GetLbl(lblSameWindow)
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
