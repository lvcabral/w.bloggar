VERSION 5.00
Begin VB.Form frmFont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Format Font"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Font.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSize 
      Height          =   315
      ItemData        =   "Font.frx":000C
      Left            =   3825
      List            =   "Font.frx":0025
      TabIndex        =   6
      Top             =   540
      Width           =   750
   End
   Begin VB.CommandButton cmdColor 
      Height          =   285
      Left            =   2325
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   570
      Width           =   285
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3420
      TabIndex        =   10
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2190
      TabIndex        =   9
      Top             =   1440
      Width           =   1155
   End
   Begin VB.TextBox txtColor 
      Height          =   330
      Left            =   1185
      MaxLength       =   7
      TabIndex        =   3
      Top             =   540
      Width           =   1455
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      Left            =   1185
      TabIndex        =   8
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      ItemData        =   "Font.frx":003E
      Left            =   1185
      List            =   "Font.frx":0066
      TabIndex        =   1
      Top             =   120
      Width           =   3390
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Size:"
      Height          =   195
      Index           =   3
      Left            =   3375
      TabIndex        =   5
      Top             =   600
      Width           =   345
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Cl&ass:"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "C&olor:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Width           =   435
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Font:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   390
   End
End
Attribute VB_Name = "frmFont"
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

Private Sub cboFont_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboFont, KeyAscii
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim i As Integer
    If Index = 0 Then
        HtmlTag = "<font "
        If Trim(cboFont.Text) <> "" Then
            HtmlTag = HtmlTag & "face=""" & cboFont.Text & """ "
        End If
        If Trim(txtColor.Text) <> "" Then
            HtmlTag = HtmlTag & "color=""" & txtColor.Text & """ "
        End If
        If Trim(cboSize.Text) <> "" Then
            HtmlTag = HtmlTag & "size=""" & cboSize.Text & """ "
        End If
        If Trim(cboClass.Text) <> "" Then
            HtmlTag = HtmlTag & "class=""" & cboClass.Text & """ "
        End If
        HtmlTag = Trim(HtmlTag) & ">"
        If Trim(cboClass.Text) <> "" Then
            SaveCombo cboClass, "Classes", "Class"
        End If
    Else
        HtmlTag = ""
    End If
    Me.Hide
End Sub

Private Sub cmdColor_Click()
Dim strBGR As String, strRGB As String
    objColor.hWndParent = Me.hwnd
    If objColor.ShowColor Then
        SaveColors
        strBGR = Right("0000" & Hex(objColor.Color), 6)
        strRGB = Right(strBGR, 2) & Mid(strBGR, 3, 2) & Left(strBGR, 2)
        txtColor.Text = "#" & strRGB
    End If
    txtColor.SetFocus
End Sub

Private Sub Form_Activate()
    cboFont.SelStart = Len(cboFont.Text)
End Sub

Private Sub Form_Load()
Dim i As Integer, strAux As String
On Error Resume Next
    LocalizeForm
    cmdColor.Picture = frmPost.acbMain.Tools("miColor").GetPicture(ddITNormal)
    HtmlTag = ""
    'Load Classes
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    For i = 0 To 30
        strAux = objXMLReg.GetSetting(App.Title, "Classes", "Class" & Format(i, "00"), "**")
        If strAux = "**" Then
            Exit For
        Else
            cboClass.AddItem strAux
        End If
    Next
    Set objXMLReg = Nothing
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblFormatFont)
    lblField(0).Caption = GetLbl(lblFontFace) & ":"
    lblField(1).Caption = GetLbl(lblColor) & ":"
    lblField(2).Caption = GetLbl(lblClass) & ":"
    lblField(3).Caption = GetLbl(lblSize) & ":"
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub

