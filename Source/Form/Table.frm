VERSION 5.00
Begin VB.Form frmTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Table"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Table.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRows 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   1
      Top             =   105
      Width           =   615
   End
   Begin VB.TextBox txtCols 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   3
      Top             =   105
      Width           =   615
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1485
      MaxLength       =   5
      TabIndex        =   5
      Top             =   525
      Width           =   615
   End
   Begin VB.TextBox txtPadding 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   9
      Top             =   945
      Width           =   615
   End
   Begin VB.TextBox txtSpacing 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   11
      Top             =   945
      Width           =   615
   End
   Begin VB.TextBox txtBorder 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   7
      Top             =   525
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Top             =   1425
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3435
      TabIndex        =   13
      Top             =   1425
      Width           =   1155
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Rows:"
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   450
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Columns:"
      Height          =   195
      Index           =   5
      Left            =   2640
      TabIndex        =   2
      Top             =   165
      Width           =   660
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   585
      Width           =   480
   End
   Begin VB.Label lblField 
      Caption         =   "Cell Padding:"
      Height          =   405
      Index           =   2
      Left            =   150
      TabIndex        =   8
      Top             =   975
      Width           =   1320
   End
   Begin VB.Label lblField 
      Caption         =   "Cell Spacing:"
      Height          =   405
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   975
      Width           =   1290
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Border:"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   6
      Top             =   585
      Width           =   540
   End
End
Attribute VB_Name = "frmTable"
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

Private Sub cmdButton_Click(Index As Integer)
Dim r As Integer, c As Integer
Dim strSep As String, strMrg As String
    If Index = 0 Then
        If Val(txtRows.Text) = 0 Or Val(txtCols.Text) = 0 Then
            MsgBox GetMsg(msgEnterRowCol), vbInformation
            txtRows.SetFocus
            Exit Sub
        End If
        If gPostID = "main" Or gPostID = "archiveIndex" Or Not gBlog.PreviewAutoBR Then
            strSep = vbCrLf
            strMrg = "     "
        Else
            strSep = ""
            strMrg = ""
        End If
        HtmlTag = "<table"
        If Trim(txtWidth.Text) <> "" Then
            HtmlTag = HtmlTag & " width=""" & txtWidth.Text & """"
        End If
        If Trim(txtBorder.Text) <> "" Then
            HtmlTag = HtmlTag & " border=""" & txtBorder.Text & """"
        End If
        If Trim(txtSpacing.Text) <> "" Then
            HtmlTag = HtmlTag & " cellspacing=""" & txtSpacing.Text & """"
        End If
        If Trim(txtPadding.Text) <> "" Then
            HtmlTag = HtmlTag & " cellpadding=""" & txtPadding.Text & """"
        End If
        HtmlTag = HtmlTag & "> " & strSep
        For r = 1 To Val(txtRows.Text)
            HtmlTag = HtmlTag & "<tr>" & strSep
            For c = 1 To Val(txtCols.Text)
                HtmlTag = HtmlTag & strMrg & "<td>" & GetLbl(lblTextHere) & "</td>" & strSep
            Next
            HtmlTag = HtmlTag & "</tr> " & strSep
        Next
        HtmlTag = HtmlTag & "</table>" & strSep
        Set objXMLReg = New XMLRegistry
        objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
        objXMLReg.SaveSetting App.Title, "Table", "Width", txtWidth.Text
        objXMLReg.SaveSetting App.Title, "Table", "Border", txtBorder.Text
        objXMLReg.SaveSetting App.Title, "Table", "Spacing", txtSpacing.Text
        objXMLReg.SaveSetting App.Title, "Table", "Padding", txtPadding.Text
        objXMLReg.SaveSetting App.Title, "Table", "Rows", txtRows.Text
        objXMLReg.SaveSetting App.Title, "Table", "Cols", txtCols.Text
        Set objXMLReg = Nothing
    Else
        HtmlTag = ""
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
    LocalizeForm
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    txtWidth.Text = objXMLReg.GetSetting(App.Title, "Table", "Width", "")
    txtBorder.Text = objXMLReg.GetSetting(App.Title, "Table", "Border", "")
    txtSpacing.Text = objXMLReg.GetSetting(App.Title, "Table", "Spacing", "")
    txtPadding.Text = objXMLReg.GetSetting(App.Title, "Table", "Padding", "")
    txtRows.Text = objXMLReg.GetSetting(App.Title, "Table", "Rows", "1")
    txtCols.Text = objXMLReg.GetSetting(App.Title, "Table", "Cols", "1")
    Set objXMLReg = Nothing
End Sub

Private Sub txtBorder_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCols_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPadding_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRows_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpacing_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        If KeyAscii <> Asc("%") Then KeyAscii = 0
    End If
End Sub


Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblInsertTable)
    lblField(0).Caption = GetLbl(lblWidth) & ":"
    lblField(1).Caption = GetLbl(lblBorder) & ":"
    lblField(2).Caption = GetLbl(lblCellPadding) & ":"
    lblField(3).Caption = GetLbl(lblCellSpacing) & ":"
    lblField(4).Caption = GetLbl(lblRows) & ":"
    lblField(5).Caption = GetLbl(lblColumns) & ":"
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub

