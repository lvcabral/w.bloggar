VERSION 5.00
Begin VB.Form frmCustom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Tags"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Custom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboCustom 
      Height          =   315
      ItemData        =   "Custom.frx":000C
      Left            =   1665
      List            =   "Custom.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   135
      Width           =   1530
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   6
      Top             =   1950
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   1950
      Width           =   1155
   End
   Begin VB.TextBox txtClose 
      Height          =   345
      Left            =   1650
      TabIndex        =   5
      Top             =   1455
      Width           =   2385
   End
   Begin VB.TextBox txtOpen 
      Height          =   345
      Left            =   1650
      TabIndex        =   3
      Top             =   990
      Width           =   2385
   End
   Begin VB.TextBox txtMenu 
      Height          =   345
      Left            =   1650
      TabIndex        =   1
      Top             =   540
      Width           =   2385
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Tag HotKey:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   165
      Width           =   900
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Tag &Close:"
      Height          =   195
      Index           =   3
      Left            =   105
      TabIndex        =   4
      Top             =   1530
      Width           =   765
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Tag Open:"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   2
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Menu Caption:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   615
      Width           =   1050
   End
End
Attribute VB_Name = "frmCustom"
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

Private Sub cboCustom_Click()
On Error Resume Next
    txtMenu.Text = gSettings.CustomTag(cboCustom.ListIndex + 1, 1)
    txtOpen.Text = gSettings.CustomTag(cboCustom.ListIndex + 1, 2)
    txtClose.Text = gSettings.CustomTag(cboCustom.ListIndex + 1, 3)
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next
    If Index = 0 Then 'OK
        If Trim(txtMenu.Text) = "" And Not (Trim(txtOpen.Text) = "" And Trim(txtClose.Text) = "") Then
            MsgBox GetMsg(msgEnterMenuCaption), vbInformation
            txtMenu.SetFocus
            Exit Sub
        ElseIf Trim(txtMenu.Text) <> "" And Trim(txtOpen.Text) = "" And Trim(txtClose.Text) = "" Then
            MsgBox GetMsg(msgEnterCustomTag), vbInformation
            txtOpen.SetFocus
            Exit Sub
        End If
        gSettings.CustomTag(cboCustom.ListIndex + 1, 1) = txtMenu.Text
        gSettings.CustomTag(cboCustom.ListIndex + 1, 2) = txtOpen.Text
        gSettings.CustomTag(cboCustom.ListIndex + 1, 3) = txtClose.Text
        SaveCustomTags
        If Trim(txtMenu.Text) <> "" Then
            frmPost.acbMain.Bands("bndPopCustom").Tools("miCustomF" & cboCustom.ListIndex + 1).Caption = gSettings.CustomTag(cboCustom.ListIndex + 1, 1)
        Else
            frmPost.acbMain.Bands("bndPopCustom").Tools("miCustomF" & cboCustom.ListIndex + 1).Caption = "� " & GetLbl(lblClickToEdit) & " �"
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
Dim oTool As ActiveBar2LibraryCtl.Tool
Dim oKey As ShortCut
On Error Resume Next
    For Each oTool In frmPost.acbMain.Bands("bndPopCustom").Tools
        If Left(oTool.Name, 9) = "miCustomF" Then
            Set oKey = oTool.ShortCuts(0)
            cboCustom.AddItem oKey.Value
        End If
    Next
    Set oKey = Nothing
    Set oTool = Nothing
    LocalizeForm
    
End Sub

Public Sub ShowForm(ByVal intCustomTag As Integer)
    cboCustom.ListIndex = intCustomTag - 1
    If Not Me.Visible Then
        Me.Show vbModal, frmPost
        Unload Me
    End If
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = Replace(frmPost.acbMain.Tools("miCustom").Caption, "&", "")
    lblField(0).Caption = GetLbl(lblTagHotkey) & ":"
    lblField(1).Caption = GetLbl(lblMenuCaption) & ":"
    lblField(2).Caption = GetLbl(lblTagOpen) & ":"
    lblField(3).Caption = GetLbl(lblTagClose) & ":"
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
