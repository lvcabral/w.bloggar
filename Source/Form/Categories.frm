VERSION 5.00
Begin VB.Form frmCategories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categories"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Categories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Top             =   2580
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   2580
      Width           =   1155
   End
   Begin VB.ListBox lstCategs 
      Height          =   2310
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   105
      Width           =   3690
   End
End
Attribute VB_Name = "frmCategories"
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

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim b As Integer, strCategs As String
    If Index = 0 Then
        strCategs = ""
        For b = 0 To lstCategs.ListCount - 1
            If lstCategs.Selected(b) Then
                If gAccount.GetCategMethod = API_MT Or gAccount.GetCategMethod = API_B2 Then
                    strCategs = strCategs & Format(lstCategs.ItemData(b), CATEG_ID_MASK) & vbTab
                Else
                    strCategs = strCategs & lstCategs.List(b) & vbTab
                End If
            End If
        Next
        If frmPost.cboPostCat.ListIndex > 0 Then
            If gAccount.GetCategMethod = API_MT Or gAccount.GetCategMethod = API_B2 Then
                strCategs = Format(frmPost.cboPostCat.ItemData(frmPost.cboPostCat.ListIndex), CATEG_ID_MASK) & vbTab & strCategs
            Else
                strCategs = frmPost.cboPostCat.List(frmPost.cboPostCat.ListIndex) & vbTab & strCategs
            End If
        End If
        frmPost.PostData.Categories = strCategs
        frmPost.Changed = True
        Unload Me
    Else
        Unload Me
    End If
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim c As Integer
    LocalizeForm
    For c = 1 To frmPost.cboPostCat.ListCount - 2
        If c <> frmPost.cboPostCat.ListIndex Then
            lstCategs.AddItem frmPost.cboPostCat.List(c)
            lstCategs.ItemData(lstCategs.NewIndex) = frmPost.cboPostCat.ItemData(c)
            If gAccount.GetCategMethod = API_MT Or gAccount.GetCategMethod = API_B2 Then
                If InStr(frmPost.PostData.Categories, Format(frmPost.cboPostCat.ItemData(c), CATEG_ID_MASK)) > 0 Then
                    lstCategs.Selected(lstCategs.NewIndex) = True
                End If
            Else
                If InStr(frmPost.PostData.Categories, frmPost.cboPostCat.List(c)) > 0 Then
                    lstCategs.Selected(lstCategs.NewIndex) = True
                End If
            End If
        End If
    Next
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblCategories)
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
