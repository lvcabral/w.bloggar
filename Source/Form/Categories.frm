VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Begin VB.Form frmCategories 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Categories"
   ClientHeight    =   4890
   ClientLeft      =   360
   ClientTop       =   705
   ClientWidth     =   6210
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
   ScaleHeight     =   4890
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   4890
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "v"
      Top             =   0
      Width           =   6210
      _cx             =   10954
      _cy             =   8625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   2
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"Categories.frx":000C
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   3375
         ScaleHeight     =   600
         ScaleWidth      =   2745
         TabIndex        =   2
         Top             =   4200
         Width           =   2745
         Begin VB.CommandButton cmdButton 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   465
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdButton 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   465
            Index           =   1
            Left            =   1440
            TabIndex        =   3
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.ListBox lstCategs 
         Columns         =   3
         Height          =   3885
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   90
         Width           =   6030
      End
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
    LoadFormSettings Me, frmPost.Left + 300, frmPost.Top + 300, 6300, 5300
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

Private Sub Form_Unload(Cancel As Integer)
    SaveFormSettings Me
End Sub
