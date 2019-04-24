VERSION 5.00
Object = "{871CA184-A1A9-4E38-999E-E1AC439C0698}#1.0#0"; "DTPicker.ocx"
Begin VB.Form frmAdvanced 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Post Options"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Advanced.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTrackback 
      Height          =   1215
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2520
      Width           =   4275
   End
   Begin VB.ComboBox cboFilters 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1020
      Width           =   4275
   End
   Begin VB.ComboBox cboComments 
      Height          =   315
      ItemData        =   "Advanced.frx":000C
      Left            =   180
      List            =   "Advanced.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox cboPings 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3225
      TabIndex        =   11
      Top             =   3840
      Width           =   1230
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   3840
      Width           =   1230
   End
   Begin VB.CheckBox chkCurrentDate 
      Caption         =   "&Use Current Date and Time"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1500
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin DTPicker.DateTimePick dtpTime 
      Height          =   315
      Left            =   3120
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Format          =   9
      BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin DTPicker.DateTimePick dtpDate 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Format          =   0
      BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MinDate         =   25569
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "Send &TrackBack Pings to:"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   13
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "Text &Filters:"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   870
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "Allow &Comments:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "Allow &Pings:"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "&Date/Time:"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Top             =   1860
      Width           =   795
   End
End
Attribute VB_Name = "frmAdvanced"
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

Private Sub chkCurrentDate_Click()
On Error Resume Next
    dtpDate.Enabled = (chkCurrentDate.Value <> vbChecked)
    dtpTime.Enabled = (chkCurrentDate.Value <> vbChecked)
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next
    If Index = 0 Then 'OK
        If gPostID = "" And cboComments.ListIndex = cboComments.ListCount - 1 Then
            frmPost.PostData.AllowComments = -1
        Else
            frmPost.PostData.AllowComments = cboComments.ListIndex
        End If
        If gPostID = "" And cboPings.ListIndex = cboPings.ListCount - 1 Then
            frmPost.PostData.AllowPings = -1
        Else
            frmPost.PostData.AllowPings = cboPings.ListIndex
        End If
        If gPostID = "" And cboFilters.ListIndex = cboFilters.ListCount - 1 Then
            frmPost.PostData.TextFilter = ""
        Else
            frmPost.PostData.TextFilter = frmPost.TextFilters(cboFilters.ListIndex + 1)(1)
        End If
        If chkCurrentDate.Enabled And chkCurrentDate.Value = vbChecked Then
            frmPost.PostData.DateTime = CDate(0)
        Else
            frmPost.PostData.DateTime = CDate(Format(dtpDate.SelDate, "Short Date") + " " + Format(dtpTime.SelDate, "Long Time"))
        End If
        frmPost.PostData.TrackBack = txtTrackback
        frmPost.Changed = True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim t As Integer
    cboComments.AddItem GetLbl(lblNone)
    cboComments.AddItem GetLbl(lblOpen)
    cboComments.AddItem GetLbl(lblClosed)
    cboPings.AddItem GetLbl(lblNo)
    cboPings.AddItem GetLbl(lblYes)
    For t = 1 To frmPost.TextFilters.Count
        cboFilters.AddItem frmPost.TextFilters(t)(2)
    Next
    If gPostID = "" Then
        cboComments.AddItem "(" & LCase(GetLbl(lblDefault)) & ")"
        cboPings.AddItem "(" & LCase(GetLbl(lblDefault)) & ")"
        cboFilters.AddItem "(" & LCase(GetLbl(lblDefault)) & ")"
        If frmPost.PostData.AllowComments = -1 Then
            cboComments.ListIndex = cboComments.ListCount - 1
        Else
            cboComments.ListIndex = frmPost.PostData.AllowComments
        End If
        If frmPost.PostData.AllowPings = -1 Then
            cboPings.ListIndex = cboPings.ListCount - 1
        Else
            cboPings.ListIndex = frmPost.PostData.AllowPings
        End If
        If frmPost.PostData.DateTime <> CDate(0) Then
            chkCurrentDate.Value = vbUnchecked
            dtpDate.SelDate = frmPost.PostData.DateTime
            dtpTime.SelDate = frmPost.PostData.DateTime
        End If
        cboFilters.ListIndex = cboFilters.ListCount - 1
    Else
        chkCurrentDate.Value = vbUnchecked
        chkCurrentDate.Enabled = False
        cboFilters.ListIndex = frmPost.TextFilters(frmPost.PostData.TextFilter)(0)
        cboComments.ListIndex = frmPost.PostData.AllowComments
        cboPings.ListIndex = frmPost.PostData.AllowPings
        dtpDate.SelDate = frmPost.PostData.DateTime
        dtpTime.SelDate = frmPost.PostData.DateTime
    End If
    txtTrackback.Text = frmPost.PostData.TrackBack
    'Disable object according to Account settings
    cboComments.Enabled = gAccount.AllowComments
    cboPings.Enabled = gAccount.AllowPings
    cboFilters.Enabled = gAccount.TextFilters
    chkCurrentDate.Enabled = gAccount.PostDate
    If chkCurrentDate.Value = vbChecked Then
        dtpDate.Enabled = False
        dtpTime.Enabled = False
    Else
        dtpDate.Enabled = gAccount.PostDate
        dtpTime.Enabled = gAccount.PostDate
    End If
    txtTrackback.Enabled = gAccount.TrackBack
    lblFields(0).Enabled = gAccount.AllowComments
    lblFields(1).Enabled = gAccount.AllowPings
    lblFields(2).Enabled = gAccount.TextFilters
    lblFields(3).Enabled = gAccount.PostDate
    lblFields(4).Enabled = gAccount.TrackBack
    LocalizeForm
End Sub
Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblAdvPostOpt)
    lblFields(0).Caption = GetLbl(lblAllowComments) & ":"
    lblFields(1).Caption = GetLbl(lblAllowPings) & ":"
    lblFields(2).Caption = GetLbl(lblTextFilters) & ":"
    lblFields(3).Caption = GetLbl(lblDateTime) & ":"
    lblFields(4).Caption = GetLbl(lblSendTrackbackTo) & ":"
    chkCurrentDate.Caption = GetLbl(lblUseCurrentDT)
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub

