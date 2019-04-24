VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4035
   ClientLeft      =   2715
   ClientTop       =   2415
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2785.029
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2490
      Left            =   165
      Picture         =   "About.frx":000C
      ScaleHeight     =   2430
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   195
      Width           =   915
      Begin VB.Image imgIcon 
         Height          =   720
         Left            =   60
         Picture         =   "About.frx":111D
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4215
      TabIndex        =   0
      Top             =   3045
      Width           =   1365
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4215
      TabIndex        =   1
      Top             =   3495
      Width           =   1350
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Marcelo Lv Cabral"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   180
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   "https://lvcabral.com"
      Top             =   3570
      Width           =   3975
   End
   Begin VB.Label lblTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Developed and Maintained by"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label lblBrazil 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This Software was Made in Brazil"
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   3135
      Width           =   3975
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "This Freeware tool has portions of Jan G.P. Sijm code to XML-RPC access. To get more information click:"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   1365
      Width           =   3885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      Height          =   255
      Index           =   1
      Left            =   1455
      TabIndex        =   9
      Top             =   2310
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
      Height          =   270
      Index           =   0
      Left            =   1455
      TabIndex        =   7
      Top             =   2025
      Width           =   420
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "marcelo@lvcabral.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1995
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2310
      Width           =   1920
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "https://github.com/lvcabral/w.bloggar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1995
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2025
      Width           =   3300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1977.474
      Y2              =   1977.474
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1455
      TabIndex        =   5
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1455
      TabIndex        =   3
      Top             =   210
      UseMnemonic     =   0   'False
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1987.827
      Y2              =   1987.827
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   1455
      TabIndex        =   4
      Top             =   525
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdSysInfo_Click()
On Error Resume Next
Dim strSysInfo As String

    strSysInfo = GetOSVer & vbCrLf & _
                 "MS Internet Explorer " & GetIeVer & vbCrLf & _
                 "MS XML Parser " & GetXMLVer & vbCrLf & _
                 "Windows LCID: " & GetLocaleName(GetUserDefaultLCID) & " - " & GetUserDefaultLCID & vbCrLf & _
                 "w.bloggar LCID: " & GetLocaleName(gLCID) & " - " & gLCID
    MsgBox strSysInfo, vbInformation
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = GetLbl(lblSobre) & " " & App.Title
    lblVer.Caption = GetLbl(lblVersion) & " " & App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
    lblTitle.Caption = App.ProductName
    lblCopyright.Caption = App.LegalCopyright
    If gLCID <> 1033 Then
        lblDesc.Caption = GetLbl(lblDescription)
        lblBrazil.Caption = GetLbl(lblMadeInBrazil)
        lblTrans.Caption = GetLbl(lblTranslated)
        lblLink(2).Caption = GetLbl(lblTranslatorName)
        lblLink(2).ToolTipText = GetLbl(lblTranslatorURL)
        cmdOK.Caption = GetLbl(lblOK)
        cmdSysInfo.Caption = GetLbl(lblSysInfo)
    End If
    lblLink(0).MouseIcon = frmPost.imgHand.Picture
    lblLink(1).MouseIcon = frmPost.imgHand.Picture
    lblLink(2).MouseIcon = frmPost.imgHand.Picture
End Sub

Private Sub lblLink_Click(Index As Integer)
    Screen.MousePointer = vbArrowHourglass
    Select Case Index
    Case 0
        Call ShellExecute(Me.hwnd, "open", lblLink(Index).Caption, vbNullString, CurDir$, SW_SHOW)
    Case 1
        Call ShellExecute(Me.hwnd, "open", "mailto:" & lblLink(Index).Caption, vbNullString, CurDir$, SW_SHOW)
    Case 2
        Call ShellExecute(Me.hwnd, "open", lblLink(Index).ToolTipText, vbNullString, CurDir$, SW_SHOW)
    End Select
    Screen.MousePointer = vbNormal
End Sub
