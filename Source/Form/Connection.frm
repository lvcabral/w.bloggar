VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Properties"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Connection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCMS 
      Caption         =   "Content Management System"
      Height          =   1185
      Left            =   120
      TabIndex        =   23
      Top             =   45
      Width           =   3975
      Begin VB.TextBox txtAlias 
         Height          =   315
         Left            =   1410
         MaxLength       =   25
         TabIndex        =   3
         Top             =   735
         Width           =   2385
      End
      Begin VB.ComboBox cboCMS 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   255
         Width           =   2085
      End
      Begin VB.Image imgCMS 
         Height          =   240
         Left            =   3555
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "A&ccount Alias:"
         Height          =   195
         Index           =   8
         Left            =   225
         TabIndex        =   2
         Top             =   795
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "&Blog Tool:"
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   0
         Top             =   315
         Width           =   705
      End
   End
   Begin VB.CheckBox chkProxy 
      Appearance      =   0  'Flat
      Caption         =   "Proxy Server"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   270
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame fraProxy 
      Caption         =   "  "
      Enabled         =   0   'False
      Height          =   1320
      Left            =   120
      TabIndex        =   20
      Top             =   3495
      Width           =   3975
      Begin VB.TextBox txtProxyPort 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1395
         MaxLength       =   5
         TabIndex        =   17
         Top             =   780
         Width           =   720
      End
      Begin VB.TextBox txtProxyServer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         MaxLength       =   255
         TabIndex        =   15
         Top             =   375
         Width           =   2400
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Por&t:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   810
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "&Address:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   405
         Width           =   645
      End
   End
   Begin VB.Frame fraBlogger 
      Caption         =   "API Server"
      Height          =   2025
      Left            =   120
      TabIndex        =   21
      Top             =   1365
      Width           =   3975
      Begin VB.TextBox txtTimeout 
         Alignment       =   1  'Right Justify
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1395
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1515
         Width           =   720
      End
      Begin VB.CheckBox chkSecure 
         Appearance      =   0  'Flat
         Caption         =   "HTTPS"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3060
         TabIndex        =   10
         Top             =   1140
         Width           =   825
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1395
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1110
         Width           =   720
      End
      Begin VB.TextBox txtHost 
         Height          =   315
         Left            =   1395
         MaxLength       =   255
         TabIndex        =   5
         Top             =   285
         Width           =   2400
      End
      Begin VB.TextBox txtPage 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1395
         MaxLength       =   255
         TabIndex        =   7
         Top             =   705
         Width           =   2400
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Index           =   6
         Left            =   2190
         TabIndex        =   22
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Ti&meout:"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   11
         Top             =   1545
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "Por&t:"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   8
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "&Host:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   330
         Width           =   390
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "P&age:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Top             =   735
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1710
      TabIndex        =   18
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2940
      TabIndex        =   19
      Top             =   4920
      Width           =   1155
   End
End
Attribute VB_Name = "frmConnection"
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

Private Sub cboCMS_Click()
Dim strSec As String, strIni As String, bolSvc As Boolean, strIco As String
    If cboCMS.ListIndex > 0 Then
        strSec = "CMS-" & Format(cboCMS.ListIndex, "00")
        strIni = App.Path & "\CMS\CMS.ini"
        strIco = App.Path & "\CMS\" & ReadINI(strSec, "Icon", strIni)
        If FileExists(strIco) Then
            imgCMS.Picture = LoadPicture(strIco)
        Else
            imgCMS.Picture = LoadPicture()
        End If
        bolSvc = CBool(Val(ReadINI(strSec, "Service", strIni)))
        txtHost.Text = ReadINI(strSec, "Host", strIni)
        txtPage.Text = ReadINI(strSec, "Page", strIni)
        txtPort.Text = ReadINI(strSec, "Port", strIni)
        chkSecure.Value = Val(ReadINI(strSec, "Https", strIni))
    Else
        imgCMS.Picture = LoadPicture()
        bolSvc = False
    End If
    txtHost.Enabled = Not bolSvc
    txtPage.Enabled = Not bolSvc
    txtPort.Enabled = Not bolSvc
    chkSecure.Enabled = Not bolSvc
End Sub

Private Sub chkProxy_Click()
    If chkProxy.Value Then
        fraProxy.Enabled = True
        lblField(3).Enabled = True
        lblField(4).Enabled = True
        txtProxyServer.Enabled = True
        txtProxyPort.Enabled = True
        txtProxyServer.BackColor = vbWindowBackground
        txtProxyPort.BackColor = vbWindowBackground
    Else
        fraProxy.Enabled = False
        lblField(3).Enabled = False
        lblField(4).Enabled = False
        txtProxyServer.Enabled = False
        txtProxyPort.Enabled = False
        txtProxyServer.BackColor = vbButtonFace
        txtProxyPort.BackColor = vbButtonFace
    End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
    If Index = 0 Then
        If Trim(txtAlias.Text) = "" Then
            MsgBox "Enter an Alias for this Account!", vbInformation
            txtAlias.SetFocus
            Exit Sub
        ElseIf Trim(txtHost.Text) = "" Then
            MsgBox GetMsg(msgEnterHost), vbInformation
            txtHost.SetFocus
            Exit Sub
        ElseIf Trim(txtPage.Text) = "" Then
            MsgBox GetMsg(msgEnterPage), vbInformation
            txtPage.SetFocus
            Exit Sub
        ElseIf Val(txtPort.Text) <= 0 Then
            MsgBox GetMsg(msgEnterPort), vbInformation
            txtPort.SetFocus
            Exit Sub
        ElseIf Val(txtTimeout.Text) <= 0 Then
            MsgBox GetMsg(msgEnterTimeout), vbInformation
            txtTimeout.SetFocus
            Exit Sub
        End If
        If chkProxy.Value Then
            If Trim(txtProxyServer.Text) = "" Then
                MsgBox GetMsg(msgEnterProxy), vbInformation
                txtProxyServer.SetFocus
                Exit Sub
            ElseIf Val(txtProxyPort.Text) <= 0 Then
                MsgBox GetMsg(msgEnterPort), vbInformation
                txtProxyPort.SetFocus
                Exit Sub
            End If
        End If
        gAccount.CMS = cboCMS.ListIndex
        gAccount.Service = Not txtHost.Enabled
        gAccount.Alias = txtAlias.Text
        gAccount.Host = txtHost.Text
        gAccount.Page = txtPage.Text
        gAccount.Port = Val(txtPort.Text)
        gAccount.Secure = chkSecure.Value
        gAccount.Timeout = Val(txtTimeout.Text)
        gAccount.UseProxy = chkProxy.Value
        If gAccount.UseProxy Then
            gAccount.ProxyServer = txtProxyServer.Text
            gAccount.ProxyPort = Val(txtProxyPort.Text)
        End If
        If gAccount.User <> "" Then SaveAccount
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    LocalizeForm
    'Populate Fields
    LoadCMS
    If gAccount.Alias <> "" Then 'Existing Account
        txtAlias.Text = gAccount.Alias
        txtHost.Text = gAccount.Host
        txtPage.Text = gAccount.Page
        txtPort.Text = gAccount.Port
        chkSecure.Value = Abs(gAccount.Secure)
    Else 'New Account
        cboCMS.ListIndex = 1
    End If
    txtTimeout.Text = gAccount.Timeout
    chkProxy.Value = Abs(gAccount.UseProxy)
    txtProxyServer.Text = gAccount.ProxyServer
    txtProxyPort.Text = Format(gAccount.ProxyPort, "#")
End Sub

Private Sub LocalizeForm()
On Error Resume Next
    'Alterar -> Me.Caption = GetLbl(lblConnection)
    lblField(0).Caption = GetLbl(lblHost) & ":"
    lblField(1).Caption = GetLbl(lblPage) & ":"
    lblField(2).Caption = GetLbl(lblPort) & ":"
    'Eliminar -> lblAlert.Caption = GetLbl(lblDontChange)
    chkProxy.Caption = GetLbl(lblProxy)
    lblField(3).Caption = GetLbl(lblAddress) & ":"
    lblField(4).Caption = GetLbl(lblPort) & ":"
    lblField(5).Caption = GetLbl(lblTimeout) & ":"
    lblField(6).Caption = GetLbl(lblSeconds)
    'Eliminar -> cmdRestore.Caption = GetLbl(lblRestoreBlogger)
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
End Sub

Private Sub LoadCMS()
On Error Resume Next
Dim i As Integer, l As Integer
Dim strIni As String
strIni = App.Path & "\CMS\CMS.ini"
    l = ReadINI("CMS", "Count", strIni)
    cboCMS.Clear
    cboCMS.AddItem "� Custom Tool �"
    For i = 1 To l
        cboCMS.AddItem ReadINI("CMS-" & Format(i, "00"), "Name", strIni)
    Next
    cboCMS.ListIndex = gAccount.CMS
End Sub
