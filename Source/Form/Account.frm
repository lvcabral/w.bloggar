VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{ED442B9F-ADE2-11D4-B868-00606E3BC2C9}#1.0#0"; "ActiveCbo.ocx"
Begin VB.Form frmAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Properties"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Account.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne tabIndex 
      Height          =   3060
      Left            =   120
      TabIndex        =   5
      Top             =   1620
      Width           =   4245
      _cx             =   7488
      _cy             =   5397
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
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "API Server|Proxy Server|Custom"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin SizerOneLibCtl.ElasticOne ElasticOne3 
         Height          =   2685
         Left            =   45
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   330
         Width           =   4155
         _cx             =   7329
         _cy             =   4736
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
         Align           =   0
         AutoSizeChildren=   0
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
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         Begin VB.CheckBox chkUTF8 
            Appearance      =   0  'Flat
            Caption         =   "UTF-8"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3105
            TabIndex        =   15
            Top             =   1650
            Width           =   780
         End
         Begin VB.CommandButton cmdReloadBlogs 
            Caption         =   "&Reload Blogs List"
            Height          =   345
            Left            =   1335
            TabIndex        =   16
            Top             =   2025
            Width           =   2595
         End
         Begin VB.TextBox txtTimeout 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1335
            MaxLength       =   5
            TabIndex        =   14
            Top             =   1620
            Width           =   720
         End
         Begin VB.CheckBox chkSecure 
            Appearance      =   0  'Flat
            Caption         =   "HTTPS"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3105
            TabIndex        =   12
            Top             =   1245
            Width           =   825
         End
         Begin VB.TextBox txtPort 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1335
            MaxLength       =   5
            TabIndex        =   11
            Top             =   1215
            Width           =   720
         End
         Begin VB.TextBox txtHost 
            Height          =   315
            Left            =   1335
            MaxLength       =   255
            TabIndex        =   7
            Top             =   390
            Width           =   2595
         End
         Begin VB.TextBox txtPage 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1335
            MaxLength       =   255
            TabIndex        =   9
            Top             =   810
            Width           =   2595
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "seconds"
            Height          =   195
            Index           =   6
            Left            =   2130
            TabIndex        =   51
            Top             =   1650
            Width           =   585
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Ti&meout:"
            Height          =   195
            Index           =   5
            Left            =   165
            TabIndex        =   13
            Top             =   1650
            Width           =   630
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Por&t:"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   10
            Top             =   1245
            Width           =   360
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Host:"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   6
            Top             =   435
            Width           =   390
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "P&ath:"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   8
            Top             =   840
            Width           =   390
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   2685
         Left            =   5190
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   330
         Width           =   4155
         _cx             =   7329
         _cy             =   4736
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
         Align           =   0
         AutoSizeChildren=   0
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
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         Begin VB.TextBox txtMoreTag 
            Height          =   315
            Index           =   1
            Left            =   2865
            TabIndex        =   42
            Text            =   "</more_text>"
            Top             =   2205
            Width           =   1110
         End
         Begin VB.TextBox txtMoreTag 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   41
            Text            =   "<more_text>"
            Top             =   2205
            Width           =   1110
         End
         Begin VB.ComboBox cboPostAPI 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   165
            Width           =   2310
         End
         Begin VB.ComboBox cboCategAPI 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   570
            Width           =   2310
         End
         Begin VB.ComboBox cboTemplAPI 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   975
            Width           =   2310
         End
         Begin VB.TextBox txtTitleTag 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   35
            Text            =   "<title>"
            Top             =   1395
            Width           =   1110
         End
         Begin VB.TextBox txtTitleTag 
            Height          =   315
            Index           =   1
            Left            =   2865
            TabIndex        =   36
            Text            =   "</title>"
            Top             =   1395
            Width           =   1110
         End
         Begin VB.TextBox txtCategTag 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   38
            Text            =   "<category>"
            Top             =   1800
            Width           =   1110
         End
         Begin VB.TextBox txtCategTag 
            Height          =   315
            Index           =   1
            Left            =   2865
            TabIndex        =   39
            Text            =   "</category>"
            Top             =   1800
            Width           =   1110
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "More Text Tags:"
            Height          =   195
            Index           =   16
            Left            =   165
            TabIndex        =   40
            Top             =   2265
            Width           =   1185
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Posts:"
            Height          =   195
            Index           =   9
            Left            =   165
            TabIndex        =   28
            Top             =   225
            Width           =   450
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Categories:"
            Height          =   195
            Index           =   10
            Left            =   165
            TabIndex        =   30
            Top             =   630
            Width           =   840
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Templates:"
            Height          =   195
            Index           =   11
            Left            =   165
            TabIndex        =   32
            Top             =   1035
            Width           =   795
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Title Tags:"
            Height          =   195
            Index           =   12
            Left            =   165
            TabIndex        =   34
            Top             =   1485
            Width           =   750
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Category Tags:"
            Height          =   195
            Index           =   13
            Left            =   165
            TabIndex        =   37
            Top             =   1860
            Width           =   1125
         End
      End
      Begin VB.PictureBox picProxy 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   2685
         Left            =   4890
         ScaleHeight     =   2685
         ScaleWidth      =   4155
         TabIndex        =   46
         Top             =   330
         Width           =   4155
         Begin VB.TextBox txtPassword 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1470
            MaxLength       =   255
            PasswordChar    =   "*"
            TabIndex        =   27
            Top             =   2235
            Width           =   1485
         End
         Begin VB.TextBox txtUser 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            MaxLength       =   255
            TabIndex        =   25
            Top             =   1830
            Width           =   1485
         End
         Begin VB.OptionButton optProxy 
            Appearance      =   0  'Flat
            Caption         =   "Use the following Proxy Settings:"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   675
            Width           =   3975
         End
         Begin VB.OptionButton optProxy 
            Appearance      =   0  'Flat
            Caption         =   "Use Internet Explorer Proxy Settings"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   375
            Width           =   3975
         End
         Begin VB.OptionButton optProxy 
            Appearance      =   0  'Flat
            Caption         =   "Don't use Proxy"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   90
            Value           =   -1  'True
            Width           =   3975
         End
         Begin VB.TextBox txtProxyServer 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            MaxLength       =   255
            TabIndex        =   21
            Top             =   1020
            Width           =   2400
         End
         Begin VB.TextBox txtProxyPort 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1470
            MaxLength       =   5
            TabIndex        =   23
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Password:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   15
            Left            =   195
            TabIndex        =   26
            Top             =   2265
            Width           =   750
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&User:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   14
            Left            =   195
            TabIndex        =   24
            Top             =   1860
            Width           =   390
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Address:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   20
            Top             =   1050
            Width           =   645
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Por&t:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   22
            Top             =   1455
            Width           =   360
         End
      End
   End
   Begin VB.Frame fraCMS 
      Caption         =   "Content Management System"
      Height          =   1485
      Left            =   120
      TabIndex        =   45
      Top             =   45
      Width           =   4245
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   1215
         Left            =   75
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   195
         Width           =   4080
         _cx             =   7197
         _cy             =   2143
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
         Align           =   0
         AutoSizeChildren=   0
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
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         Begin VB.ComboBox cboPing 
            Height          =   315
            ItemData        =   "Account.frx":000C
            Left            =   2580
            List            =   "Account.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   900
            Width           =   1395
         End
         Begin rdActiveCombo.ActiveCombo cboCMS 
            Height          =   330
            Left            =   1620
            TabIndex        =   1
            Top             =   75
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   582
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IconColorDepth  =   2
            ShowIcons       =   -1  'True
            Style           =   2
         End
         Begin VB.TextBox txtAlias 
            Height          =   315
            Left            =   1605
            MaxLength       =   25
            TabIndex        =   3
            Top             =   495
            Width           =   2385
         End
         Begin VB.CheckBox chkWeblogs 
            Appearance      =   0  'Flat
            Caption         =   "Ping:"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1605
            TabIndex        =   4
            Top             =   945
            Width           =   2490
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Blog Tool:"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   0
            Top             =   135
            Width           =   705
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "A&ccount Alias:"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   2
            Top             =   555
            Width           =   1020
         End
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1995
      TabIndex        =   43
      Top             =   4770
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3210
      TabIndex        =   44
      Top             =   4770
      Width           =   1155
   End
   Begin VB.Image imgCustom 
      Height          =   240
      Left            =   240
      Picture         =   "Account.frx":003E
      Top             =   4800
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmAccount"
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
Private bolService As Boolean

Private Sub cboCMS_Click()
Dim strSec As String, strIni As String
    If cboCMS.ListIndex < cboCMS.ListCount - 1 Then
        strSec = "CMS-" & Format(cboCMS.ItemData(cboCMS.ListIndex), "00")
        strIni = App.Path & "\CMS\CMS.ini"
        bolService = CBool(Val(ReadINI(strSec, "Service", strIni)))
        txtHost.Text = ReadINI(strSec, "Host", strIni)
        txtPage.Text = ReadINI(strSec, "Page", strIni)
        txtPort.Text = ReadINI(strSec, "Port", strIni)
        chkSecure.Value = Val(ReadINI(strSec, "Https", strIni))
        tabIndex.TabEnabled(2) = False
    Else
        tabIndex.TabEnabled(2) = True
        bolService = False
    End If
    tabIndex.CurrTab = 0
    txtHost.Enabled = Not bolService
    txtPage.Enabled = Not bolService
    txtPort.Enabled = Not bolService
    chkSecure.Enabled = Not bolService
End Sub

Private Sub chkSecure_Click()
    If chkSecure.Value = vbChecked Then
        txtPort.Text = "443"
    Else
        txtPort.Text = "80"
    End If
End Sub

Private Sub chkWeblogs_Click()
    cboPing.Enabled = chkWeblogs.Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
    If Index = 0 Then
        If Trim(txtAlias.Text) = "" Then
            MsgBox GetMsg(msgEnterAlias), vbInformation
            tabIndex.CurrTab = 0
            txtAlias.SetFocus
            Exit Sub
        ElseIf Trim(txtHost.Text) = "" Then
            MsgBox GetMsg(msgEnterHost), vbInformation
            tabIndex.CurrTab = 0
            txtHost.SetFocus
            Exit Sub
        ElseIf Trim(txtPage.Text) = "" Then
            MsgBox GetMsg(msgEnterPage), vbInformation
            tabIndex.CurrTab = 0
            txtPage.SetFocus
            Exit Sub
        ElseIf Val(txtPort.Text) <= 0 Then
            MsgBox GetMsg(msgEnterPort), vbInformation
            tabIndex.CurrTab = 0
            txtPort.SetFocus
            Exit Sub
        ElseIf Val(txtTimeout.Text) <= 0 Then
            MsgBox GetMsg(msgEnterTimeout), vbInformation
            tabIndex.CurrTab = 0
            txtTimeout.SetFocus
            Exit Sub
        End If
        If optProxy(2).Value Then
            If Trim(txtProxyServer.Text) = "" Then
                MsgBox GetMsg(msgEnterProxy), vbInformation
                tabIndex.CurrTab = 1
                txtProxyServer.SetFocus
                Exit Sub
            ElseIf Val(txtProxyPort.Text) <= 0 Then
                MsgBox GetMsg(msgEnterPort), vbInformation
                tabIndex.CurrTab = 1
                txtProxyPort.SetFocus
                Exit Sub
            End If
        End If
        If cboCMS.ListIndex < cboCMS.ListCount - 1 Then
            gAccount.CMS = cboCMS.ItemData(cboCMS.ListIndex)
        Else
            gAccount.CMS = CMS_CUSTOM
        End If
        gAccount.Service = bolService
        gAccount.Alias = txtAlias.Text
        gAccount.Host = txtHost.Text
        gAccount.Page = txtPage.Text
        gAccount.Port = Val(txtPort.Text)
        gAccount.Secure = chkSecure.Value
        gAccount.UTF8 = chkUTF8.Value
        gAccount.Timeout = Val(txtTimeout.Text)
        If optProxy(0).Value Then
            gAccount.UseProxy = 0
            gAccount.ProxyServer = ""
            gAccount.ProxyPort = 0
            gAccount.ProxyUser = ""
            gAccount.ProxyPassword = ""
        ElseIf optProxy(1).Value Then
            gAccount.UseProxy = 1
            gAccount.ProxyServer = ""
            gAccount.ProxyPort = 0
            gAccount.ProxyUser = ""
            gAccount.ProxyPassword = ""
        Else
            gAccount.UseProxy = 2
            gAccount.ProxyServer = txtProxyServer.Text
            gAccount.ProxyPort = Val(txtProxyPort.Text)
            gAccount.ProxyUser = txtUser.Text
            gAccount.ProxyPassword = txtPassword.Text
        End If
        If chkWeblogs.Value = vbChecked Then
            gAccount.PingWeblogs = cboPing.ListIndex + 1
        Else
            gAccount.PingWeblogs = 0
        End If
        'Custom Account API Settings
        If gAccount.CMS = CMS_CUSTOM Then
            Select Case cboPostAPI.ListIndex
            Case 0: gAccount.PostMethod = API_BLOGGER
            Case 1: gAccount.PostMethod = API_METAWEBLOG
            End Select
            gAccount.GetPostsMethod = gAccount.PostMethod
            Select Case cboCategAPI.ListIndex
            Case 0:    gAccount.GetCategMethod = API_NOTSUPPORTED
            Case 1, 2: gAccount.GetCategMethod = API_METAWEBLOG
            End Select
            gAccount.MultiCategory = (cboCategAPI.ListIndex = 2)
            Select Case cboTemplAPI.ListIndex
            Case 0: gAccount.TemplateMethod = API_NOTSUPPORTED
            Case 1: gAccount.TemplateMethod = API_BLOGGER
            End Select
            gAccount.TitleTag1 = txtTitleTag(0).Text
            gAccount.TitleTag2 = txtTitleTag(1).Text
            gAccount.CategTag1 = txtCategTag(0).Text
            gAccount.CategTag2 = txtCategTag(1).Text
            gAccount.MoreTextTag1 = txtMoreTag(0).Text
            gAccount.MoreTextTag2 = txtMoreTag(1).Text
        End If
        'Save Account
        If gAccount.User <> "" Then SaveAccount
    End If
    Unload Me
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub cmdReloadBlogs_Click()
    If LoadBlogs(False) Then
        LoadBlogSettings
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    LocalizeForm
    'Populate Fields
    LoadCombos
    txtAlias.Text = gAccount.Alias
    txtHost.Text = gAccount.Host
    txtPage.Text = gAccount.Page
    txtPort.Text = gAccount.Port
    chkSecure.Value = Abs(gAccount.Secure)
    chkUTF8.Value = Abs(gAccount.UTF8)
    If CBool(gAccount.PingWeblogs) Then
        chkWeblogs.Value = vbChecked
        cboPing.ListIndex = gAccount.PingWeblogs - 1
        cboPing.Enabled = True
    Else
        chkWeblogs.Value = vbUnchecked
        cboPing.ListIndex = 0
        cboPing.Enabled = False
    End If
    If gAccount.CMS = CMS_CUSTOM Then
        cboPostAPI.ListIndex = IIf(gAccount.PostMethod = API_BLOGGER, 0, 1)
        If gAccount.MultiCategory Then
            cboCategAPI.ListIndex = 2
        Else
            cboCategAPI.ListIndex = IIf(gAccount.GetCategMethod = API_NOTSUPPORTED, 0, 1)
        End If
        cboTemplAPI.ListIndex = IIf(gAccount.TemplateMethod = API_NOTSUPPORTED, 0, 1)
    End If
    txtTitleTag(0).Text = gAccount.TitleTag1
    txtTitleTag(1).Text = gAccount.TitleTag2
    txtCategTag(0).Text = gAccount.CategTag1
    txtCategTag(1).Text = gAccount.CategTag2
    txtMoreTag(0).Text = gAccount.MoreTextTag1
    txtMoreTag(1).Text = gAccount.MoreTextTag2
    cmdReloadBlogs.Enabled = True
    txtTimeout.Text = gAccount.Timeout
    optProxy(gAccount.UseProxy).Value = True
    txtProxyServer.Text = gAccount.ProxyServer
    txtProxyPort.Text = Format(gAccount.ProxyPort, "#")
    txtUser.Text = gAccount.ProxyUser
    txtPassword.Text = gAccount.ProxyPassword
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub LoadCombos()
On Error Resume Next
    Call LoadCMSCombo(cboCMS, imgCustom.Picture)
    Call LoadPingCombo(cboPing)
    cboCMS_Click
    'Populate API Method Combos
    cboPostAPI.AddItem "Blogger API"
    cboPostAPI.AddItem "metaWeblog API"
    cboPostAPI.ListIndex = 0
    cboCategAPI.AddItem GetLbl(lblNotSupported)
    cboCategAPI.AddItem "metaWeblog API - Single"
    cboCategAPI.AddItem "metaWeblog API - Multi"
    cboCategAPI.ListIndex = 0
    cboTemplAPI.AddItem GetLbl(lblNotSupported)
    cboTemplAPI.AddItem "Blogger API"
    cboTemplAPI.ListIndex = 1
End Sub

Private Sub optProxy_Click(Index As Integer)
    If optProxy(2).Value Then
        lblField(3).Enabled = True
        lblField(4).Enabled = True
        lblField(14).Enabled = True
        lblField(15).Enabled = True
        txtProxyServer.Enabled = True
        txtProxyPort.Enabled = True
        txtUser.Enabled = True
        txtPassword.Enabled = True
        txtProxyServer.BackColor = vbWindowBackground
        txtProxyPort.BackColor = vbWindowBackground
        txtUser.BackColor = vbWindowBackground
        txtPassword.BackColor = vbWindowBackground
    Else
        lblField(3).Enabled = False
        lblField(4).Enabled = False
        lblField(14).Enabled = False
        lblField(15).Enabled = False
        txtProxyServer.Enabled = False
        txtProxyPort.Enabled = False
        txtUser.Enabled = False
        txtPassword.Enabled = False
        txtProxyServer.BackColor = vbButtonFace
        txtProxyPort.BackColor = vbButtonFace
        txtUser.BackColor = vbButtonFace
        txtPassword.BackColor = vbButtonFace
    End If
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblConnection)
    tabIndex.TabCaption(0) = GetLbl(lblBloggerInfo)
    tabIndex.TabCaption(1) = GetLbl(lblProxy)
    tabIndex.TabCaption(2) = GetLbl(lblCustom)
    fraCMS.Caption = GetLbl(lblCMS)
    lblField(0).Caption = GetLbl(lblHost) & ":"
    lblField(1).Caption = GetLbl(lblPage) & ":"
    lblField(2).Caption = GetLbl(lblPort) & ":"
    lblField(3).Caption = GetLbl(lblAddress) & ":"
    lblField(4).Caption = GetLbl(lblPort) & ":"
    lblField(5).Caption = GetLbl(lblTimeout) & ":"
    lblField(6).Caption = GetLbl(lblSeconds)
    lblField(7).Caption = GetLbl(lblBlogTool) & ":"
    lblField(8).Caption = GetLbl(lblAccountAlias) & ":"
    lblField(9).Caption = GetLbl(lblPosts) & ":"
    lblField(10).Caption = GetLbl(lblCategories) & ":"
    lblField(11).Caption = GetLbl(lblTemplates) & ":"
    lblField(12).Caption = GetLbl(lblTitleTags) & ":"
    lblField(13).Caption = GetLbl(lblCategTags) & ":"
    lblField(14).Caption = GetLbl(lblUser) & ":"
    lblField(15).Caption = GetLbl(lblPassword) & ":"
    lblField(16).Caption = GetLbl(lblMoreTextTags) & ":"
    optProxy(0).Caption = GetLbl(lblNoProxy)
    optProxy(1).Caption = GetLbl(lblIEProxy)
    optProxy(2).Caption = GetLbl(lblMyProxy)
    chkWeblogs.Caption = GetLbl(lblPing) & ":"
    cmdReloadBlogs.Caption = GetLbl(lblReloadBlogs)
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
