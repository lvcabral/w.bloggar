VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne tabIndex 
      Height          =   3930
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   4950
      _cx             =   8731
      _cy             =   6932
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
      Caption         =   "General|Code Editor|Post Files"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   0   'False
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   -150
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
      Begin VB.Frame fraEditor 
         Caption         =   "Settings"
         Height          =   3255
         Left            =   5745
         TabIndex        =   32
         Top             =   480
         Width           =   4560
         Begin SizerOneLibCtl.ElasticOne ElasticOne1 
            Height          =   2970
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   240
            Width           =   4260
            _cx             =   7514
            _cy             =   5239
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
            Begin VB.CheckBox chkAutoConvert 
               Appearance      =   0  'Flat
               Caption         =   "Auto-Convert Extended Characters"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   30
               TabIndex        =   16
               Top             =   1485
               Value           =   1  'Checked
               Width           =   4335
            End
            Begin VB.CheckBox chkXHTML 
               Appearance      =   0  'Flat
               Caption         =   "Use XHTML Compatible Tags"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   30
               TabIndex        =   15
               Top             =   1140
               Width           =   4335
            End
            Begin VB.TextBox txtTabSpc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3855
               MaxLength       =   2
               TabIndex        =   22
               Top             =   2535
               Width           =   375
            End
            Begin VB.CheckBox chkColorize 
               Appearance      =   0  'Flat
               Caption         =   "Colorize HTML Code"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   30
               TabIndex        =   13
               Top             =   465
               Value           =   1  'Checked
               Width           =   4335
            End
            Begin VB.ComboBox cboFontFace 
               Height          =   315
               ItemData        =   "Settings.frx":000C
               Left            =   1380
               List            =   "Settings.frx":000E
               Sorted          =   -1  'True
               TabIndex        =   18
               Top             =   2100
               Width           =   2850
            End
            Begin VB.ComboBox cboFontSize 
               Height          =   315
               ItemData        =   "Settings.frx":0010
               Left            =   1380
               List            =   "Settings.frx":002C
               TabIndex        =   20
               Top             =   2535
               Width           =   825
            End
            Begin VB.CheckBox chkClear 
               Appearance      =   0  'Flat
               Caption         =   "Clear Editor after Post"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   30
               TabIndex        =   12
               Top             =   135
               Value           =   1  'Checked
               Width           =   4335
            End
            Begin VB.CheckBox chkHTML 
               Appearance      =   0  'Flat
               Caption         =   "Show HTML Toolbar on Template Edit"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   30
               TabIndex        =   14
               Top             =   795
               Value           =   1  'Checked
               Width           =   4335
            End
            Begin VB.Label lblField 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tab S&paces:"
               Height          =   195
               Index           =   6
               Left            =   2850
               TabIndex        =   21
               Top             =   2580
               Width           =   885
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               Caption         =   "&Font:"
               Height          =   195
               Index           =   3
               Left            =   30
               TabIndex        =   17
               Top             =   2145
               Width           =   390
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               Caption         =   "Font &Size:"
               Height          =   195
               Index           =   2
               Left            =   30
               TabIndex        =   19
               Top             =   2580
               Width           =   720
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000014&
               Index           =   0
               X1              =   15
               X2              =   4200
               Y1              =   1950
               Y2              =   1950
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000010&
               Index           =   1
               X1              =   30
               X2              =   4245
               Y1              =   1935
               Y2              =   1935
            End
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne5 
         Height          =   3255
         Left            =   6045
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   480
         Width           =   4560
         _cx             =   8043
         _cy             =   5741
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
         Begin VB.Frame fraDefault 
            Caption         =   "Default Post"
            Height          =   1380
            Left            =   0
            TabIndex        =   38
            Top             =   1860
            Width           =   4545
            Begin SizerOneLibCtl.ElasticOne ElasticOne7 
               Height          =   1050
               Left            =   120
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   255
               Width           =   4320
               _cx             =   7620
               _cy             =   1852
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
               Begin VB.TextBox txtDefault 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   26
                  Top             =   360
                  Width           =   3975
               End
               Begin VB.CommandButton cmdFile 
                  Caption         =   "..."
                  Height          =   300
                  Index           =   0
                  Left            =   3990
                  TabIndex        =   27
                  Top             =   360
                  Width           =   315
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "New Post button loads this file (optional):"
                  Height          =   195
                  Index           =   8
                  Left            =   0
                  TabIndex        =   25
                  Top             =   105
                  Width           =   2970
               End
            End
         End
         Begin VB.Frame fraFiles 
            Caption         =   "Settings"
            Height          =   1770
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   4545
            Begin VB.CheckBox chkOpenLast 
               Appearance      =   0  'Flat
               Caption         =   "Reopen last used .post file at startup"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   150
               TabIndex        =   24
               Top             =   705
               Value           =   1  'Checked
               Width           =   4185
            End
            Begin SizerOneLibCtl.ElasticOne ElasticOne8 
               Height          =   630
               Left            =   60
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   210
               Width           =   4350
               _cx             =   7673
               _cy             =   1111
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
               Begin VB.CheckBox chkAssociate 
                  Appearance      =   0  'Flat
                  Caption         =   "Associate *.post files with w.bloggar"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   90
                  TabIndex        =   23
                  Top             =   150
                  Value           =   1  'Checked
                  Width           =   4185
               End
            End
         End
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Settings"
         Height          =   3255
         Left            =   195
         TabIndex        =   33
         Top             =   480
         Width           =   4560
         Begin SizerOneLibCtl.ElasticOne ElasticOne3 
            Height          =   2835
            Left            =   120
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   255
            Width           =   4410
            _cx             =   7779
            _cy             =   5001
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
            BorderWidth     =   0
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
            Begin VB.ComboBox cboBlogListSize 
               Height          =   315
               ItemData        =   "Settings.frx":004E
               Left            =   1560
               List            =   "Settings.frx":0082
               TabIndex        =   5
               Text            =   "cboBlogListSize"
               Top             =   1260
               Width           =   825
            End
            Begin VB.ComboBox cboLang 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   2100
               Width           =   2700
            End
            Begin VB.CheckBox chkDefBrowser 
               Appearance      =   0  'Flat
               Caption         =   "Open Web Pages using the Default Browser"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   60
               TabIndex        =   3
               Top             =   930
               Width           =   4335
            End
            Begin VB.ComboBox cboSkin 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   1680
               Width           =   2700
            End
            Begin VB.CheckBox chkSilent 
               Appearance      =   0  'Flat
               Caption         =   "Don't Show Message Boxes on Success"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   60
               TabIndex        =   2
               Top             =   645
               Width           =   4335
            End
            Begin VB.CheckBox chkTray 
               Appearance      =   0  'Flat
               Caption         =   "Minimize to Tray"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   60
               TabIndex        =   0
               Top             =   75
               Value           =   1  'Checked
               Width           =   4335
            End
            Begin VB.CheckBox chkMinimized 
               Appearance      =   0  'Flat
               Caption         =   "Start Minimized"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   60
               TabIndex        =   1
               Top             =   360
               Width           =   4335
            End
            Begin VB.ComboBox cboDictionary 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   2520
               Width           =   2700
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               Caption         =   "pixels"
               Height          =   195
               Index           =   5
               Left            =   2490
               TabIndex        =   40
               Top             =   1305
               Width           =   405
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               Caption         =   "Blog List &Size:"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   4
               Top             =   1305
               Width           =   975
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               Caption         =   "Language:"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   8
               Top             =   2145
               Width           =   765
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               Caption         =   "Toolbar S&kin:"
               Height          =   195
               Index           =   7
               Left            =   60
               TabIndex        =   6
               Top             =   1740
               Width           =   930
            End
            Begin VB.Label lblField 
               AutoSize        =   -1  'True
               Caption         =   "Dictionary:"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   10
               Top             =   2565
               Width           =   780
            End
         End
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2655
      TabIndex        =   28
      Top             =   4155
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3915
      TabIndex        =   29
      Top             =   4155
      Width           =   1155
   End
End
Attribute VB_Name = "frmSettings"
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

Private Sub cboFontFace_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboFontFace, KeyAscii
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim lngOldStart As Long, lngOldLen As Long
    If Index = 0 Then
        If Val(cboBlogListSize.Text) < 50 Or Val(cboBlogListSize.Text) > 200 Then
            MsgBox GetMsg(msgInvalidBlogListSize), vbInformation
            tabIndex.CurrTab = 0
            cboBlogListSize.SetFocus
            Exit Sub
        ElseIf cboFontFace.ListIndex < 0 Then
            MsgBox GetMsg(msgInvalidFontFace), vbInformation
            tabIndex.CurrTab = 1
            cboFontFace.SetFocus
            Exit Sub
        ElseIf Val(cboFontSize.Text) < 6 Or Val(cboFontSize.Text) > 32 Then
            MsgBox GetMsg(msgInvalidFontSize), vbInformation
            tabIndex.CurrTab = 1
            cboFontSize.SetFocus
            Exit Sub
        ElseIf Val(txtTabSpc.Text) < 1 Then
            MsgBox GetMsg(msgEnterTabSpaces), vbInformation
            tabIndex.CurrTab = 1
            txtTabSpc.SetFocus
            Exit Sub
        ElseIf Trim(txtDefault.Text) <> "" And _
               Not FileExists(txtDefault.Text) Then
            MsgBox GetMsg(msgDefPostNotFound), vbInformation
            tabIndex.CurrTab = 2
            txtDefault.SetFocus
            Exit Sub
        End If
        'Write General Settings
        gSettings.Tray = chkTray.Value
        gSettings.StartMinimized = chkMinimized.Value
        gSettings.ClearPost = chkClear.Value
        gSettings.ColorizeCode = chkColorize.Value
        gSettings.ShowHtmlBar = chkHTML.Value
        gSettings.AutoConvert = chkAutoConvert.Value
        gSettings.Silent = chkSilent.Value
        gSettings.DefaultBrowser = chkDefBrowser.Value
        gSettings.BlogListSize = Val(cboBlogListSize.Text)
        If cboSkin.ListIndex >= 0 Then
            gSettings.SkinFolder = App.Path & "\skins\" & cboSkin.Text
        End If
        If cboLang.ListIndex >= 0 Then
            gSettings.AppLCID = cboLang.ItemData(cboLang.ListIndex)
        End If
        If cboDictionary.ListIndex >= 0 Then
            gSettings.SpellLCID = cboDictionary.ItemData(cboDictionary.ListIndex)
        End If
        'Verify SpellChecking
        frmPost.acbMain.Tools("miTSpelling").Enabled = FileExists(App.Path & "\spell\" & gSettings.SpellLCID & ".dic")
        If gSettings.FontSize <> Val(cboFontSize.Text) Or _
           gSettings.FontFace <> cboFontFace.Text Then
            frmPost.txtPost.Font.Size = Val(cboFontSize.Text)
            frmPost.txtPost.Font.Name = cboFontFace.Text
            frmPost.txtMore.Font.Size = Val(cboFontSize.Text)
            frmPost.txtMore.Font.Name = cboFontFace.Text
            frmPost.txtExcerpt.Font.Size = Val(cboFontSize.Text)
            frmPost.txtExcerpt.Font.Name = cboFontFace.Text
        End If
        gSettings.TabSpaces = Val(txtTabSpc.Text)
        If Not frmPost.txtPost.AutoColorize And gSettings.ColorizeCode Then
            frmPost.txtPost.AutoColorize = gSettings.ColorizeCode
            frmPost.txtPost.Colorize
        ElseIf frmPost.txtPost.AutoColorize And Not gSettings.ColorizeCode Then
            frmPost.txtPost.AutoColorize = gSettings.ColorizeCode
            lngOldStart = frmPost.txtPost.SelStart
            lngOldLen = frmPost.txtPost.SelLength
            frmPost.txtPost.Text = frmPost.txtPost.Text
            frmPost.txtPost.SelStart = lngOldStart
            frmPost.txtPost.SelLength = lngOldLen
        ElseIf gSettings.FontSize <> Val(cboFontSize.Text) Or _
               gSettings.FontFace <> cboFontFace.Text Then
            frmPost.txtPost.AutoColorize = gSettings.ColorizeCode
            frmPost.txtPost.Colorize
        End If
        gSettings.FontFace = cboFontFace.Text
        gSettings.FontSize = Val(cboFontSize.Text)
        gSettings.XHTML = chkXHTML.Value
        'Post File Settings
        gSettings.OpenLastFile = chkOpenLast.Value
        Call Associate(chkAssociate.Value = vbChecked)
        gSettings.PostTemplate = txtDefault.Text
        'Set the Blog List Size
        frmPost.acbMain.Bands("bndTools").Tools("miBlogs").Width = gSettings.BlogListSize * Screen.TwipsPerPixelX
        'Set Tray Option
        frmPost.acfPost.MinimizeToTray = gSettings.Tray
        'Save on Registry
        SaveAppSettings
        'Reload the selected Skin
        frmPost.LoadSkin
    End If
    Unload Me
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub cmdFile_Click(Index As Integer)
On Error Resume Next
Dim oFile As New FileDialog
    oFile.DialogTitle = fraDefault.Caption
    oFile.Filter = GetMsg(msgFileFilter)
    oFile.Flags = cdlFileMustExist Or cdlLongnames Or cdlHideReadOnly
    oFile.hWndParent = Me.hwnd
    oFile.ShowOpen
    If oFile.FileName <> "" Then
        txtDefault.Text = oFile.FileName
    End If
    Set oFile = Nothing
End Sub

Private Sub txtTabSpc_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    tabIndex.CurrTab = 0
    LocalizeForm
    chkTray.Value = Abs(gSettings.Tray)
    chkMinimized.Value = Abs(gSettings.StartMinimized)
    chkClear.Value = Abs(gSettings.ClearPost)
    chkColorize.Value = Abs(gSettings.ColorizeCode)
    chkHTML.Value = Abs(gSettings.ShowHtmlBar)
    chkAutoConvert.Value = Abs(gSettings.AutoConvert)
    chkSilent.Value = Abs(gSettings.Silent)
    chkDefBrowser.Value = Abs(gSettings.DefaultBrowser)
    chkXHTML.Value = Abs(gSettings.XHTML)
    LoadFonts
    LoadSkins
    LoadLangs
    LoadDicts
    cboFontSize.Text = gSettings.FontSize
    txtTabSpc.Text = gSettings.TabSpaces
    chkAssociate.Value = Abs(IsAssociated())
    chkOpenLast.Value = Abs(gSettings.OpenLastFile)
    txtDefault.Text = gSettings.PostTemplate
    cboBlogListSize.Text = Format(gSettings.BlogListSize)
End Sub

Private Sub LoadFonts()
On Error Resume Next
Dim f As Integer, t As Integer
Dim i As Integer
    t = Screen.FontCount
    For f = 0 To t - 1
        cboFontFace.AddItem Screen.Fonts(f)
    Next
    i = -1
    For f = 0 To t - 1
        If cboFontFace.List(f) = gSettings.FontFace Then
            i = f
            Exit For
        End If
    Next
    If i < 0 Then i = 0
    cboFontFace.ListIndex = i
End Sub

Private Sub LoadSkins()
Dim strSkin As String
Dim s As Integer
On Error Resume Next
    strSkin = Dir(App.Path & "\skins\", vbDirectory)
    Do While strSkin <> ""
        If Left(strSkin, 1) <> "." Then
            If FileExists(App.Path & "\skins\" & strSkin & "\skin.ini") Then
                cboSkin.AddItem strSkin
            End If
        End If
        strSkin = Dir()
    Loop
    For s = 0 To cboSkin.ListCount - 1
        If LCase(cboSkin.List(s)) = LCase(GetNamePart(gSettings.SkinFolder)) Then
            cboSkin.ListIndex = s
            Exit For
        End If
    Next
    If cboSkin.ListIndex = -1 And cboSkin.ListCount > 0 Then
        cboSkin.ListIndex = 0
    End If
End Sub

Private Sub LoadDicts()
Dim strDict As String
Dim d As Integer
On Error Resume Next
    strDict = Dir(App.Path & "\Spell\*.dic")
    Do While strDict <> ""
        If Len(strDict) = 8 Then
            strDict = Left(strDict, 4)
            cboDictionary.AddItem GetLocaleName(Val(strDict))
            cboDictionary.ItemData(cboDictionary.NewIndex) = Val(strDict)
        End If
        strDict = Dir()
    Loop
    For d = 0 To cboDictionary.ListCount - 1
        If Val(cboDictionary.ItemData(d)) = gSettings.SpellLCID Then
            cboDictionary.ListIndex = d
            Exit For
        End If
    Next
    If cboDictionary.ListIndex = -1 And cboDictionary.ListCount > 0 Then
        cboDictionary.ListIndex = 0
    End If
End Sub

Private Sub LoadLangs()
Dim strLang As String
Dim d As Integer
On Error Resume Next
    strLang = Dir(App.Path & "\Lang\*.lng")
    Do While strLang <> ""
        If Len(strLang) = 8 Then
            strLang = Left(strLang, 4)
            cboLang.AddItem GetLocaleName(Val(strLang))
            cboLang.ItemData(cboLang.NewIndex) = Val(strLang)
        End If
        strLang = Dir()
    Loop
    For d = 0 To cboLang.ListCount - 1
        If Val(cboLang.ItemData(d)) = gSettings.AppLCID Then
            cboLang.ListIndex = d
            Exit For
        End If
    Next
    If cboLang.ListIndex = -1 And cboLang.ListCount > 0 Then
        cboLang.ListIndex = 0
    End If
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblOptions)
    tabIndex.TabCaption(0) = GetLbl(lblGeneral)
    tabIndex.TabCaption(1) = GetLbl(lblCodeEditor)
    tabIndex.TabCaption(2) = GetLbl(lblPostFiles)
    
    fraOptions.Caption = GetLbl(lblSettings)
    chkTray.Caption = GetLbl(lblMinimizeTray)
    chkMinimized.Caption = GetLbl(lblStartMin)
    chkClear.Caption = GetLbl(lblClearAfter)
    chkColorize.Caption = GetLbl(lblColorize)
    chkHTML.Caption = GetLbl(lblShowHTML)
    chkAutoConvert.Caption = GetLbl(lblAutoConvert)
    chkSilent.Caption = GetLbl(lblSilentPost)
    chkDefBrowser.Caption = GetLbl(lblDefaultBrowser)
    chkXHTML.Caption = GetLbl(lblUseXHTML)
    chkOpenLast.Caption = GetLbl(lblReopenLastFile)
    lblField(0).Caption = GetLbl(lblLanguage) & ":"
    lblField(1).Caption = GetLbl(lblDictionary) & ":"
    lblField(2).Caption = GetLbl(lblFontSize) & ":"
    lblField(3).Caption = GetLbl(lblFontFace) & ":"
    lblField(4).Caption = GetLbl(lblBlogListSize) & ":"
    lblField(5).Caption = GetLbl(lblPixels)
    lblField(6).Caption = GetLbl(lblTabSpaces) & ":"
    lblField(7).Caption = GetLbl(lblToolbarSkin) & ":"
    lblField(8).Caption = GetLbl(lblLoadDefault) & ":"
    fraEditor.Caption = GetLbl(lblSettings)
    fraFiles.Caption = GetLbl(lblSettings)
    chkAssociate.Caption = GetLbl(lblAssociate)
    fraDefault.Caption = GetLbl(lblDefaultPost)
    
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
