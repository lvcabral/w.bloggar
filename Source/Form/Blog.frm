VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Begin VB.Form frmBlog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Blog Properties"
   ClientHeight    =   6060
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
   Icon            =   "Blog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne tabIndex 
      Height          =   5355
      Left            =   105
      TabIndex        =   31
      Top             =   120
      Width           =   4950
      _cx             =   8731
      _cy             =   9446
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
      Caption         =   "Preview|Upload|Media"
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4680
         Left            =   195
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   480
         Width           =   4560
         _cx             =   8043
         _cy             =   8255
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
         Begin VB.Frame fraHTML 
            Caption         =   "Preview Format:"
            Height          =   3810
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   4545
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               HasDC           =   0   'False
               Height          =   3525
               Left            =   75
               ScaleHeight     =   3525
               ScaleWidth      =   4440
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   225
               Width           =   4440
               Begin VB.TextBox txtCSS 
                  Height          =   315
                  Left            =   45
                  TabIndex        =   3
                  Top             =   960
                  Width           =   4290
               End
               Begin VB.CheckBox chkConvertBR 
                  Appearance      =   0  'Flat
                  Caption         =   "Convert Line-Break to <br> tag"
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   60
                  TabIndex        =   12
                  Top             =   2775
                  Value           =   1  'Checked
                  Width           =   3660
               End
               Begin VB.TextBox txtAlign 
                  Height          =   315
                  Left            =   45
                  TabIndex        =   9
                  Top             =   2280
                  Width           =   2070
               End
               Begin VB.TextBox txtPostStyle 
                  Height          =   315
                  Left            =   2265
                  TabIndex        =   7
                  Top             =   1620
                  Width           =   2070
               End
               Begin VB.TextBox txtBody 
                  Height          =   315
                  Left            =   45
                  TabIndex        =   1
                  Top             =   330
                  Width           =   4290
               End
               Begin VB.TextBox txtTitleStyle 
                  Height          =   315
                  Left            =   45
                  TabIndex        =   5
                  Top             =   1620
                  Width           =   2070
               End
               Begin VB.TextBox txtWidth 
                  Height          =   315
                  Left            =   2265
                  TabIndex        =   11
                  Top             =   2280
                  Width           =   2070
               End
               Begin VB.CommandButton cmdDefault 
                  Caption         =   "&Restore Defaults"
                  Height          =   345
                  Left            =   615
                  TabIndex        =   13
                  Top             =   3090
                  Width           =   3150
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "CSS Tag:"
                  Height          =   195
                  Index           =   18
                  Left            =   45
                  TabIndex        =   2
                  Top             =   705
                  Width           =   660
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Post Alignment:"
                  Height          =   195
                  Index           =   6
                  Left            =   45
                  TabIndex        =   8
                  Top             =   2025
                  Width           =   1125
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Post Style / Class:"
                  Height          =   195
                  Index           =   5
                  Left            =   2280
                  TabIndex        =   6
                  Top             =   1365
                  Width           =   1305
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Body Tag:"
                  Height          =   195
                  Index           =   4
                  Left            =   45
                  TabIndex        =   0
                  Top             =   75
                  Width           =   735
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Title Style / Class:"
                  Height          =   195
                  Index           =   0
                  Left            =   45
                  TabIndex        =   4
                  Top             =   1365
                  Width           =   1290
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Post Width:"
                  Height          =   195
                  Index           =   14
                  Left            =   2265
                  TabIndex        =   10
                  Top             =   2025
                  Width           =   840
               End
            End
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000014&
            BorderStyle     =   3  'Dot
            Height          =   735
            Index           =   2
            Left            =   15
            Top             =   3945
            Width           =   4515
         End
         Begin VB.Image imgTip 
            Height          =   480
            Index           =   1
            Left            =   45
            Top             =   4035
            Width           =   480
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000010&
            BorderStyle     =   3  'Dot
            Height          =   735
            Index           =   3
            Left            =   0
            Top             =   3930
            Width           =   4515
         End
         Begin VB.Label lblTip 
            Caption         =   "Note: Use this (per blog) settings to make the preview as close as possible of your blog's template."
            Height          =   600
            Index           =   1
            Left            =   585
            TabIndex        =   36
            Top             =   4005
            Width           =   3750
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   4680
         Left            =   5745
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   480
         Width           =   4560
         _cx             =   8043
         _cy             =   8255
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
         Begin VB.Frame fraFTP 
            Caption         =   "Upload:"
            Height          =   3375
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   4545
            Begin SizerOneLibCtl.ElasticOne ElasticOne4 
               Height          =   3075
               Left            =   150
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   210
               Width           =   4335
               _cx             =   7646
               _cy             =   5424
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
               Begin VB.OptionButton optUpload 
                  Appearance      =   0  'Flat
                  Caption         =   "Blog API"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   2310
                  TabIndex        =   15
                  Top             =   75
                  Width           =   2040
               End
               Begin VB.OptionButton optUpload 
                  Appearance      =   0  'Flat
                  Caption         =   "FTP Server"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   60
                  TabIndex        =   14
                  Top             =   75
                  Width           =   2040
               End
               Begin VB.TextBox txtPath 
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1305
                  MaxLength       =   255
                  TabIndex        =   19
                  Top             =   1005
                  Width           =   2925
               End
               Begin VB.TextBox txtHost 
                  Height          =   315
                  Left            =   1305
                  MaxLength       =   255
                  TabIndex        =   17
                  Top             =   585
                  Width           =   2925
               End
               Begin VB.CheckBox chkProxy 
                  Appearance      =   0  'Flat
                  Caption         =   "&Use Account Proxy"
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   1890
                  TabIndex        =   22
                  Top             =   1425
                  Width           =   2385
               End
               Begin VB.TextBox txtPort 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1305
                  MaxLength       =   5
                  TabIndex        =   21
                  Text            =   "21"
                  Top             =   1410
                  Width           =   480
               End
               Begin VB.TextBox txtUser 
                  Height          =   315
                  Left            =   1305
                  TabIndex        =   24
                  Top             =   1815
                  Width           =   1755
               End
               Begin VB.TextBox txtPassword 
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1320
                  PasswordChar    =   "*"
                  TabIndex        =   26
                  Top             =   2235
                  Width           =   1755
               End
               Begin VB.TextBox txtLink 
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1320
                  MaxLength       =   255
                  TabIndex        =   28
                  Top             =   2640
                  Width           =   2925
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H80000010&
                  Index           =   1
                  X1              =   15
                  X2              =   4230
                  Y1              =   405
                  Y2              =   405
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H80000014&
                  Index           =   0
                  X1              =   0
                  X2              =   4185
                  Y1              =   420
                  Y2              =   420
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Remote P&ath:"
                  Height          =   195
                  Index           =   8
                  Left            =   75
                  TabIndex        =   18
                  Top             =   1065
                  Width           =   990
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "&Host:"
                  Height          =   195
                  Index           =   9
                  Left            =   75
                  TabIndex        =   16
                  Top             =   630
                  Width           =   390
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Por&t:"
                  Height          =   195
                  Index           =   10
                  Left            =   75
                  TabIndex        =   20
                  Top             =   1470
                  Width           =   360
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "&User:"
                  Height          =   195
                  Index           =   11
                  Left            =   75
                  TabIndex        =   23
                  Top             =   1875
                  Width           =   390
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "&Password:"
                  Height          =   195
                  Index           =   12
                  Left            =   75
                  TabIndex        =   25
                  Top             =   2295
                  Width           =   750
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "URL to &Link:"
                  Height          =   195
                  Index           =   13
                  Left            =   75
                  TabIndex        =   27
                  Top             =   2700
                  Width           =   855
               End
            End
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000014&
            BorderStyle     =   3  'Dot
            Height          =   1140
            Index           =   1
            Left            =   15
            Top             =   3540
            Width           =   4530
         End
         Begin VB.Image imgTip 
            Height          =   480
            Index           =   0
            Left            =   45
            Picture         =   "Blog.frx":000C
            Top             =   3615
            Width           =   480
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000010&
            BorderStyle     =   3  'Dot
            Height          =   1140
            Index           =   0
            Left            =   0
            Top             =   3525
            Width           =   4530
         End
         Begin VB.Label lblTip 
            Caption         =   $"Blog.frx":08D6
            Height          =   1005
            Index           =   0
            Left            =   615
            TabIndex        =   34
            Top             =   3615
            Width           =   3750
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne9 
         Height          =   4680
         Left            =   6045
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   480
         Width           =   4560
         _cx             =   8043
         _cy             =   8255
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
         Begin VB.Frame fraMedia 
            Caption         =   "Add Media Information"
            Height          =   3645
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   4545
            Begin SizerOneLibCtl.ElasticOne ElasticOne10 
               Height          =   3360
               Left            =   150
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   210
               Width           =   4335
               _cx             =   7646
               _cy             =   5927
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
               Begin VB.CheckBox chkMediaLink 
                  Appearance      =   0  'Flat
                  Caption         =   "Add Link to Media Info on Artist Name"
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   0
                  TabIndex        =   55
                  Top             =   1740
                  Width           =   4290
               End
               Begin VB.CommandButton cmdMediaField 
                  Caption         =   "Artist"
                  Height          =   315
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   52
                  Top             =   2805
                  Width           =   1050
               End
               Begin VB.CommandButton cmdMediaField 
                  Caption         =   "Duration"
                  Height          =   315
                  Index           =   3
                  Left            =   3240
                  TabIndex        =   51
                  Top             =   2805
                  Width           =   1050
               End
               Begin VB.CommandButton cmdMediaField 
                  Caption         =   "Album"
                  Height          =   315
                  Index           =   2
                  Left            =   2160
                  TabIndex        =   50
                  Top             =   2805
                  Width           =   1050
               End
               Begin VB.CommandButton cmdMediaField 
                  Caption         =   "Title"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   49
                  Top             =   2805
                  Width           =   1050
               End
               Begin VB.OptionButton optMedia 
                  Appearance      =   0  'Flat
                  Caption         =   "Automatic Insert on Bottom of the New Posts"
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   2
                  Left            =   0
                  TabIndex        =   48
                  Top             =   1395
                  Width           =   4410
               End
               Begin VB.OptionButton optMedia 
                  Appearance      =   0  'Flat
                  Caption         =   "Automatic Insert on Top of the New Posts"
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   1
                  Left            =   0
                  TabIndex        =   47
                  Top             =   1080
                  Width           =   4410
               End
               Begin VB.OptionButton optMedia 
                  Appearance      =   0  'Flat
                  Caption         =   "Manual Insert: Click on Status Bar Icon or Press F11"
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   0
                  Left            =   0
                  TabIndex        =   46
                  Top             =   780
                  Value           =   -1  'True
                  Width           =   4410
               End
               Begin VB.TextBox txtMedia 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   44
                  Top             =   2415
                  Width           =   4290
               End
               Begin VB.Image Image1 
                  Height          =   240
                  Left            =   15
                  Picture         =   "Blog.frx":0996
                  Top             =   60
                  Width           =   240
               End
               Begin VB.Label lblMediaOpt 
                  Caption         =   $"Blog.frx":0F20
                  Height          =   705
                  Left            =   360
                  TabIndex        =   54
                  Top             =   60
                  Width           =   3915
               End
               Begin VB.Label lblPlaceholders 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   ":: These placeholders can also be used on Post body ::"
                  Height          =   225
                  Left            =   -15
                  TabIndex        =   53
                  Top             =   3135
                  Width           =   4335
               End
               Begin VB.Label lblField 
                  AutoSize        =   -1  'True
                  Caption         =   "Media String:"
                  Height          =   195
                  Index           =   20
                  Left            =   0
                  TabIndex        =   45
                  Top             =   2160
                  Width           =   945
               End
            End
         End
         Begin VB.Label lblTip 
            Caption         =   $"Blog.frx":0FA7
            Height          =   780
            Index           =   2
            Left            =   615
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   3855
            Width           =   3750
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000010&
            BorderStyle     =   3  'Dot
            Height          =   900
            Index           =   5
            Left            =   0
            Top             =   3765
            Width           =   4530
         End
         Begin VB.Image imgTip 
            Height          =   480
            Index           =   2
            Left            =   60
            Top             =   3855
            Width           =   480
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000014&
            BorderStyle     =   3  'Dot
            Height          =   900
            Index           =   4
            Left            =   15
            Top             =   3780
            Width           =   4530
         End
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2685
      TabIndex        =   29
      Top             =   5565
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3915
      TabIndex        =   30
      Top             =   5565
      Width           =   1155
   End
End
Attribute VB_Name = "frmBlog"
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
Public Enum BlogSettingsTabs
    enuTabPreview
    enuTabUpload
    enuTabMedia
End Enum


Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim lngOldStart As Long, lngOldLen As Long
    If Index = 0 Then
        If LCase(Left(Trim(txtBody.Text), 5)) <> "<body" Or _
               Right(Trim(txtBody.Text), 1) <> ">" Then
            MsgBox GetMsg(msgInvalidBody), vbInformation
            tabIndex.CurrTab = enuTabPreview
            txtBody.SetFocus
            Exit Sub
        ElseIf Len(Trim(txtHost.Text)) > 0 And Val(txtPort.Text) <= 0 Then
            MsgBox GetMsg(msgEnterPort), vbInformation
            tabIndex.CurrTab = enuTabUpload
            txtPort.SetFocus
            Exit Sub
        End If
        'Write Blog Settings
        gBlog.PreviewBody = txtBody.Text
        gBlog.PreviewCSS = txtCSS.Text
        gBlog.PreviewTitle = txtTitleStyle.Text
        gBlog.PreviewStyle = txtPostStyle.Text
        gBlog.PreviewAlign = txtAlign.Text
        gBlog.PreviewWidth = txtWidth.Text
        gBlog.PreviewAutoBR = CBool(chkConvertBR.Value)
        gBlog.APIUpload = optUpload(1).Value
        gBlog.FTPHost = txtHost.Text
        txtPath.Text = Replace(txtPath.Text, "\", "/")
        gBlog.FTPPath = txtPath.Text & IIf(Right(txtPath.Text, 1) = "/", "", "/")
        gBlog.FTPPort = Val(txtPort.Text)
        gBlog.FTPProxy = CBool(chkProxy.Value)
        gBlog.FTPUser = txtUser.Text
        gBlog.FTPPassword = txtPassword.Text
        txtLink.Text = Replace(txtLink.Text, "\", "/")
        If Trim(txtLink.Text) <> "" Then
            gBlog.FTPLink = txtLink.Text & IIf(Right(txtLink.Text, 1) = "/", "", "/")
        Else
            gBlog.FTPLink = ""
        End If
        If optMedia(0).Value Then
            gBlog.MediaInsert = 0
        ElseIf optMedia(1).Value Then
            gBlog.MediaInsert = 1
        Else
            gBlog.MediaInsert = 2
        End If
        gBlog.MediaLink = CBool(chkMediaLink.Value)
        gBlog.MediaString = txtMedia.Text
        SaveBlogSettings
    End If
    Unload Me
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub cmdDefault_Click()
    txtBody.Text = BODYTAG
    txtCSS.Text = CSSTAG
    txtTitleStyle.Text = TITLESTYLE
    txtPostStyle.Text = POSTSTYLE
    txtAlign.Text = POSTALIGN
End Sub

Private Sub cmdMediaField_Click(Index As Integer)
On Error Resume Next
    Select Case Index
    Case 0: txtMedia.SelText = "%T%"
    Case 1: txtMedia.SelText = "%A%"
    Case 2: txtMedia.SelText = "%B%"
    Case 3: txtMedia.SelText = "%D%"
    End Select
    txtMedia.SetFocus
End Sub

Private Sub lblTip_Click(Index As Integer)
    If Index = 2 Then
        frmPost.OpenWebPage "http://wbloggar.com/download.php"
    End If
End Sub

Private Sub Form_Load()
    tabIndex.CurrTab = enuTabPreview
    LocalizeForm
    lblTip(2).MouseIcon = frmPost.imgHand.Picture
    imgTip(1).Picture = imgTip(0).Picture
    imgTip(2).Picture = imgTip(0).Picture
    RefreshBlog
End Sub

Private Sub RefreshBlog()
On Error Resume Next
    fraHTML.Caption = GetLbl(lblPreviewFmt) & ": " & gBlogs(frmPost.CurrentBlog).Name
    fraFTP.Caption = GetLbl(lblUpload) & ": " & gBlogs(frmPost.CurrentBlog).Name
    fraMedia.Caption = GetLbl(lblAddMediaInfo) & ": " & gBlogs(frmPost.CurrentBlog).Name
    txtBody.Text = gBlog.PreviewBody
    txtCSS.Text = gBlog.PreviewCSS
    txtTitleStyle.Text = gBlog.PreviewTitle
    txtPostStyle.Text = gBlog.PreviewStyle
    txtAlign.Text = gBlog.PreviewAlign
    txtWidth.Text = gBlog.PreviewWidth
    chkConvertBR.Value = Abs(gBlog.PreviewAutoBR)
    chkProxy.Enabled = gAccount.UseProxy
    optUpload(1).Caption = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Name", App.Path & "\CMS\CMS.ini")
    If gAccount.UploadMethod = API_METAWEBLOG Then
        optUpload(1).Value = Abs(gBlog.APIUpload)
    Else
        optUpload(0).Value = True
        optUpload(1).Enabled = False
    End If
    txtHost.Text = gBlog.FTPHost
    txtPath.Text = gBlog.FTPPath
    txtPort.Text = gBlog.FTPPort
    chkProxy.Value = Abs(IIf(chkProxy.Enabled, gBlog.FTPProxy, False))
    txtUser.Text = gBlog.FTPUser
    txtPassword.Text = gBlog.FTPPassword
    txtLink.Text = gBlog.FTPLink
    optMedia(gBlog.MediaInsert).Value = True
    chkMediaLink.Value = Abs(gBlog.MediaLink)
    If gBlog.MediaString = "" Then gBlog.MediaString = Replace(MEDIASTR, "%1%", GetLbl(lblListening))
    txtMedia.Text = gBlog.MediaString
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    'Me.Caption = GetLbl(lblSettings)
    tabIndex.TabCaption(enuTabPreview) = GetLbl(lblPreview)
    tabIndex.TabCaption(enuTabUpload) = GetLbl(lblUpload)
    tabIndex.TabCaption(enuTabMedia) = GetLbl(lblMedia)
    
    fraHTML.Caption = GetLbl(lblPreviewFmt)
    lblField(4).Caption = GetLbl(lblBodyTag) & ":"
    lblField(18).Caption = GetLbl(lblCSSTag) & ":"
    lblField(0).Caption = GetLbl(lblTitleStyle) & ":"
    lblField(5).Caption = GetLbl(lblPostStyle) & ":"
    lblField(6).Caption = GetLbl(lblPostAlign) & ":"
    lblField(14).Caption = GetLbl(lblPostWidth) & ":"
    chkConvertBR.Caption = GetLbl(lblConvertBR)
    cmdDefault.Caption = GetLbl(lblRestoreDef)
    lblTip(1).Caption = GetLbl(lblTipBlog)
    
    fraFTP.Caption = GetLbl(lblUpload)
    optUpload(0).Caption = GetLbl(lblFTPServer)
    lblField(9).Caption = GetLbl(lblHost) & ":"
    lblField(8).Caption = GetLbl(lblRemotePath) & ":"
    chkProxy.Caption = GetLbl(lblUseAccountProxy)
    lblField(10).Caption = GetLbl(lblPort) & ":"
    lblField(11).Caption = GetLbl(lblUser) & ":"
    lblField(12).Caption = GetLbl(lblPassword) & ":"
    lblField(13).Caption = GetLbl(lblLinkURL) & ":"
    lblTip(0).Caption = GetLbl(lblTipUpload)
    
    fraMedia.Caption = GetLbl(lblAddMediaInfo)
    lblMediaOpt.Caption = GetLbl(lblMediaOptions)
    optMedia(0).Caption = GetLbl(lblMediaManual)
    optMedia(1).Caption = GetLbl(lblMediaAutoTop)
    optMedia(2).Caption = GetLbl(lblMediaAutoBottom)
    chkMediaLink.Caption = GetLbl(lblMediaLink)
    lblField(20).Caption = GetLbl(lblMediaString)
    cmdMediaField(0).Caption = GetLbl(lblMediaTitle)
    cmdMediaField(1).Caption = GetLbl(lblMediaArtist)
    cmdMediaField(2).Caption = GetLbl(lblMediaAlbum)
    cmdMediaField(3).Caption = GetLbl(lblMediaDuration)
    lblPlaceholders.Caption = GetLbl(lblMediaCodes)
    lblTip(2).Caption = GetLbl(lblTipMedia)
    
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub

Private Sub optUpload_Click(Index As Integer)
    If Index = 0 Then
        txtHost.BackColor = vbWindowBackground
        txtPath.BackColor = vbWindowBackground
        txtPort.BackColor = vbWindowBackground
        txtUser.BackColor = vbWindowBackground
        txtPassword.BackColor = vbWindowBackground
        txtLink.BackColor = vbWindowBackground
        txtHost.Enabled = True
        txtPath.Enabled = True
        txtPort.Enabled = True
        txtUser.Enabled = True
        txtPassword.Enabled = True
        txtLink.Enabled = True
        chkProxy.Enabled = gAccount.UseProxy
    Else
        txtHost.BackColor = vbButtonFace
        txtPath.BackColor = vbButtonFace
        txtPort.BackColor = vbButtonFace
        txtUser.BackColor = vbButtonFace
        txtPassword.BackColor = vbButtonFace
        txtLink.BackColor = vbButtonFace
        txtHost.Enabled = False
        txtPath.Enabled = False
        txtPort.Enabled = False
        txtUser.Enabled = False
        txtPassword.Enabled = False
        txtLink.Enabled = False
        chkProxy.Enabled = False
    End If
End Sub
