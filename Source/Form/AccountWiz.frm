VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{ED442B9F-ADE2-11D4-B868-00606E3BC2C9}#1.0#0"; "ActiveCbo.ocx"
Begin VB.Form frmAccountWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Account Wizard"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AccountWiz.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWizard 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   6060
      TabIndex        =   51
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "&Next >"
      Height          =   375
      Index           =   1
      Left            =   4380
      TabIndex        =   49
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdWizard 
      Caption         =   "< &Back"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   50
      Top             =   4800
      Width           =   1335
   End
   Begin SizerOneLibCtl.TabOne tabWizard 
      Height          =   4635
      Left            =   0
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   0
      Width           =   7515
      _cx             =   13256
      _cy             =   8176
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
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "Tab&1|Tab&2|Tab&3|Tab&4|Tab&5|Tab&6|Tab&7"
      Align           =   1
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   0   'False
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   100
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4500
         Index           =   0
         Left            =   8130
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   15
         Width           =   7485
         _cx             =   13203
         _cy             =   7938
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
         Begin VB.CommandButton cmdImportAccounts 
            Height          =   330
            Left            =   4980
            Picture         =   "AccountWiz.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Import Accounts and Settings..."
            Top             =   1320
            Width           =   375
         End
         Begin VB.ComboBox cboPing 
            Height          =   315
            ItemData        =   "AccountWiz.frx":03B5
            Left            =   3510
            List            =   "AccountWiz.frx":03C2
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2655
            Width           =   1395
         End
         Begin VB.CheckBox chkProxy 
            Appearance      =   0  'Flat
            Caption         =   "&Use Proxy Server"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2400
            TabIndex        =   9
            Top             =   3480
            Width           =   2385
         End
         Begin VB.CheckBox chkWeblogs 
            Appearance      =   0  'Flat
            Caption         =   "Ping:"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2400
            TabIndex        =   7
            Top             =   2700
            Width           =   1665
         End
         Begin VB.TextBox txtAlias 
            Height          =   315
            Left            =   2400
            MaxLength       =   25
            TabIndex        =   6
            Top             =   1815
            Width           =   2505
         End
         Begin rdActiveCombo.ActiveCombo cboCMS 
            Height          =   330
            Left            =   2400
            TabIndex        =   3
            Top             =   1320
            Width           =   2505
            _ExtentX        =   4419
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
         Begin VB.Label lblMessages 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Check if you need a Proxy to access internet web pages. The settings will be configured on a next step."
            ForeColor       =   &H80000011&
            Height          =   615
            Index           =   6
            Left            =   2640
            TabIndex        =   87
            Top             =   3780
            Width           =   4320
         End
         Begin VB.Label lblMessages 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Check if you want w.bloggar to notify one of the listed sites when you publish a new post at your blog."
            ForeColor       =   &H80000011&
            Height          =   465
            Index           =   5
            Left            =   2640
            TabIndex        =   86
            Top             =   2970
            Width           =   4320
         End
         Begin VB.Label lblMessages 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter here a name to identify your account at w.bloggar status bar. Ex. My Personal Blog"
            ForeColor       =   &H80000011&
            Height          =   435
            Index           =   4
            Left            =   2400
            TabIndex        =   85
            Top             =   2160
            Width           =   3975
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "A&ccount Alias:"
            Height          =   195
            Index           =   8
            Left            =   900
            TabIndex        =   5
            Top             =   1860
            Width           =   1020
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Blog Tool:"
            Height          =   195
            Index           =   7
            Left            =   900
            TabIndex        =   2
            Top             =   1380
            Width           =   705
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   0
            X1              =   0
            X2              =   7500
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Image imgHeader 
            Appearance      =   0  'Flat
            Height          =   1050
            Index           =   0
            Left            =   6600
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblMessages 
            BackStyle       =   0  'Transparent
            Caption         =   "In which tool or service your blog was created?"
            Height          =   495
            Index           =   3
            Left            =   900
            TabIndex        =   60
            Top             =   480
            Width           =   5640
         End
         Begin VB.Label lblPageTit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Content Management System"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   420
            TabIndex        =   59
            Top             =   180
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000005&
            Height          =   1050
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   7515
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   4500
         Left            =   15
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   15
         Width           =   7485
         _cx             =   13203
         _cy             =   7938
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
         BackColor       =   -2147483643
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
         Begin VB.OptionButton optHaveBlog 
            BackColor       =   &H80000005&
            Caption         =   "No, please help me to create one"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   1
            Top             =   3360
            Width           =   4635
         End
         Begin VB.OptionButton optHaveBlog 
            BackColor       =   &H80000005&
            Caption         =   "Yes, I want to add it as a new account"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   0
            Top             =   2580
            Value           =   -1  'True
            Width           =   4635
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   4935
            Left            =   0
            Picture         =   "AccountWiz.frx":03E7
            ScaleHeight     =   4935
            ScaleWidth      =   2295
            TabIndex        =   54
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               HasDC           =   0   'False
               Height          =   915
               Left            =   1290
               Picture         =   "AccountWiz.frx":D27B
               ScaleHeight     =   915
               ScaleWidth      =   915
               TabIndex        =   76
               Top             =   255
               Width           =   915
            End
         End
         Begin VB.Label lblMessages 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "This option will guide you to subscribe to a supported blog provider"
            ForeColor       =   &H80000011&
            Height          =   435
            Index           =   2
            Left            =   2880
            TabIndex        =   71
            Top             =   3660
            Width           =   3855
         End
         Begin VB.Label lblMessages 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "This option guides you to configure w.bloggar to post to your existing blog."
            ForeColor       =   &H80000011&
            Height          =   435
            Index           =   1
            Left            =   2880
            TabIndex        =   70
            Top             =   2880
            Width           =   3855
         End
         Begin VB.Label lblPageTit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Do you already have a blog?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3630
            TabIndex        =   69
            Top             =   2220
            Width           =   2415
         End
         Begin VB.Label lblMessages 
            BackStyle       =   0  'Transparent
            Caption         =   $"AccountWiz.frx":D757
            Height          =   1245
            Index           =   0
            Left            =   2700
            TabIndex        =   56
            Top             =   1020
            Width           =   4335
         End
         Begin VB.Label lblPageTit 
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to the 'Add Account Wizard'"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   0
            Left            =   2460
            TabIndex        =   55
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   4830
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4500
         Index           =   1
         Left            =   9330
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   15
         Width           =   7485
         _cx             =   13203
         _cy             =   7938
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
         Begin VB.TextBox txtUser 
            Height          =   315
            Left            =   2040
            TabIndex        =   44
            Top             =   1560
            Width           =   1635
         End
         Begin VB.CheckBox chkSavePwd 
            Appearance      =   0  'Flat
            Caption         =   "&Save Password"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            TabIndex        =   47
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   46
            Top             =   1980
            Width           =   1635
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&User:"
            Height          =   195
            Index           =   4
            Left            =   900
            TabIndex        =   43
            Top             =   1620
            Width           =   390
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Password:"
            Height          =   195
            Index           =   3
            Left            =   900
            TabIndex        =   45
            Top             =   2040
            Width           =   750
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   0
            X2              =   7500
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Image imgHeader 
            Appearance      =   0  'Flat
            Height          =   1050
            Index           =   4
            Left            =   6600
            Picture         =   "AccountWiz.frx":D82B
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblPageTit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account User and Password"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   420
            TabIndex        =   63
            Top             =   180
            Width           =   2340
         End
         Begin VB.Label lblMessages 
            BackStyle       =   0  'Transparent
            Caption         =   "What are your account login settings?"
            Height          =   495
            Index           =   12
            Left            =   900
            TabIndex        =   62
            Top             =   480
            Width           =   5640
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000005&
            Height          =   1050
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   7515
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4500
         Index           =   2
         Left            =   8730
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   15
         Width           =   7485
         _cx             =   13203
         _cy             =   7938
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
         Begin VB.CheckBox chkUTF8 
            Appearance      =   0  'Flat
            Caption         =   "UTF-8"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5175
            TabIndex        =   32
            Top             =   3390
            Value           =   1  'Checked
            Width           =   780
         End
         Begin VB.TextBox txtPage 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2100
            MaxLength       =   255
            TabIndex        =   28
            Top             =   2400
            Width           =   3855
         End
         Begin VB.TextBox txtHost 
            Height          =   315
            Left            =   2100
            MaxLength       =   255
            TabIndex        =   26
            Top             =   1380
            Width           =   3855
         End
         Begin VB.TextBox txtPort 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2100
            MaxLength       =   5
            TabIndex        =   30
            Top             =   3345
            Width           =   720
         End
         Begin VB.CheckBox chkSecure 
            Appearance      =   0  'Flat
            Caption         =   "HTTPS"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3060
            TabIndex        =   31
            Top             =   3360
            Width           =   825
         End
         Begin VB.Label lblMessages 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter here the full path to the API endpoint of your blog tool. Start with slash (/) Ex: /xmlrpc.php"
            ForeColor       =   &H80000011&
            Height          =   435
            Index           =   10
            Left            =   2100
            TabIndex        =   89
            Top             =   2760
            Width           =   4680
         End
         Begin VB.Label lblMessages 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter here the host address without the protocol prefix (http://) and the ending slash. Ex: www.yoursite.com"
            ForeColor       =   &H80000011&
            Height          =   435
            Index           =   9
            Left            =   2100
            TabIndex        =   88
            Top             =   1740
            Width           =   4680
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "P&ath:"
            Height          =   195
            Index           =   1
            Left            =   1140
            TabIndex        =   27
            Top             =   2460
            Width           =   390
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Host:"
            Height          =   195
            Index           =   0
            Left            =   1140
            TabIndex        =   25
            Top             =   1440
            Width           =   390
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Por&t:"
            Height          =   195
            Index           =   2
            Left            =   1140
            TabIndex        =   29
            Top             =   3375
            Width           =   360
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   2
            X1              =   0
            X2              =   7500
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lblPageTit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Connection Settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   420
            TabIndex        =   67
            Top             =   180
            Width           =   2430
         End
         Begin VB.Label lblMessages 
            BackStyle       =   0  'Transparent
            Caption         =   "Where your blog tool is hosted?"
            Height          =   495
            Index           =   8
            Left            =   900
            TabIndex        =   66
            Top             =   480
            Width           =   5640
         End
         Begin VB.Image imgHeader 
            Appearance      =   0  'Flat
            Height          =   1050
            Index           =   2
            Left            =   6600
            Top             =   0
            Width           =   900
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000005&
            Height          =   1050
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Width           =   7515
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4500
         Index           =   3
         Left            =   9030
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   15
         Width           =   7485
         _cx             =   13203
         _cy             =   7938
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
         Begin VB.TextBox txtProxyPort 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2250
            MaxLength       =   5
            TabIndex        =   38
            Top             =   2535
            Width           =   720
         End
         Begin VB.TextBox txtProxyServer 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2250
            MaxLength       =   255
            TabIndex        =   36
            Top             =   2130
            Width           =   2400
         End
         Begin VB.OptionButton optProxy 
            Appearance      =   0  'Flat
            Caption         =   "Use Internet Explorer Proxy Settings"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   900
            TabIndex        =   33
            Top             =   1485
            Value           =   -1  'True
            Width           =   3975
         End
         Begin VB.OptionButton optProxy 
            Appearance      =   0  'Flat
            Caption         =   "Use the following Proxy Settings:"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   900
            TabIndex        =   34
            Top             =   1785
            Width           =   3975
         End
         Begin VB.TextBox txtProxyUser 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2250
            MaxLength       =   255
            TabIndex        =   40
            Top             =   2940
            Width           =   1485
         End
         Begin VB.TextBox txtProxyPassword 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2250
            MaxLength       =   255
            PasswordChar    =   "*"
            TabIndex        =   42
            Top             =   3345
            Width           =   1485
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Por&t:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   6
            Left            =   975
            TabIndex        =   37
            Top             =   2565
            Width           =   360
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Address:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   5
            Left            =   975
            TabIndex        =   35
            Top             =   2160
            Width           =   645
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&User:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   14
            Left            =   975
            TabIndex        =   39
            Top             =   2970
            Width           =   390
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "&Password:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   15
            Left            =   975
            TabIndex        =   41
            Top             =   3375
            Width           =   750
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   3
            X1              =   0
            X2              =   7500
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Image imgHeader 
            Appearance      =   0  'Flat
            Height          =   1050
            Index           =   3
            Left            =   6600
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblMessages 
            BackStyle       =   0  'Transparent
            Caption         =   "What are the proxy server settings?"
            Height          =   495
            Index           =   11
            Left            =   900
            TabIndex        =   74
            Top             =   480
            Width           =   5640
         End
         Begin VB.Label lblPageTit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Proxy Server"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   420
            TabIndex        =   73
            Top             =   180
            Width           =   1845
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            Height          =   1050
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   7515
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4500
         Index           =   4
         Left            =   8430
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   15
         Width           =   7485
         _cx             =   13203
         _cy             =   7938
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
         Begin VB.TextBox txtMoreTag 
            Height          =   315
            Index           =   1
            Left            =   3600
            TabIndex        =   24
            Text            =   "</more_text>"
            Top             =   3480
            Width           =   1110
         End
         Begin VB.TextBox txtMoreTag 
            Height          =   315
            Index           =   0
            Left            =   2415
            TabIndex        =   23
            Text            =   "<more_text>"
            Top             =   3480
            Width           =   1110
         End
         Begin VB.TextBox txtCategTag 
            Height          =   315
            Index           =   1
            Left            =   3600
            TabIndex        =   21
            Text            =   "</category>"
            Top             =   3075
            Width           =   1110
         End
         Begin VB.TextBox txtCategTag 
            Height          =   315
            Index           =   0
            Left            =   2415
            TabIndex        =   20
            Text            =   "<category>"
            Top             =   3075
            Width           =   1110
         End
         Begin VB.TextBox txtTitleTag 
            Height          =   315
            Index           =   1
            Left            =   3600
            TabIndex        =   18
            Text            =   "</title>"
            Top             =   2670
            Width           =   1110
         End
         Begin VB.TextBox txtTitleTag 
            Height          =   315
            Index           =   0
            Left            =   2415
            TabIndex        =   17
            Text            =   "<title>"
            Top             =   2670
            Width           =   1110
         End
         Begin VB.ComboBox cboTemplAPI 
            Height          =   315
            Left            =   2415
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2250
            Width           =   2310
         End
         Begin VB.ComboBox cboCategAPI 
            Height          =   315
            Left            =   2415
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1845
            Width           =   2310
         End
         Begin VB.ComboBox cboPostAPI 
            Height          =   315
            Left            =   2415
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1440
            Width           =   2310
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "More Text Tags:"
            Height          =   195
            Index           =   16
            Left            =   900
            TabIndex        =   22
            Top             =   3540
            Width           =   1185
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Category Tags:"
            Height          =   195
            Index           =   13
            Left            =   900
            TabIndex        =   19
            Top             =   3135
            Width           =   1125
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Title Tags:"
            Height          =   195
            Index           =   12
            Left            =   900
            TabIndex        =   16
            Top             =   2760
            Width           =   750
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Templates:"
            Height          =   195
            Index           =   11
            Left            =   900
            TabIndex        =   14
            Top             =   2310
            Width           =   795
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Categories:"
            Height          =   195
            Index           =   10
            Left            =   900
            TabIndex        =   12
            Top             =   1905
            Width           =   840
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Posts:"
            Height          =   195
            Index           =   9
            Left            =   900
            TabIndex        =   10
            Top             =   1500
            Width           =   450
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   4
            X1              =   0
            X2              =   7500
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lblPageTit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Blog Tool Settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   420
            TabIndex        =   79
            Top             =   180
            Width           =   2205
         End
         Begin VB.Label lblMessages 
            BackStyle       =   0  'Transparent
            Caption         =   "What API Methods does your blog support?"
            Height          =   495
            Index           =   7
            Left            =   900
            TabIndex        =   78
            Top             =   480
            Width           =   5640
         End
         Begin VB.Image imgHeader 
            Appearance      =   0  'Flat
            Height          =   1050
            Index           =   1
            Left            =   6600
            Top             =   0
            Width           =   900
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000005&
            Height          =   1050
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   7515
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   4500
         Index           =   5
         Left            =   9630
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   15
         Width           =   7485
         _cx             =   13203
         _cy             =   7938
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
         Begin SHDocVwCtl.WebBrowser webSubscribe 
            Height          =   3135
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   7215
            ExtentX         =   12726
            ExtentY         =   5530
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   0
            X2              =   7500
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label lblPageTit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subscribe to a Blog Provider"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   420
            TabIndex        =   83
            Top             =   180
            Width           =   2370
         End
         Begin VB.Label lblMessages 
            BackStyle       =   0  'Transparent
            Caption         =   "Choose a service to subscribe and get your blog!"
            Height          =   495
            Index           =   13
            Left            =   900
            TabIndex        =   82
            Top             =   480
            Width           =   5640
         End
         Begin VB.Image imgHeader 
            Appearance      =   0  'Flat
            Height          =   1050
            Index           =   5
            Left            =   6600
            Top             =   0
            Width           =   900
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000005&
            Height          =   1050
            Left            =   0
            TabIndex        =   84
            Top             =   0
            Width           =   7515
         End
      End
   End
   Begin VB.Image imgGhost 
      Appearance      =   0  'Flat
      Height          =   1050
      Index           =   1
      Left            =   1500
      Picture         =   "AccountWiz.frx":EB9E
      Top             =   4740
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgGhost 
      Appearance      =   0  'Flat
      Height          =   1050
      Index           =   0
      Left            =   540
      Picture         =   "AccountWiz.frx":1007C
      Top             =   4740
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgCustom 
      Height          =   240
      Left            =   240
      Picture         =   "AccountWiz.frx":116BB
      Top             =   4860
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   7500
      Y1              =   4635
      Y2              =   4635
   End
End
Attribute VB_Name = "frmAccountWiz"
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
'Private Variables
Private bolService As Boolean
Private intNextTab As Integer
Private intSaveAcc As Integer
Private strSavePwd As String

'Private Constants
Private Enum enumTabs
    TAB_WELCOME
    TAB_CMS
    TAB_CUSTOM
    TAB_CONNECTION
    TAB_PROXY
    TAB_LOGIN
    TAB_SUBSCRIBE
End Enum

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
    Else
        bolService = False
    End If
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

Private Sub cmdImportAccounts_Click()
    If frmPost.ImportSettings() Then
        Unload Me
        Unload frmLogin
    End If
End Sub

Private Sub cmdWizard_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim intTab As Integer
    intTab = tabWizard.CurrTab
    Select Case Index
    Case 0 'Back
        If intTab > 0 Then
            tabWizard.CurrTab = tabWizard.TabData(intTab)
        End If
    Case 1 'Next or Finish
        Select Case intTab
        Case TAB_WELCOME
            If optHaveBlog(0).Value Then
                intNextTab = TAB_CMS
            Else
                intNextTab = TAB_SUBSCRIBE
            End If
        Case TAB_CMS
            If Len(Trim(txtAlias.Text)) = 0 Then
                MsgBox GetMsg(msgEnterAlias), vbInformation
                txtAlias.SetFocus
                Exit Sub
            End If
            If cboCMS.ListIndex < cboCMS.ListCount - 1 Then
                If bolService Then
                    If chkProxy.Value = vbChecked Then
                        intNextTab = TAB_PROXY
                    Else
                        intNextTab = TAB_LOGIN
                    End If
                Else
                    intNextTab = TAB_CONNECTION
                End If
            Else
                txtHost.Text = ""
                txtPage.Text = ""
                chkSecure.Value = False
                intNextTab = TAB_CUSTOM
            End If
        Case TAB_CUSTOM
            If bolService Then
                If chkProxy.Value = vbChecked Then
                    intNextTab = TAB_PROXY
                Else
                    intNextTab = TAB_LOGIN
                End If
            Else
                intNextTab = TAB_CONNECTION
            End If
        Case TAB_CONNECTION
            If Trim(txtHost.Text) = "" Then
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
            End If
            If chkProxy.Value = vbChecked Then
                intNextTab = TAB_PROXY
            Else
                intNextTab = TAB_LOGIN
            End If
        Case TAB_PROXY
            If optProxy(1).Value Then
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
            intNextTab = TAB_LOGIN
        Case TAB_LOGIN
            If Trim(txtUser.Text) = "" Then
                MsgBox GetMsg(msgEnterUser), vbInformation
                txtUser.SetFocus
                Exit Sub
            ElseIf Trim(txtPassword.Text) = "" Then
                MsgBox GetMsg(msgEnterPassword), vbInformation
                txtPassword.SetFocus
                Exit Sub
            End If
            'Set Account Object
            gAccount.User = txtUser.Text
            gAccount.Password = txtPassword.Text
            gAccount.SavePassword = chkSavePwd.Value
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
            gAccount.Timeout = 30
            gAccount.UTF8 = chkUTF8.Value
            If chkProxy.Value <> vbChecked Then
                gAccount.UseProxy = 0
                gAccount.ProxyServer = ""
                gAccount.ProxyPort = 0
                gAccount.ProxyUser = ""
                gAccount.ProxyPassword = ""
            ElseIf optProxy(0).Value Then
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
            Else
                LoadCMS
            End If
            'Check the Account Connection
            If LoadBlogs(False) Then
                SaveAccount
                frmPost.Account = gAccount.Alias
            Else
                gAccount.User = ""
                gAccount.Password = ""
                Exit Sub
            End If
            Unload Me
            Unload frmLogin
            Exit Sub
        Case TAB_SUBSCRIBE
            Unload Me
            Exit Sub
        End Select
        tabWizard.TabData(intNextTab) = intTab
        tabWizard.CurrTab = intNextTab
    Case 2 'Cancel
        RestoreAccount
        Unload Me
    End Select
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim strAux As String
    tabWizard.CurrTab = 0
    tabWizard.Position = tpTop
    tabWizard.TabHeight = 1
    cmdWizard(0).Enabled = False
    LoadCombos
    'Load Blog Service List
    strAux = GetBinaryFile(App.Path & "\CMS\CMS.htm")
    strAux = Replace(strAux, "<i8>Name</i8>", GetLbl(lblName))
    strAux = Replace(strAux, "<i8>Features</i8>", GetLbl(lblFeatures))
    strAux = Replace(strAux, "<i8>Subscribe</i8>", GetLbl(lblSubscribe))
    strAux = Replace(strAux, "<i8>Language</i8>", GetLbl(lblLanguage))
    strAux = Replace(strAux, "<i8>Free</i8>", GetLbl(lblFree))
    strAux = Replace(strAux, "<i8>Trial</i8>", GetLbl(lblTrial))
    SaveBinaryFile App.Path & "\CMS\CMSTMP.htm", strAux
    webSubscribe.Navigate2 App.Path & "\CMS\CMSTMP.htm"
    'Fill Image Headers
    imgHeader(0).Picture = imgGhost(0).Picture
    imgHeader(1).Picture = imgGhost(0).Picture
    imgHeader(2).Picture = imgGhost(1).Picture
    imgHeader(3).Picture = imgGhost(1).Picture
    imgHeader(5).Picture = imgGhost(0).Picture
    'Save Current Account Data
    intSaveAcc = gAccount.Current
    strSavePwd = gAccount.Password
    'Get Next Account ID
    gAccount.Current = NextAccount()
    LoadAccount
    LocalizeForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then RestoreAccount
End Sub

Private Sub LoadCombos()
On Error Resume Next
    Call LoadCMSCombo(cboCMS, imgCustom.Picture, True)
    Call LoadPingCombo(cboPing, True)
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
    cboPing.ListIndex = 0
    cboPing.ListIndex = 0
    cboPing.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Kill App.Path & "\CMS\CMSTMP.htm"
End Sub

Private Sub optProxy_Click(Index As Integer)
    If optProxy(1).Value Then
        lblField(5).Enabled = True
        lblField(6).Enabled = True
        lblField(14).Enabled = True
        lblField(15).Enabled = True
        txtProxyServer.Enabled = True
        txtProxyPort.Enabled = True
        txtProxyUser.Enabled = True
        txtProxyPassword.Enabled = True
        txtProxyServer.BackColor = vbWindowBackground
        txtProxyPort.BackColor = vbWindowBackground
        txtProxyUser.BackColor = vbWindowBackground
        txtProxyPassword.BackColor = vbWindowBackground
    Else
        lblField(5).Enabled = False
        lblField(6).Enabled = False
        lblField(14).Enabled = False
        lblField(15).Enabled = False
        txtProxyServer.Enabled = False
        txtProxyPort.Enabled = False
        txtProxyUser.Enabled = False
        txtProxyPassword.Enabled = False
        txtProxyServer.BackColor = vbButtonFace
        txtProxyPort.BackColor = vbButtonFace
        txtProxyUser.BackColor = vbButtonFace
        txtProxyPassword.BackColor = vbButtonFace
    End If
End Sub

Private Sub tabWizard_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If NewTab = 0 Then
        cmdWizard(0).Enabled = False
    Else
        cmdWizard(0).Enabled = True
    End If
    If NewTab = TAB_LOGIN Then
        cmdWizard(1).Caption = "&" & GetLbl(lblFinish)
    Else
        cmdWizard(1).Caption = "&" & GetLbl(lblNext) & " >"
    End If
    If NewTab = TAB_SUBSCRIBE Then
        cmdWizard(1).Enabled = False
    Else
        cmdWizard(1).Enabled = True
    End If
    Select Case NewTab
    Case TAB_CMS
        cboCMS.SetFocus
    Case TAB_CUSTOM
        cboPostAPI.SetFocus
    Case TAB_CONNECTION
        txtHost.SetFocus
    Case TAB_LOGIN
        txtUser.SetFocus
    End Select
End Sub

Private Sub RestoreAccount()
    'Restore Current Account Settings
    gAccount.Current = intSaveAcc
    LoadAccount
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    Call objXMLReg.SaveSetting(App.Title, "Settings", "Account", Format(gAccount.Current))
    gAccount.Password = strSavePwd
    Set objXMLReg = Nothing
End Sub

Private Function NextAccount() As Integer
On Error GoTo ErrorHandler
Dim strReg As String, strUser As String
Dim intNext As Integer, a As Integer
    intNext = -1
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    For a = 0 To 99
        strReg = "Accounts/a" & Format(a, "00")
        strUser = objXMLReg.GetSetting(App.Title, strReg, "User", "*")
        If strUser = "*" Then Exit For
        If CBool(objXMLReg.GetSetting(App.Title, strReg, "Deleted", "0")) Then
            intNext = a
            Exit For
        End If
    Next
    Set objXMLReg = Nothing
    If intNext < 0 Then intNext = a
    NextAccount = intNext
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Function

Private Sub webSubscribe_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If (Left(URL, 4) = "http" Or Left(URL, 4) = "ftp:") And URL <> "http:///" Then
        OpenWebPage URL
        Cancel = True
    End If
End Sub

Private Sub OpenWebPage(ByVal strWebPage As String)
On Error Resume Next
    #If compIE Then
        If gSettings.DefaultBrowser Then
            Call ShellExecute(Me.hwnd, "open", strWebPage, vbNullString, CurDir$, SW_SHOW)
        Else
            webSubscribe.Navigate strWebPage, 5
        End If
    #Else
        Call ShellExecute(Me.hwnd, "open", strWebPage, vbNullString, CurDir$, SW_SHOW)
    #End If
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
Dim m As Integer
    Me.Caption = GetLbl(lblAddAccWizard)
    lblPageTit(0).Caption = GetLbl(lblWelcomeTo) & " '" & GetLbl(lblAddAccWizard) & "'"
    lblPageTit(1).Caption = GetLbl(lblDoYouHaveBlog)
    lblPageTit(2).Caption = GetLbl(lblCMS)
    lblPageTit(3).Caption = GetLbl(lblCustBlogSettings)
    lblPageTit(4).Caption = GetLbl(lblAccConnSettings)
    lblPageTit(5).Caption = GetLbl(lblAccProxyServer)
    lblPageTit(6).Caption = GetLbl(lblAccUsrPwd)
    lblPageTit(7).Caption = GetLbl(lblSubscribeProvider)
    lblField(0).Caption = GetLbl(lblHost) & ":"
    lblField(1).Caption = GetLbl(lblPage) & ":"
    lblField(2).Caption = GetLbl(lblPort) & ":"
    lblField(3).Caption = GetLbl(lblPassword) & ":"
    lblField(4).Caption = GetLbl(lblUser) & ":"
    lblField(5).Caption = GetLbl(lblTimeout) & ":"
    lblField(6).Caption = GetLbl(lblSeconds)
    lblField(5).Caption = GetLbl(lblAddress) & ":"
    lblField(6).Caption = GetLbl(lblPort) & ":"
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
    For m = 0 To 13
        lblMessages(m).Caption = GetMsg(msgWizMsg00 + m)
    Next
    cmdImportAccounts.ToolTipText = frmPost.acbMain.Bands("bndPopFile").Tools("miImportSettings").Caption
    optHaveBlog(0).Caption = GetLbl(lblYesHaveBlog)
    optHaveBlog(1).Caption = GetLbl(lblNoHaveBlog)
    optProxy(0).Caption = GetLbl(lblIEProxy)
    optProxy(1).Caption = GetLbl(lblMyProxy)
    chkSavePwd.Caption = GetLbl(lblSavePassword)
    chkWeblogs.Caption = GetLbl(lblPing) & ":"
    chkProxy.Caption = GetLbl(lblUseProxy)
    cmdWizard(0).Caption = "< &" & GetLbl(lblBack)
    cmdWizard(1).Caption = "&" & GetLbl(lblNext) & " >"
    cmdWizard(2).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
