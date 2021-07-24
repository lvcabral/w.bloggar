VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{8E22FD0B-91ED-11D2-8865-EAF032485D5B}#1.4#0"; "ActiveForm.ocx"
Begin VB.Form frmPost 
   Caption         =   " :: w.bloggar :: "
   ClientHeight    =   6600
   ClientLeft      =   1935
   ClientTop       =   345
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Post.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "frmPost"
   LockControls    =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   9930
   Begin VB.Timer tmrToolBar 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   630
      Top             =   5910
   End
   Begin rdActiveForm.ActiveForm acfPost 
      Left            =   75
      Top             =   5925
      _ExtentX        =   794
      _ExtentY        =   688
      MinWidth        =   5000
      MinHeight       =   3500
      MinimizeToTray  =   -1  'True
      RestoreMode     =   3
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 acbMain 
      Align           =   1  'Align Top
      Height          =   5715
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9930
      _LayoutVersion  =   1
      _ExtentX        =   17515
      _ExtentY        =   10081
      _DataPath       =   ""
      Bands           =   "Post.frx":1CFA
      Begin VB.PictureBox picPost 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   4095
         Left            =   15
         ScaleHeight     =   4095
         ScaleWidth      =   8700
         TabIndex        =   2
         Top             =   1155
         Width           =   8700
         Begin SizerOneLibCtl.TabOne tabPost 
            Height          =   3510
            Left            =   60
            TabIndex        =   5
            Top             =   525
            Width           =   8565
            _cx             =   15108
            _cy             =   6191
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
            Caption         =   "Editor|More|Preview"
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
            BorderWidth     =   -60
            BoldCurrent     =   0   'False
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   270
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            Begin SizerOneLibCtl.ElasticOne pnlMore 
               Height          =   3030
               Left            =   9270
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   375
               Width           =   8355
               _cx             =   14737
               _cy             =   5345
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
               AutoSizeChildren=   8
               BorderWidth     =   0
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   1
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   1350
               TagPosition     =   0
               Style           =   0
               TagSplit        =   2
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   4
               GridCols        =   3
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"Post.frx":2C744
               Begin SizerOneLibCtl.ElasticOne ElasticOne2 
                  Height          =   735
                  Left            =   0
                  TabIndex        =   19
                  TabStop         =   0   'False
                  Top             =   1920
                  Width           =   1335
                  _cx             =   2355
                  _cy             =   1296
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
                  Begin VB.Label lblExcEntry 
                     AutoSize        =   -1  'True
                     Caption         =   "Excerpt:"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   20
                     Top             =   60
                     Width           =   615
                  End
               End
               Begin VB.TextBox txtKeywords 
                  Height          =   315
                  Left            =   1395
                  TabIndex        =   17
                  Tag             =   "Keywords:"
                  Top             =   2715
                  Width           =   6960
               End
               Begin SizerOneLibCtl.ElasticOne ElasticOne1 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   15
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   6885
                  _cx             =   12144
                  _cy             =   556
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
                  Begin VB.Label lblExtEntry 
                     AutoSize        =   -1  'True
                     Caption         =   "Extended Entry:"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   18
                     Top             =   60
                     Width           =   1185
                  End
               End
               Begin VB.CommandButton cmdMore 
                  Caption         =   "Advanced"
                  Height          =   315
                  Left            =   6945
                  TabIndex        =   13
                  Top             =   0
                  Width           =   1410
               End
               Begin wbloggar.HtmlEdit txtExcerpt 
                  Height          =   735
                  Left            =   1395
                  TabIndex        =   16
                  Top             =   1920
                  Width           =   6960
                  _ExtentX        =   12277
                  _ExtentY        =   1296
                  Entities        =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "Post.frx":2C7B3
               End
               Begin wbloggar.HtmlEdit txtMore 
                  Height          =   1485
                  Left            =   0
                  TabIndex        =   14
                  Top             =   375
                  Width           =   8355
                  _ExtentX        =   14737
                  _ExtentY        =   2619
                  Entities        =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "Post.frx":2C7CF
               End
            End
            Begin SizerOneLibCtl.ElasticOne pnlEditor 
               Height          =   3030
               Left            =   105
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   375
               Width           =   8355
               _cx             =   14737
               _cy             =   5345
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
               AutoSizeChildren=   8
               BorderWidth     =   0
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   1
               WordWrap        =   -1  'True
               MaxChildSize    =   0
               MinChildSize    =   0
               TagWidth        =   1000
               TagPosition     =   0
               Style           =   0
               TagSplit        =   2
               PicturePos      =   4
               CaptionStyle    =   0
               ResizeFonts     =   0   'False
               GridRows        =   2
               GridCols        =   5
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"Post.frx":2C7EB
               Begin VB.CommandButton cmdCategories 
                  Caption         =   "..."
                  Height          =   315
                  Left            =   8010
                  TabIndex        =   10
                  ToolTipText     =   "More Categories"
                  Top             =   0
                  Width           =   345
               End
               Begin VB.ComboBox cboPostCat 
                  Height          =   315
                  Left            =   6240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Tag             =   "Category:"
                  Top             =   0
                  Width           =   1710
               End
               Begin VB.TextBox txtPostTit 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   8
                  Tag             =   "Title:"
                  Top             =   0
                  Width           =   3990
               End
               Begin wbloggar.HtmlEdit txtPost 
                  Height          =   2655
                  Left            =   0
                  TabIndex        =   11
                  Top             =   375
                  Width           =   8355
                  _ExtentX        =   14737
                  _ExtentY        =   4683
                  Entities        =   0   'False
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "Post.frx":2C85B
               End
            End
            Begin SHDocVwCtl.WebBrowser webPreview 
               Height          =   3030
               Left            =   9570
               TabIndex        =   7
               Top             =   375
               Width           =   8355
               ExtentX         =   14737
               ExtentY         =   5345
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
               Location        =   "http:///"
            End
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "Post:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   285
            Left            =   480
            TabIndex        =   3
            Top             =   105
            UseMnemonic     =   0   'False
            Width           =   630
         End
         Begin VB.Image imgStatus 
            Height          =   240
            Left            =   150
            Top             =   135
            Width           =   240
         End
         Begin VB.Label lblBack 
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   375
            Left            =   60
            TabIndex        =   4
            Top             =   60
            UseMnemonic     =   0   'False
            Width           =   8565
         End
      End
   End
   Begin VB.TextBox txtCommand 
      Height          =   300
      Left            =   2100
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image imgMediaPlayer 
      Height          =   240
      Index           =   1
      Left            =   4020
      Picture         =   "Post.frx":2C877
      Top             =   6000
      Width           =   240
   End
   Begin VB.Image imgMediaPlayer 
      Height          =   240
      Index           =   0
      Left            =   3720
      Picture         =   "Post.frx":2CE01
      Top             =   6000
      Width           =   240
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   1440
      Picture         =   "Post.frx":2D38B
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon16 
      Height          =   240
      Left            =   1140
      Picture         =   "Post.frx":2D4E1
      Top             =   5940
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPost"
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
Private strCurrentFile As String
Private bolChanged As Boolean
Private strTagBegin As String
Private strTagEnd As String
Private strMusic As String
' Used to control the Insert key state
Private bolInsBody As Boolean
Private bolInsMore As Boolean
Private bolInsExpt As Boolean
' Used in multi-level undo/redo scheme
Private trapUndo As Boolean         ' Locks the document while
                                    ' undo/redo is performed
Private trapMove As Boolean

Private UndoStack As Collection
Private RedoStack As Collection

Private UndoMore As Collection
Private RedoMore As Collection

Private UndoExcerpt As Collection
Private RedoExcerpt As Collection

' Most Recently Used Files
Private objMRU    As clsMRUFileList

'Properties
Public PostData As New PostData
Public TextFilters As New Collection
Public objSpellCheck As clsSpell

Public Sub EditPost(udtPost As PostData)
On Error GoTo ErrorHandler
    ClearUndo
    strCurrentFile = ""
    DisplayMRU
    txtPostTit.Text = udtPost.Title
    If gAccount.GetPostsMethod = API_MT Or _
       gAccount.GetPostsMethod = API_METAWEBLOG Or _
       gAccount.Extended Or _
       gAccount.Excerpt Or _
       gAccount.Keywords Then
        txtPost.Text = udtPost.Text
        txtMore.Text = udtPost.More
        txtExcerpt.Text = udtPost.Excerpt
        txtKeywords.Text = udtPost.Keywords
        If gSettings.ColorizeCode Then
            txtPost.Colorize
            txtMore.Colorize
            txtExcerpt.Colorize
        End If
        If UndoMore.Count >= 1 Then UndoMore(UndoMore.Count).Text = txtMore.TextRTF
        If UndoExcerpt.Count >= 1 Then UndoExcerpt(UndoExcerpt.Count).Text = txtExcerpt.TextRTF
        'Advanced Post Options
        If gAccount.AdvancedOptions Then
            PostData.AllowComments = udtPost.AllowComments
            PostData.AllowPings = udtPost.AllowPings
            PostData.DateTime = udtPost.DateTime
            PostData.TextFilter = udtPost.TextFilter
            PostData.TrackBack = udtPost.TrackBack
        End If
    Else
        txtPost.Text = udtPost.Text
        If gSettings.ColorizeCode Then txtPost.Colorize
    End If
    PostData.Categories = ""
    If cboPostCat.Visible Then
        If cboPostCat.ListCount > 0 Then cboPostCat.ListIndex = 0
        Select Case gAccount.GetPostsMethod
        Case API_B2
            SearchItemData cboPostCat, Val(udtPost.Categories)
        Case API_METAWEBLOG
            If udtPost.Categories <> "" Then
                SearchComboBox cboPostCat, (Split(udtPost.Categories, vbTab)(0))
                PostData.Categories = udtPost.Categories
            End If
        Case API_MT
            If udtPost.Categories <> "" Then
                SearchItemData cboPostCat, Val(Split(udtPost.Categories, vbTab)(0))
                PostData.Categories = udtPost.Categories
            End If
        Case Else
            If udtPost.Categories <> "" Then
                SearchComboBox cboPostCat, udtPost.Categories
            End If
        End Select
    End If
    If UndoStack.Count >= 1 Then UndoStack(UndoStack.Count).Text = txtPost.TextRTF
    gPostID = udtPost.PostID
    Set PostData = udtPost
    imgStatus.Picture = acbMain.Tools("miPosts").GetPicture(0)
    lblStatus.Caption = GetLbl(lblPost) & ": " & gPostID
    Me.Message = GetLbl(lblPostedAt) & " " & udtPost.DateTime
    TemplateMode False
    tabPost.CurrTab = TAB_EDITOR
    strCurrentFile = ""
    bolChanged = False
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".EditPost"
End Sub

Public Sub TemplateMode(ByVal bolMode As Boolean)
On Error GoTo ErrorHandler
Static HTMLWasVisible As Boolean
Static InTemplateMode As Boolean
    'Set Menu
    acbMain.Tools("miPost").Visible = Not bolMode
    acbMain.Tools("miSaveTemplate").Visible = bolMode
    If gAccount.CMS = CMS_BLOGGER Or gAccount.CMS = CMS_BLOGGERPRO Then
        acbMain.Tools("miPublish").Visible = Not bolMode
        acbMain.Tools("miPublish").Enabled = Not bolMode
        acbMain.Tools("miPublishTemplate").Visible = bolMode
    Else
        acbMain.Tools("miPublishTemplate").Visible = False
        acbMain.Tools("miPublish").Visible = True
        acbMain.Tools("miPublish").Enabled = Not bolMode
    End If
    acbMain.Tools("miMultipost").Enabled = Not bolMode
    If gSettings.ShowHtmlBar Then
        If bolMode Then
            HTMLWasVisible = acbMain.Bands("bndHTML").Visible
            acbMain.Bands("bndHTML").Visible = True
        ElseIf acbMain.Bands("bndHTML").Visible And InTemplateMode Then
            acbMain.Bands("bndHTML").Visible = HTMLWasVisible
        End If
    End If
    acbMain.RecalcLayout
    'hide/show title and category fields
    If bolMode Or Not SupportsTitle() Then
        txtPostTit.Visible = False
    Else
        txtPostTit.Visible = True
    End If
    If pnlEditor.Grid(gsColWidth, 3) > 15 Then
        cboPostCat.Visible = Not bolMode
        cmdCategories.Visible = Not bolMode
    End If
    If bolMode Then
        txtPostTit.Text = ""
        If cboPostCat.Visible And cboPostCat.ListCount > 0 Then cboPostCat.ListIndex = 0
        pnlEditor.Grid(gsRowHeight, 0) = 15
    ElseIf Not SupportsTitle() Then
        pnlEditor.Grid(gsRowHeight, 0) = 15
    Else
        pnlEditor.Grid(gsRowHeight, 0) = 375
    End If
    'More Tab
    If bolMode Then
        tabPost.TabVisible(TAB_MORE) = False
    ElseIf gAccount.PostMethod = API_MT Or gAccount.MoreTextTag1 <> "" Or gAccount.MoreTextTag2 <> "" Then
        tabPost.TabVisible(TAB_MORE) = True
    End If
    InTemplateMode = bolMode
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".TemplateMode"
End Sub

Public Sub NewPost()
On Error Resume Next
    'Clear Editor Area
    txtPostTit.Text = ""
    If cboPostCat.ListCount > 0 Then cboPostCat.ListIndex = 0
    If gSettings.PostTemplate <> "" And FileExists(gSettings.PostTemplate) Then
        Call LoadPostFile(gSettings.PostTemplate, True)
    Else
        txtPost.Text = ""
    End If
    If txtPost.Text <> "" Then txtPost.Colorize
    txtMore.Text = ""
    txtExcerpt.Text = ""
    txtKeywords.Text = ""
    gPostID = ""
    strCurrentFile = ""
    DisplayMRU
    imgStatus.Picture = acbMain.Tools("miNew").GetPicture(0)
    lblStatus.Caption = GetLbl(lblPost) & ": " & GetMsg(msgNewPost)
    tabPost.CurrTab = TAB_EDITOR
    Me.Message = ""
    TemplateMode False
    If txtPostTit.Visible Then
        txtPostTit.SetFocus
    ElseIf txtPost.Visible Then
        txtPost.SetFocus
    End If
    Set PostData = New PostData
    bolChanged = False
End Sub

Private Function SavePostData(Optional ByVal strFilePath As String) As Boolean
On Error GoTo ErrorHandler
Dim bolResult As Boolean
    If strFilePath = "" Then
        Dim oFile As New FileDialog
        oFile.DialogTitle = acbMain.Tools("miSaveAs").ToolTipText
        oFile.Filter = GetMsg(msgFileFilter)
        oFile.Flags = cdlOverWritePrompt Or cdlPathMustExist Or cdlLongnames Or cdlHideReadOnly
        If SafeFileName(txtPostTit.Text) <> "" Then
            oFile.FileName = SafeFileName(txtPostTit.Text) & ".post"
        End If
        oFile.hWndParent = Me.hwnd
        If oFile.ShowSave() Then strFilePath = oFile.FileName
        Set oFile = Nothing
    End If
    If strFilePath <> "" Then
        PostData.Title = txtPostTit.Text
        PostData.Text = txtPost.Text
        PostData.More = txtMore.Text
        PostData.Excerpt = txtExcerpt.Text
        PostData.Keywords = txtKeywords.Text
        bolResult = PostData.SaveData(strFilePath)
        bolChanged = Not bolResult
        objMRU.AddFile strFilePath
        DisplayMRU
        If strCurrentFile <> strFilePath Then
            strCurrentFile = strFilePath
            If gPostID <> "" Then
                If gPostID = "main" Then
                    lblStatus.Caption = GetMsg(msgMainTemplate) & " - " & objMRU.CompressFileName(strFilePath)
                ElseIf gPostID = "archiveIndex" Then
                    lblStatus.Caption = GetMsg(msgArchiveTemplate) & " - " & objMRU.CompressFileName(strFilePath)
                Else
                    lblStatus.Caption = GetLbl(lblPost) & ": " & gPostID & " - " & objMRU.CompressFileName(strFilePath)
                End If
            Else
                lblStatus.Caption = GetLbl(lblPost) & ": " & GetMsg(msgNewPost) & " - " & objMRU.CompressFileName(strFilePath)
            End If
        End If
        SavePostData = bolResult
    End If
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".SavePostDate"
End Function

Private Sub ExportSettings()
On Error GoTo ErrorHandler
Dim bolResult As Boolean
Dim strFilePath As String
Dim oFile As New FileDialog
    oFile.DialogTitle = Replace(acbMain.Tools("miExportSettings").Caption, "...", "")
    oFile.Filter = GetMsg(msgXMLFileFilter)
    oFile.Flags = cdlPathMustExist Or cdlLongnames Or cdlHideReadOnly
    oFile.FileName = GetNamePart(XML_SETTINGS)
    oFile.hWndParent = Me.hwnd
    oFile.DefaultExt = "xml"
    If oFile.ShowSave() Then strFilePath = oFile.FileName
    Set oFile = Nothing
    If strFilePath <> "" Then
        Call ShellCopyFile(Me.hwnd, gAppDataPath & XML_SETTINGS, strFilePath)
    End If
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".ExportSettings"
End Sub

Public Function ImportSettings(Optional blnSilent As Boolean) As Boolean
On Error Resume Next
Dim oFile As New FileDialog, strAux As String
    oFile.DialogTitle = Replace(acbMain.Tools("miImportSettings").Caption, "...", "")
    oFile.Filter = GetMsg(msgXMLFileFilter)
    oFile.Flags = cdlFileMustExist Or cdlLongnames Or cdlHideReadOnly
    oFile.hWndParent = Me.hwnd
    oFile.ShowOpen
    If oFile.FileName <> "" Then
        Set objXMLReg = New XMLRegistry
        If objXMLReg.OpenXMLFile(oFile.FileName, False) Then
            strAux = objXMLReg.GetSetting(App.Title, "Settings", "Account", "*")
            If Err.Number <> 0 Or strAux = "*" Or Not IsNumeric(strAux) Then
                MsgBox GetMsg(msgInvalidSettings), vbExclamation
                GoTo ExitNow
            End If
        Else
            MsgBox GetMsg(msgErrOpenSettings), vbCritical
            GoTo ExitNow
        End If
        Call ShellCopyFile(Me.hwnd, oFile.FileName, gAppDataPath & XML_SETTINGS, True)
        Kill gAppDataPath & "\blogs" & Format(strAux, "00") & ".xml" ' Delete existing blog list
        If Not blnSilent Then MsgBox GetMsg(msgSettingsImported), vbInformation
        Unload Me
        ImportSettings = True
    End If
ExitNow:
    Set objXMLReg = Nothing
    Set oFile = Nothing
End Function

Private Sub InsertList(ByVal tag As String)
On Error GoTo ErrorHandler
Dim intAux As Integer, strAux As String
Dim objText As Object
Dim intTab As Integer
Dim bolColor As Boolean
Dim aLine() As String
Dim i As Integer
    If Me.ActiveControl.Name = "txtPostTit" Then
        Set objText = txtPostTit
        intTab = TAB_EDITOR
        bolColor = False
    ElseIf Me.ActiveControl.Name = "txtPost" Then
        Set objText = txtPost
        intTab = TAB_EDITOR
        bolColor = True
    ElseIf Me.ActiveControl.Name = "txtMore" Then
        Set objText = txtMore
        intTab = TAB_MORE
        bolColor = True
    ElseIf Me.ActiveControl.Name = "txtExcerpt" Then
        Set objText = txtExcerpt
        intTab = TAB_MORE
        bolColor = True
    ElseIf Me.ActiveControl.Name = "txtKeywords" Then
        Set objText = txtKeywords
        intTab = TAB_MORE
        bolColor = False
    Else
        Exit Sub
    End If
    trapMove = False
    If objText.SelLength > 0 And InStr(objText.SelText, vbLf) > 0 Then
        If InStr(objText.SelText, vbCrLf) > 0 Then
            aLine = Split(objText.SelText, vbCrLf)
        Else
            aLine = Split(objText.SelText, vbLf)
        End If
        intAux = objText.SelStart
        strAux = "<" & tag & ">"
        For i = 0 To UBound(aLine)
            strAux = strAux & "<li>" & aLine(i) & "</li>" & IIf(i = UBound(aLine), "", vbCrLf)
        Next
        strAux = strAux & "</" & tag & ">"
        objText.SelText = strAux
        If gSettings.ColorizeCode And bolColor Then
            objText.Colorize intAux, intAux + Len(strAux) + 1
        End If
    Else
        InsertTag "<" & tag & "><li>", "</li></" & tag & ">"
    End If
    tabPost.CurrTab = intTab
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".InsertList"
End Sub

Private Sub InsertTag(ByVal strBegin As String, ByVal strEnd As String)
On Error GoTo ErrorHandler
Dim intAux As Integer, strAux As String
Dim objText As Object
Dim intTab As Integer
Dim bolColor As Boolean
    If strBegin = "" And strEnd = "" Then Exit Sub
    If Me.ActiveControl.Name = "txtPostTit" Then
        Set objText = txtPostTit
        intTab = TAB_EDITOR
        bolColor = False
    ElseIf Me.ActiveControl.Name = "txtPost" Then
        Set objText = txtPost
        intTab = TAB_EDITOR
        bolColor = True
    ElseIf Me.ActiveControl.Name = "txtMore" Then
        Set objText = txtMore
        intTab = TAB_MORE
        bolColor = True
    ElseIf Me.ActiveControl.Name = "txtExcerpt" Then
        Set objText = txtExcerpt
        intTab = TAB_MORE
        bolColor = True
    ElseIf Me.ActiveControl.Name = "txtKeywords" Then
        Set objText = txtKeywords
        intTab = TAB_MORE
        bolColor = False
    Else
        Exit Sub
    End If
    trapMove = False
    If objText.SelLength > 0 Then
        intAux = objText.SelStart
        strAux = strBegin & objText.SelText & strEnd
        objText.SelText = strAux
        objText.SelStart = intAux + Len(strAux)
        If gSettings.ColorizeCode And bolColor Then
            objText.Colorize intAux, objText.SelStart
        End If
    Else
        intAux = objText.SelStart
        objText.SelText = strBegin & strEnd
        If Right(strBegin, 1) = ">" Or Right(strBegin, 1) = " " Then
            objText.SelStart = intAux + Len(strBegin)
        Else
            objText.SelStart = intAux + Len(strBegin & strEnd)
        End If
        If gSettings.ColorizeCode And bolColor Then
            objText.Colorize intAux, objText.SelStart + Len(strEnd)
        End If
    End If
    tabPost.CurrTab = intTab
    trapMove = True
    strTagBegin = strBegin
    strTagEnd = strEnd
    acbMain.Tools("miRepeatTag").Enabled = True
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".InsertTag"
End Sub

Private Sub UploadFile()
    frmUpload.Show vbModal, Me
    If frmUpload.HtmlTag = "I" Then
        frmImage.Show vbModal, Me
        If frmImage.HtmlTag <> "" Then InsertTag frmImage.HtmlTag, ""
        Unload frmImage
        Set frmImage = Nothing
    ElseIf frmUpload.HtmlTag = "L" Then
        frmLink.Show vbModal, Me
        If frmLink.HtmlTag <> "" Then InsertTag frmLink.HtmlTag, "</a>"
        Unload frmLink
        Set frmLink = Nothing
    ElseIf frmUpload.HtmlTag <> "" Then
        InsertTag frmUpload.HtmlTag, ""
    End If
    Unload frmUpload
    Set frmUpload = Nothing
End Sub

Private Function GetColor() As String
    objColor.hWndParent = Me.hwnd
    If objColor.ShowColor Then
        SaveColors
        GetColor = Rgb2Html(objColor.Color)
    End If
End Function

Private Sub OpenBlogPage(ByVal intBlog As Integer)
    OpenWebPage gBlogs(intBlog).URL
End Sub

Private Sub OpenCMSPage()
    OpenWebPage ReadINI("CMS-" & Format(gAccount.CMS, "00"), "URL", App.Path & "\CMS\CMS.ini")
End Sub

Public Sub OpenWebPage(ByVal strWebPage As String)
On Error Resume Next
    #If compIE Then
        If gSettings.DefaultBrowser Then
            Call ShellExecute(Me.hwnd, "open", strWebPage, vbNullString, CurDir$, SW_SHOW)
        Else
            webPreview.Navigate strWebPage, 5
        End If
    #Else
        Call ShellExecute(Me.hwnd, "open", strWebPage, vbNullString, CurDir$, SW_SHOW)
    #End If
End Sub

Private Sub OpenDocument(ByVal strDocPath As String)
On Error Resume Next
    #If compIE Then
        If gSettings.DefaultBrowser Then
            Call ShellExecute(Me.hwnd, "open", strDocPath, vbNullString, CurDir$, SW_SHOW)
        Else
            webPreview.Navigate "file://" & strDocPath, 5
        End If
    #Else
        Call ShellExecute(Me.hwnd, "open", strDocPath, vbNullString, CurDir$, SW_SHOW)
    #End If
End Sub

Private Sub OpenPost()
On Error Resume Next
Dim oFile As New FileDialog
    oFile.DialogTitle = acbMain.Tools("miOpen").ToolTipText
    oFile.Filter = GetMsg(msgFileFilter)
    oFile.Flags = cdlFileMustExist Or cdlLongnames Or cdlHideReadOnly
    oFile.hWndParent = Me.hwnd
    oFile.ShowOpen
    If oFile.FileName <> "" Then
        tabPost.CurrTab = TAB_EDITOR
        LoadPostFile oFile.FileName
    End If
    Set oFile = Nothing
End Sub

Private Sub ImportText()
On Error Resume Next
Dim oFile As New FileDialog
Dim strTxt As String
    oFile.DialogTitle = acbMain.Tools("miOpen").ToolTipText
    oFile.Filter = GetMsg(msgFileFilterText)
    oFile.Flags = cdlFileMustExist Or cdlLongnames Or cdlHideReadOnly
    oFile.hWndParent = Me.hwnd
    oFile.ShowOpen
    If oFile.FileName <> "" Then
        tabPost.CurrTab = TAB_EDITOR
        strTxt = GetBinaryFile(oFile.FileName)
        txtPost.SelText = strTxt
        bolChanged = True
    End If
    Set oFile = Nothing
End Sub

Private Sub CutText()
    If ActiveControl Is Nothing Then Exit Sub
    Select Case ActiveControl.Name
    Case "txtPost"
        tabPost.CurrTab = TAB_EDITOR
        ' Copy the selected text onto the Clipboard.
        Clipboard.Clear
        Clipboard.SetText txtPost.SelText
        ' Delete the selected text.
        txtPost.SelText = ""
    Case "txtPostTit", "cboPostCat", "txtMore", "txtExcerpt", "txtKeywords"
        Clipboard.Clear
        Clipboard.SetText ActiveControl.SelText
        ActiveControl.SelText = ""
    End Select
End Sub

Private Sub CopyText()
On Error Resume Next
    If ActiveControl Is Nothing Then Exit Sub
    Select Case ActiveControl.Name
    Case "txtPost"
        tabPost.CurrTab = TAB_EDITOR
        Clipboard.Clear
        Clipboard.SetText txtPost.SelText
    Case "txtPostTit", "cboPostCat", "txtMore", "txtExcerpt", "txtKeywords"
        Clipboard.Clear
        Clipboard.SetText ActiveControl.SelText
    End Select
End Sub

Private Sub PasteText(Optional bolNoRTF As Boolean)
' On Error Resume Next
Dim intAux As Integer, intLin As Integer, intFmt As Integer
    If ActiveControl Is Nothing Then Exit Sub
    Select Case ActiveControl.Name
    Case "txtPost", "txtMore", "txtExcerpt"
        trapMove = False
        'tabPost.CurrTab = TAB_EDITOR
        intAux = ActiveControl.SelStart
        If Clipboard.GetFormat(vbCFRTF) And Not bolNoRTF Then
            intFmt = vbCFRTF
            ActiveControl.SelRTF = RTF2HTML(Clipboard.GetText(vbCFRTF), "+CR" & IIf(gBlog.PreviewAutoBR, "-BR", ""))
        Else
            intFmt = vbCFText
            ActiveControl.SelText = Clipboard.GetText(vbCFText)
        End If
        intLin = ActiveControl.GetLineFirstCharIndex
        If gSettings.ColorizeCode Then ActiveControl.Colorize IIf(intLin > intAux, intAux, intLin), intAux + Len(Clipboard.GetText(intFmt)) + 1
        trapMove = True
    Case "txtPostTit", "cboPostCat", "txtKeywords"
        ActiveControl.SelText = Clipboard.GetText(vbCFText)
    End Select
End Sub

Private Sub SelectAll()
On Error Resume Next
    If ActiveControl Is Nothing Then Exit Sub
    Select Case ActiveControl.Name
    Case "txtPost"
        tabPost.CurrTab = TAB_EDITOR
        txtPost.SelStart = 0
        txtPost.SelLength = Len(txtPost.Text)
    Case "txtPostTit", "cboPostCat", "txtMore", "txtExcerpt", "txtKeywords"
        ActiveControl.SelStart = 0
        ActiveControl.SelLength = Len(ActiveControl.Text)
    End Select
End Sub

Private Sub Undo()
    If ActiveControl Is Nothing Then Exit Sub
    Select Case ActiveControl.Name
    Case "txtPost"
        tabPost.CurrTab = TAB_EDITOR
        UndoIt txtPost, UndoStack, RedoStack
    Case "txtMore"
        tabPost.CurrTab = TAB_MORE
        UndoIt txtMore, UndoMore, RedoMore
    Case "txtExcerpt"
        tabPost.CurrTab = TAB_MORE
        UndoIt txtExcerpt, UndoExcerpt, RedoExcerpt
    Case "txtPostTit"
        If SendMessage(txtPostTit.hwnd, EM_CANUNDO, 0, ByVal 0&) Then
            Call SendMessage(txtPostTit.hwnd, EM_UNDO, 0, ByVal 0&)
        End If
    End Select
End Sub

Private Sub UndoIt(txtTarget As Control, colUndo As Collection, colRedo As Collection)
Dim objElement As UndoRedo
    If colUndo.Count > 1 And trapUndo Then 'we can proceed
        trapUndo = False
        Set objElement = colUndo(colUndo.Count - 1)
        txtTarget.TextRTF = objElement.Text
        txtTarget.SelStart = objElement.SelStart
        colRedo.Add Item:=colUndo(colUndo.Count)
        colUndo.Remove colUndo.Count
    End If
    trapUndo = True
    'call the Selection Change event
    SelChangeEvent txtTarget, colUndo, colRedo
    txtTarget.SetFocus
End Sub

Private Sub Redo()
    If ActiveControl Is Nothing Then Exit Sub
    Select Case ActiveControl.Name
    Case "txtPost"
        tabPost.CurrTab = TAB_EDITOR
        RedoIt txtPost, UndoStack, RedoStack
    Case "txtMore"
        tabPost.CurrTab = TAB_MORE
        RedoIt txtMore, UndoMore, RedoMore
    Case "txtExcerpt"
        tabPost.CurrTab = TAB_MORE
        RedoIt txtExcerpt, UndoExcerpt, RedoExcerpt
    End Select
End Sub

Private Sub RedoIt(txtTarget As Control, colUndo As Collection, colRedo As Collection)
Dim objElement As UndoRedo
    If colRedo.Count > 0 And trapUndo Then
        trapUndo = False
        Set objElement = colRedo(colRedo.Count)
        txtTarget.TextRTF = objElement.Text
        txtTarget.SelStart = objElement.SelStart
        colUndo.Add Item:=objElement
        colRedo.Remove colRedo.Count
    End If
    trapUndo = True
    'call the Selection Change event
    SelChangeEvent txtTarget, colUndo, colRedo
    txtTarget.SetFocus
End Sub

Public Sub ClearUndo()
    Set UndoStack = New Collection
    Set RedoStack = New Collection
    Set UndoMore = New Collection
    Set RedoMore = New Collection
    Set UndoExcerpt = New Collection
    Set RedoExcerpt = New Collection
    acbMain.Tools("miUndo").Enabled = False
    acbMain.Tools("miRedo").Enabled = False
End Sub

Private Sub acbMain_BandOpen(ByVal Band As ActiveBar2LibraryCtl.Band, ByVal Cancel As ActiveBar2LibraryCtl.ReturnBool)
Dim objTool As ActiveBar2LibraryCtl.Tool, strMenu As String
    Select Case Band.Name
    Case "bndPopView"
        Band.Tools("miToolsBar").Checked = acbMain.Bands("bndTools").Visible
        Band.Tools("miFormatBar").Checked = acbMain.Bands("bndFormat").Visible
        Band.Tools("miHTMLBar").Checked = acbMain.Bands("bndHTML").Visible
        Band.Tools("miStatusBar").Checked = acbMain.Bands("bndStatus").Visible
    Case "bndPopTools"
        If gPostID <> "" _
           And gPostID <> "main" _
           And gPostID <> "archiveIndex" Then
            Band.Tools("miDelete").Enabled = True
            Band.Tools("miDelPublish").Enabled = True
        Else
            Band.Tools("miDelete").Enabled = False
            Band.Tools("miDelPublish").Enabled = False
        End If
    Case "bndPopPosts"
        strMenu = gSettings.PostMenu
        For Each objTool In Band.Tools
            If objTool.Name = "miLast" & strMenu Then
                objTool.Checked = True
            Else
                objTool.Checked = False
            End If
        Next
        If strMenu = "0" Then
            Band.Tools("miPostID").Checked = True
        End If
    Case "SysCustomize"
        Band.Tools(Band.Tools.Count - 1).Visible = False

    End Select
End Sub

Private Sub acbMain_ComboSelChange(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error Resume Next
Dim bolSave As Boolean
    If gPostID <> "" Then
        NewPost
        'SaveDraftPost
    ElseIf Me.Visible Then
        'Commented 4.00 - Erase the PostID to avoid wrong Blog
        'Call SaveSetting(REGISTRY_KEY, "Settings", "PostID", "")
        Me.Message = ""
    End If
    bolSave = bolChanged
    If gAccount.TemplateMethod Then
        acbMain.Tools("miTemplate").Enabled = gBlogs(Tool.CBListIndex).IsAdmin
    Else
        acbMain.Tools("miTemplate").Enabled = False
    End If
    acbMain.ApplyAll acbMain.Tools("miTemplate")
    acbMain.RecalcLayout
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    Call objXMLReg.SaveSetting(App.Title, "Accounts/a" & Format(gAccount.Current, "00"), "Blog", Format(Tool.CBListIndex))
    Set objXMLReg = Nothing
    LoadBlogSettings
    LoadCategories Not Me.Visible
    'More Tab
    If gAccount.MoreTab Then
        acbMain.Tools("miMoreText").Enabled = True
        tabPost.TabVisible(TAB_MORE) = True
        acbMain.Tools("miAdvanced").Enabled = gAccount.AdvancedOptions
        cmdMore.Visible = gAccount.AdvancedOptions
        If gAccount.Extended Then
            lblExtEntry.Enabled = True
            txtMore.BackColor = vbWindowBackground
            txtMore.Enabled = True
        Else
            lblExtEntry.Enabled = False
            txtMore.BackColor = vbButtonFace
            txtMore.Enabled = False
            txtMore.Text = ""
        End If
        If gAccount.Excerpt Then
            lblExcEntry.Enabled = True
            pnlMore.Grid(gsRowHeight, 2) = 975
            txtExcerpt.Visible = True
        Else
            lblExcEntry.Enabled = False
            pnlMore.Grid(gsRowHeight, 2) = 15
            txtExcerpt.Visible = False
            txtExcerpt.Text = ""
        End If
        If gAccount.Keywords Then
            pnlMore.Grid(gsRowHeight, 3) = 375
            txtKeywords.Visible = True
        Else
            pnlMore.Grid(gsRowHeight, 3) = 15
            txtKeywords.Visible = False
            txtKeywords.Text = ""
        End If
        If gAccount.Extended And gAccount.Excerpt And gAccount.Keywords Then
            pnlMore.Grid(gsRowSplitter, 1) = True
        Else
            pnlMore.Grid(gsRowSplitter, 1) = False
        End If
    Else
        acbMain.Tools("miMoreText").Enabled = False
        tabPost.TabVisible(TAB_MORE) = False
        txtMore.Text = ""
        txtKeywords.Text = ""
        txtExcerpt.Text = ""
    End If
    tabPost.CurrTab = TAB_EDITOR
    If txtPostTit.Visible Then
        txtPostTit.SetFocus
    ElseIf txtPost.Visible Then
        SetForegroundWindow Me.hwnd
        txtPost.SetFocus
    End If
    bolChanged = bolSave
End Sub

Private Sub acbMain_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrorHandler
Dim intPosts As Integer, strColor As String
Dim aCateg As Variant, strPost As String, strMore As String
Dim strExcerpt As String, strKeywords As String
    If InStr("miPosts*miLast1*miLast5*miLast10&miLast15*miLast20*miLastn*miPostID*miTemplate*miMainTemplate*miArchiveTemplate*miPost*miSaveTemplate*miPublish*miMultiPost*miDelete*miDelPublish*", Tool.Name & "*") > 0 And _
       gAccount.Password = "" Then
        frmLogin.cboAccount.Enabled = False
        frmLogin.Show vbModal, Me
        If gAccount.Password = "" Then Exit Sub
    End If
    Select Case Tool.Name
    Case "miNew"
        If Not CanContinue() Then Exit Sub
        NewPost
        Exit Sub
    Case "miOpen"
        If Not CanContinue() Then Exit Sub
        OpenPost
    Case "miImportText"
        ImportText
    Case "miFMRU1", "miFMRU2", "miFMRU3", "miFMRU4"
        If Not CanContinue() Then Exit Sub
        Call LoadPostFile(objMRU.file(Tool.TagVariant))
    Case "miSave"
        SavePostData strCurrentFile
    Case "miSaveAs"
        SavePostData
    Case "miExportSettings" '4.00
        ExportSettings
    Case "miImportSettings" '4.00
        ImportSettings
        Exit Sub
    Case "miPreview" '3.03
        tabPost.CurrTab = IIf(tabPost.CurrTab = TAB_PREVIEW, TAB_EDITOR, TAB_PREVIEW)
        Exit Sub
    Case "miMoreText" '3.03
        tabPost.CurrTab = IIf(tabPost.CurrTab = TAB_MORE, TAB_EDITOR, TAB_MORE)
        Exit Sub
    Case "miToolsBar" '4.00
        acbMain.Bands("bndTools").Visible = Not acbMain.Bands("bndTools").Visible
        acbMain.RecalcLayout
    Case "miFormatBar" '4.00
        acbMain.Bands("bndFormat").Visible = Not acbMain.Bands("bndFormat").Visible
        acbMain.RecalcLayout
    Case "miHTMLBar" '4.00
        acbMain.Bands("bndHTML").Visible = Not acbMain.Bands("bndHTML").Visible
        acbMain.RecalcLayout
    Case "miStatusBar" '4.00
        acbMain.Bands("bndStatus").Visible = Not acbMain.Bands("bndStatus").Visible
        acbMain.RecalcLayout
    Case "miTSpelling"
        If objSpellCheck Is Nothing Then
            Set objSpellCheck = New clsSpell
            DoEvents
        End If
        Call CheckSpelling
    Case "miUpload"
        If gBlog.APIUpload Or Trim(gBlog.FTPHost) <> "" Then
            UploadFile
        Else
            If MsgBox(GetMsg(msgFTPSettings), vbInformation + vbYesNo) = vbYes Then
                frmBlog.tabIndex.CurrTab = enuTabUpload
                frmBlog.Show vbModal, Me
                If Trim(gBlog.FTPHost) <> "" Then UploadFile
            End If
        End If
    Case "miCustomEdit"
        frmCustom.ShowForm 1
    Case "miCustomF1" To "miCustomF9"
        If gSettings.CustomTag(Tool.ID - 12100, 1) <> "" Then
            InsertTag gSettings.CustomTag(Tool.ID - 12100, 2), gSettings.CustomTag(Tool.ID - 12100, 3)
        Else
            frmCustom.ShowForm (Tool.ID - 12100)
        End If
    Case "miSettings"
        frmSettings.Show vbModal, Me
        If gSettings.AppLCID <> gLCID Then
            MsgBox GetMsg(msgRestartLanguage), vbInformation
        End If
        If Not gIsXP Then acfPost.SetTrayIcon imgIcon16.Picture
    Case "miAccounts"
        frmLogin.Show vbModal, Me
        If tabPost.CurrTab = TAB_EDITOR Then
            If txtPostTit.Visible Then
                txtPostTit.SetFocus
            ElseIf txtPost.Visible Then
                txtPost.SetFocus
            End If
        ElseIf tabPost.CurrTab = TAB_MORE Then
            txtMore.SetFocus
        End If
        Exit Sub
    Case "miConnection", "lblAccount"
        frmAccount.Show vbModal, Me
        Me.Account = gAccount.Alias
    Case "miAddAccount"
        frmAccountWiz.Show vbModal, Me
        Exit Sub
    Case "miBlogProp"
        frmBlog.Show vbModal, Me
        If tabPost.CurrTab = TAB_PREVIEW Then PreviewPost
    Case "miPosts"
        If Not CanContinue() Then Exit Sub
        Select Case gSettings.PostMenu
        Case "0":  GetPost
        Case "1":  frmRecent.ShowForm 1
        Case "5":  frmRecent.ShowForm 5
        Case "10": frmRecent.ShowForm 10
        Case "15": frmRecent.ShowForm 15
        Case "20": frmRecent.ShowForm 20
        Case "n"
            On Error Resume Next
            intPosts = Val(InputBox(GetMsg(msgHowManyPosts)))
            If intPosts > 0 Then frmRecent.ShowForm intPosts
        End Select
    Case "miLast1"
        If Not CanContinue() Then Exit Sub
        frmRecent.ShowForm 1
        gSettings.PostMenu = "1"
    Case "miLast5"
        If Not CanContinue() Then Exit Sub
        frmRecent.ShowForm 5
        gSettings.PostMenu = "5"
    Case "miLast10"
        If Not CanContinue() Then Exit Sub
        frmRecent.ShowForm 10
        gSettings.PostMenu = "10"
    Case "miLast15"
        If Not CanContinue() Then Exit Sub
        frmRecent.ShowForm 15
        gSettings.PostMenu = "15"
    Case "miLast20"
        If Not CanContinue() Then Exit Sub
        frmRecent.ShowForm 20
        gSettings.PostMenu = "20"
    Case "miLastn"
        On Error Resume Next
        If Not CanContinue() Then Exit Sub
        intPosts = Val(InputBox(GetMsg(msgHowManyPosts)))
        If intPosts > 0 Then frmRecent.ShowForm intPosts
        gSettings.PostMenu = "n"
    Case "miPostID"
        If Not CanContinue() Then Exit Sub
        GetPost
        gSettings.PostMenu = "0"
    Case "miTemplate", "miMainTemplate"
        If Not CanContinue() Then Exit Sub
        GetTemplate "main"
        strCurrentFile = ""
        If UndoStack.Count >= 1 Then UndoStack(UndoStack.Count).Text = txtPost.TextRTF
    Case "miArchiveTemplate"
        If Not CanContinue() Then Exit Sub
        GetTemplate "archiveIndex"
        strCurrentFile = ""
        If UndoStack.Count >= 1 Then UndoStack(UndoStack.Count).Text = txtPost.TextRTF
    Case "miUndo"
        Undo
        Exit Sub
    Case "miRedo"
        Redo
        Exit Sub
    Case "miCopy"
        CopyText
        Exit Sub
    Case "miCut"
        CutText
        Exit Sub
    Case "miPaste"
        PasteText
        Exit Sub
    Case "miPasteText"
        PasteText True
        Exit Sub
    Case "miSelectAll"
        SelectAll
    Case "miFind"
        If tabPost.CurrTab = TAB_PREVIEW Then
            tabPost.CurrTab = TAB_EDITOR
        End If
        frmFind.Replace = False
        frmFind.Show vbModeless, Me
        Exit Sub
    Case "miFindNext"
        If tabPost.CurrTab = TAB_PREVIEW Then
            tabPost.CurrTab = TAB_EDITOR
        End If
        frmFind.Silent = True
        If frmFind.cboSearchStr.Text <> "" Then
            frmFind.DoFind
            Unload frmFind
        Else
            frmFind.Show vbModeless, Me
            Exit Sub
        End If
    Case "miReplace"
        If tabPost.CurrTab = TAB_PREVIEW Then
            tabPost.CurrTab = TAB_EDITOR
        End If
        frmFind.Replace = True
        frmFind.Show vbModeless, Me
        Exit Sub
    Case "miBold"
        If gSettings.XHTML Then
            InsertTag "<strong>", "</strong>"
        Else
            InsertTag "<b>", "</b>"
        End If
    Case "miItalic"
        If gSettings.XHTML Then
            InsertTag "<em>", "</em>"
        Else
            InsertTag "<i>", "</i>"
        End If
    Case "miUnderline"
        InsertTag "<u>", "</u>"
    Case "miStrike"
        InsertTag "<strike>", "</strike>"
    Case "miFont"
        frmFont.Show vbModal, Me
        If frmFont.HtmlTag <> "" Then InsertTag frmFont.HtmlTag, "</font>"
        Unload frmFont
    Case "miColor"
        strColor = GetColor()
        If strColor <> "" Then InsertTag strColor, ""
    Case "miColor01" To "miColor16"
        InsertTag "<font color=""" & Rgb2Html(QBColor(Tool.TagVariant)) & """>", "</font>"
    Case "miMoreColors"
        strColor = GetColor()
        If strColor <> "" Then InsertTag "<font color=""" & strColor & """>", "</font>"
    Case "miLeft"
        InsertTag "<div align=""left"">", "</div>"
    Case "miCenter"
        InsertTag "<div align=""center"">", "</div>"
    Case "miRight"
        InsertTag "<div align=""right"">", "</div>"
    Case "miJustify"
        InsertTag "<div align=""justify"">", "</div>"
    Case "miLink"
        frmLink.Show vbModal, Me
        If frmLink.HtmlTag <> "" Then InsertTag frmLink.HtmlTag, "</a>"
        Unload frmLink
        Set frmLink = Nothing
    Case "miImage"
        frmImage.Show vbModal, Me
        If frmImage.HtmlTag <> "" Then InsertTag frmImage.HtmlTag, ""
        Unload frmImage
        Set frmImage = Nothing
    Case "miLineBreak"
        If gSettings.XHTML Then
            InsertTag "<br />", ""
        Else
            InsertTag "<br>", ""
        End If
    Case "miParagraph" '4.00
        InsertTag "<p>", "</p>"
    Case "miNbsp"
        InsertTag "&nbsp;", ""
    Case "miHorRule"
        If gSettings.XHTML Then
            InsertTag "<hr />", ""
        Else
            InsertTag "<hr>", ""
        End If
    Case "miComment"
        InsertTag "<!-- ", " -->"
    Case "miBlockquote"
        InsertTag "<blockquote>", "</blockquote>"
    Case "miTable"
        frmTable.Show vbModal, Me
        If frmTable.HtmlTag <> "" Then InsertTag frmTable.HtmlTag, ""
        Unload frmTable
    Case "miHead1"
        InsertTag "<h1>", "</h1>"
    Case "miHead2"
        InsertTag "<h2>", "</h2>"
    Case "miHead3"
        InsertTag "<h3>", "</h3>"
    Case "miSubscript"
        InsertTag "<sub>", "</sub>"
    Case "miSuperscript"
        InsertTag "<sup>", "</sup>"
    Case "miJavascript"
        InsertTag "<script language=""JavaScript"">", "</script>"
    Case "miListOrder"
        Call InsertList("ol")
    Case "miListBullet"
        Call InsertList("ul")
    Case "miRepeatTag"
        If strTagBegin <> "" Then InsertTag strTagBegin, strTagEnd
    Case "miViewPage"
        If acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex >= 0 Then
            OpenBlogPage acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex
            Exit Sub
        End If
    Case "miCMS"
        If acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex >= 0 Then
            OpenCMSPage
            Exit Sub
        End If
    Case "miAdvanced"
        Call cmdMore_Click
    Case "miSaveTemplate", "miPublishTemplate"
        If acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex >= 0 Then
            If Trim(txtPost.Text) = "" Then
                MsgBox GetMsg(msgNothingToPost), vbExclamation
                Exit Sub
            End If
            DoEvents
            Me.Message = GetMsg(msgPosting)
            'Send Template
            If SaveTemplate(txtPost.Text, _
                            gBlogs(acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex).BlogID, _
                             (Tool.Name = "miPublishTemplate"), gSettings.Silent) Then
                If gSettings.ClearPost Then NewPost
                'SaveDraftPost
            End If
            Me.Message = ""
        End If
    Case "miPost", "miPublish"
        If acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex >= 0 Then
            If Trim(txtPost.Text) = "" And Trim(txtExcerpt.Text) = "" And Trim(txtMore.Text) = "" Then
                MsgBox GetMsg(msgNothingToPost), vbExclamation
                Exit Sub
            End If
            DoEvents
            Me.Message = GetMsg(msgPosting)
            'Create Categories Array
            aCateg = CreateCategArray()
            'Verify Media Info
            strPost = txtPost.Text
            If gPostID = "" And strMusic <> "" Then
                If gBlog.MediaInsert = 1 Then 'Top
                    strPost = gBlog.MediaString & vbCrLf & strPost
                ElseIf gBlog.MediaInsert = 2 Then 'Bottom
                    strPost = txtPost.Text & vbCrLf & gBlog.MediaString
                End If
            End If
            strPost = ReplaceMediaInfo(strPost)
            strMore = ReplaceMediaInfo(txtMore.Text)
            strExcerpt = ReplaceMediaInfo(txtExcerpt.Text)
            strKeywords = txtKeywords.Text
            'Send Post
            If Post(txtPostTit.Text, strPost, strMore, strExcerpt, strKeywords, aCateg, _
                    gBlogs(acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex).BlogID, _
                    gPostID, (Tool.Name = "miPublish"), gSettings.Silent, _
                    PostData.AllowComments, PostData.AllowPings, PostData.DateTime, _
                    PostData.TrackBack, PostData.TextFilter) Then
                If gSettings.ClearPost Then
                    NewPost
                Else
                    If gAccount.PostMethod = API_MT Then
                        Call GetPost(gPostID, True)
                    Else
                        imgStatus.Picture = acbMain.Tools("miPosts").GetPicture(0)
                        If strCurrentFile <> "" Then
                            lblStatus.Caption = GetLbl(lblPost) & ": " & gPostID & " - " & objMRU.CompressFileName(strCurrentFile)
                        Else
                            lblStatus.Caption = GetLbl(lblPost) & ": " & gPostID
                        End If
                        txtPost.Text = strPost
                        txtMore.Text = strMore
                        If gSettings.ColorizeCode Then
                            txtPost.Colorize
                            txtMore.Colorize
                            txtExcerpt.Colorize
                        End If
                    End If
                End If
                'SaveDraftPost
            End If
            Me.Message = ""
        End If
    Case "miMultipost"
        If acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex >= 0 Then
            If Trim(txtPost.Text) = "" Then
                MsgBox GetMsg(msgNothingToPost), vbExclamation
                Exit Sub
            ElseIf gPostID <> "" Then
                MsgBox GetMsg(msgOnlyNewPost), vbExclamation
                Exit Sub
            End If
            frmMultipost.Show vbModal, Me
        End If
    Case "miDelete", "miDelPublish"
        If acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex >= 0 Then
            DoEvents
            If MsgBox(GetMsg(msgWantDelete), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                DeletePost gPostID, (Tool.Name = "miDelPublish"), gSettings.Silent
            End If
        End If
    Case "miHelp"
        ' Call ShellExecute(Me.hwnd, "open", App.Path & "\wbloggar.chm", vbNullString, CurDir$, SW_SHOW)
        OpenWebPage "http://web.archive.org/web/20090330164700if_/http://wbloggar.com/faq.php"
        Exit Sub
    Case "miWebPage"
        OpenWebPage "https://github.com/lvcabral/w.bloggar"
        Exit Sub
    Case "miDonate"
        OpenWebPage "https://www.paypal.com/xclick/business=paypal%40wbloggar.com&item_name=w.bloggar&item_number=1"
        Exit Sub
    Case "miLicense"
        OpenDocument App.Path & "\License.txt"
        Exit Sub
    Case "miWhatsNew"
        OpenDocument App.Path & "\WhatsNew.txt"
        Exit Sub
    Case "miAbout"
        Unload frmAbout
        If Me.WindowState <> vbMinimized Then
            frmAbout.Show vbModal, Me
        Else
            frmAbout.Show , Me
        End If
    Case "miRestore"
        Unload frmAbout
        If Me.WindowState = vbMinimized Then
            acfPost.Restore
            SetForegroundWindow Me.hwnd
        Else
            Me.WindowState = vbNormal
        End If
    Case "miMaximize"
        Unload frmAbout
        If Me.WindowState = vbMinimized Then
            acfPost.Restore vbMaximized
        Else
            Me.WindowState = vbMaximized
        End If
    Case "miExit"
        Unload Me
        Exit Sub
    Case "lblMediaPlayer"
        If Trim(strMusic) <> "" And gBlog.MediaString <> "" Then
            InsertTag ReplaceMediaInfo(gBlog.MediaString), ""
        Else
            frmBlog.tabIndex.CurrTab = enuTabMedia
            frmBlog.Show vbModal, Me
        End If
    Case "lblStatus", "lblCapsLock", "lblNumLock", "lblInsert", "lblDate"
        Exit Sub
    Case Else
        MsgBox GetMsg(msgNotYetImpl), vbInformation
    End Select
    If tabPost.CurrTab = TAB_EDITOR And tabPost.Visible And Me.ActiveControl.Name <> "txtPostTit" Then
        txtPost.SetFocus
    End If
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".ToolClick"
End Sub

Private Sub acfPost_TrayLeftClick()
    Unload frmAbout
    acfPost.Restore
End Sub

Private Sub acfPost_TrayRightClick()
    acbMain.Bands("bndPopTray").PopupMenu
End Sub

Private Sub cboPostCat_Click()
    'Flag the change
    bolChanged = True
    'Check to reload
    If cboPostCat.ListIndex = cboPostCat.ListCount - 1 Then
        LoadCategories False
    ElseIf cboPostCat.ListIndex = 0 Then
        PostData.Categories = ""
        cmdCategories.Enabled = False
    Else
        cmdCategories.Enabled = cboPostCat.ListCount > 3
    End If
End Sub

Private Sub cmdCategories_Click()
    frmCategories.Show vbModal
End Sub

Private Sub cmdMore_Click()
    frmAdvanced.Show vbModal, Me
End Sub

Private Sub Form_Initialize()
    bolInsBody = True
    bolInsMore = True
    bolInsExpt = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Toggle tabs on [CTRL+TAB]
    '3.03 - If (KeyCode = vbKeyF12 And Shift = 0) Or
    If (Shift = vbCtrlMask And KeyCode = vbKeyTab) Then
        If tabPost.TabVisible(TAB_MORE) Then
            tabPost.CurrTab = IIf(tabPost.CurrTab = 0, 1, IIf(tabPost.CurrTab = 1, 2, IIf(tabPost.CurrTab = 2, 0, 2)))
        Else
            tabPost.CurrTab = IIf(tabPost.CurrTab = TAB_EDITOR, TAB_PREVIEW, TAB_EDITOR)
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim objTool As Tool
    Me.Caption = Me.Caption & "v" & App.Major & "." & Format$(App.Minor, "00")
    'Load Translation File
    LoadLang
    'Load User Toolbar Customization
    If FileExists(gAppDataPath & "\ignore.chg") Then
        Kill gAppDataPath & "\*.chg"
    ElseIf FileExists(gAppDataPath & "\wbloggar.chg") And IsCompiled() Then
        acbMain.LoadLayoutChanges gAppDataPath & "\wbloggar.chg", ddSOFile
        GetLbl lblNone
        If Err Then LoadLang
    End If
    'Configure ActiveBar
    acbMain.AlignToForm = True
    acbMain.AutoSizeChildren = ddASClientArea
    acbMain.ClientAreaControl = picPost
    'Translate Form
    LocalizeForm
    'Add Form Skin
    LoadSkin
    'Load Most Recent Updated Files
    Set objMRU = New clsMRUFileList
    objMRU.Load
    objMRU.MaxFileCount = 4
    DisplayMRU
    'Initialize Music String
    strMusic = "*"
    'Set Form Size and Position
    LoadFormSettings Me, (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)
    acbMain.Bands("bndStatus").Tools("lblInsert").Enabled = True
    'Adjust to Minimum size if lower than this values
    If Me.Width < acfPost.MinWidth Then Me.Width = acfPost.MinWidth
    If Me.Height < acfPost.MinHeight Then Me.Height = acfPost.MinHeight
    txtPost.Font.Name = gSettings.FontFace
    txtPost.Font.Size = gSettings.FontSize
    txtMore.Font.Name = gSettings.FontFace
    txtMore.Font.Size = gSettings.FontSize
    txtExcerpt.Font.Name = gSettings.FontFace
    txtExcerpt.Font.Size = gSettings.FontSize
    Me.Account = gAccount.Alias
    acfPost.MinimizeToTray = gSettings.Tray
    If gAccount.User <> "" Then
        LoadBlogs FileExists(gAppDataPath & "\blogs" & Format(gAccount.Current, "00") & ".xml")
    End If
    If gSettings.StartMinimized Then
        DoEvents
        WindowState = vbMinimized
    End If
    'Verify SpellChecking
    acbMain.Tools("miTSpelling").Enabled = FileExists(App.Path & "\Spell\" & gSettings.SpellLCID & ".dic")
    'Set the Blog List Size
    acbMain.Bands("bndTools").Tools("miBlogs").Width = gSettings.BlogListSize * Screen.TwipsPerPixelX
    'Star Undo Control
    ClearUndo
    trapUndo = True
    trapMove = True
    txtPost.AutoColorize = gSettings.ColorizeCode
    ClearPreview
    LoadColors
    LoadCustomTags
    'Start ToolBar Timer
    tmrToolBar.Enabled = True
    'Load tray icon
    If Not gIsXP Then acfPost.SetTrayIcon imgIcon16.Picture
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If WindowState = vbNormal Then
        acbMain.Tools("miRestore").Enabled = False
    Else
        acbMain.Tools("miRestore").Enabled = True
    End If
    tmrToolBar.Enabled = (frmPost.WindowState <> vbMinimized)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bolChanged Or Not Visible Then
        Cancel = False
    ElseIf UnloadMode <= vbFormCode Or UnloadMode = vbFormMDIForm Then
        Cancel = Not CanContinue()
    End If
    If Not Cancel Then
        If gSettings.OpenLastFile Then
            gSettings.PostFile = strCurrentFile
        Else
            gSettings.PostFile = ""
        End If
        Set objXMLReg = New XMLRegistry
        objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
        Call objXMLReg.SaveSetting(App.Title, "Settings", "PostFile", gSettings.PostFile)
        Call objXMLReg.SaveSetting(App.Title, "Settings", "PostMenu", gSettings.PostMenu)
        Set objXMLReg = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set objColor = Nothing
    If Visible Or WindowState = vbMinimized Then
        objMRU.Save
        SaveFormSettings Me
        If acbMain.Tools("miSaveTemplate").Visible Then TemplateMode False
        If gLCID = gSettings.AppLCID Then
            acbMain.SaveLayoutChanges gAppDataPath & "\wbloggar.chg", ddSOFile
        Else
            Kill gAppDataPath & "\wbloggar.chg"
        End If
    End If
End Sub

Private Sub ClearPreview()
On Error Resume Next
    webPreview.Navigate2 "about:blank"
End Sub

Private Sub PreviewPost()
On Error Resume Next
Dim strPost As String, strMore As String
Dim strPreview As String, strTemp As String, strCSS As String
    strPost = txtPost.Text
    strTemp = GetTempFolder()
    If Right$(strTemp, 1) = "\" Then
        strTemp = Left(strTemp, Len(strTemp) - 1)
    End If
    'Verify if it's template or post preview
    If InStr(1, txtPost.Text, "<html>", vbTextCompare) > 0 Or _
        acbMain.Tools("miSaveTemplate").Visible Then
        'Transform relative paths into full url's
        strPost = Path2URL(strPost, " background=")
        strPost = Path2URL(strPost, " src=")
        strPost = Path2URL(strPost, " href=")
        'Save the Preview page
        strPreview = strPost
            
        SaveBinaryFile strTemp & "\preview.htm", strPreview
    Else
        'Separate the Main Text and the More Text
        strPost = Replace(strPost, "_blank", "", , , vbTextCompare)
        strMore = GetMore(strPost) & Replace(txtMore.Text, "_blank", "", , , vbTextCompare)
        strPost = GetBody(strPost)
        'Verify Media Info
        If gPostID = "" And strMusic <> "" Then
            If gBlog.MediaInsert = 1 Then
                strPost = gBlog.MediaString & vbCrLf & strPost
            ElseIf gBlog.MediaInsert = 2 Then
                strPost = strPost & vbCrLf & gBlog.MediaString
            End If
        End If
        'Add Title if Exists
        If txtPostTit.Text <> "" And txtPostTit.Visible Then
            strPost = "<div " & gBlog.PreviewTitle & ">" & txtPostTit.Text & "</div>" & vbCrLf & strPost
        End If
        'Replace LineBreak, TabKey and Media PlaceHolders
        If gBlog.PreviewAutoBR Then
            strPost = Replace(strPost, vbLf, "<br>")
            strMore = Replace(strMore, vbLf, "<br>")
        End If
        'Replace Fix Space, Media Information and
        'Transform relative paths into full url's
        strPost = Replace(strPost, "?", "&nbsp;")
        strPost = ReplaceMediaInfo(strPost)
        strPost = Path2URL(strPost, " background=")
        strPost = Path2URL(strPost, " src=")
        strPost = Path2URL(strPost, " href=")
        If Trim(strMore) <> "" Then
            strMore = Replace(strMore, "?", "&nbsp;")
            strMore = ReplaceMediaInfo(strMore)
            strMore = Path2URL(strMore, " background=")
            strMore = Path2URL(strMore, " src=")
            strMore = Path2URL(strMore, " href=")
        End If
        'Do not add the default CSS tag
        If gBlog.PreviewCSS <> CSSTAG Then
            strCSS = gBlog.PreviewCSS
        End If
        'On post preview add head and body tags
        strPreview = "<html><head><title>Bloggar Preview</title></head>" & vbCrLf & _
                     strCSS & vbCrLf & gBlog.PreviewBody & vbCrLf & _
                     "<table " & gBlog.PreviewWidth & "><tr><td><div " & _
                     gBlog.PreviewAlign & " " & gBlog.PreviewStyle & ">" & vbCrLf & _
                     "%WBTEXT%</div></td></tr></table></body></html>"
        If Trim(strMore) <> "" Then
            SaveBinaryFile strTemp & "\preview.htm", Replace(strPreview, "%WBTEXT%", strPost & "<br>" & _
                                                              "<a href=moretext.htm>" & GetLbl(lblMoreText) & "</a>")
            SaveBinaryFile strTemp & "\moretext.htm", Replace(strPreview, "%WBTEXT%", strPost & strMore & "<br>" & _
                                                              "<a href=preview.htm>" & GetLbl(lblBackToMain) & "</a>")
        Else
            SaveBinaryFile strTemp & "\preview.htm", Replace(strPreview, "%WBTEXT%", strPost & "<br>")
        End If
    End If
    'Call thePreview
    webPreview.Navigate2 "file://" & strTemp & "\preview.htm"
End Sub

Private Sub DeletePreview()
    On Error Resume Next
    Kill GetTempFolder() & "preview.htm"
    Kill GetTempFolder() & "moretext.htm"
End Sub

Private Function Path2URL(ByVal strText As String, _
                          ByVal strFind As String) As String
On Error Resume Next
Dim strURL As String, strRoot As String
Dim lngStart As Long, lngPos As Long
    lngStart = 1
    'Find the Current Folder from the Blog URL
    strURL = gBlogs(frmPost.CurrentBlog).URL
    strURL = Left(strURL, InStrRev(strURL, "/"))
    'Find the Server Root from the Blog URL
    strRoot = gBlogs(frmPost.CurrentBlog).URL
    strRoot = Left(strRoot, InStr(InStr(strRoot, "://") + 3, strRoot, "/") - 1)
    'Transform relative paths into full url's
    lngPos = InStr(lngStart, strText, strFind, vbTextCompare)
    Do Until lngPos = 0
        If LCase(Mid(strText, lngPos, Len(strFind) + 8)) <> strFind & """http://" And _
           LCase(Mid(strText, lngPos, Len(strFind) + 7)) <> strFind & "http://" Then
            'Convert the relative Path to a complete URL
            If Mid(strText, lngPos + Len(strFind), 1) = """" Then
                If LCase(Mid(strText, lngPos, Len(strFind) + 2)) <> strFind & """/" Then
                    strText = Left(strText, lngPos - 1) & strFind & """" & strURL & Mid(strText, lngPos + Len(strFind) + 1)
                Else
                    strText = Left(strText, lngPos - 1) & strFind & """" & strRoot & Mid(strText, lngPos + Len(strFind) + 1)
                End If
            Else
                If LCase(Mid(strText, lngPos, Len(strFind) + 1)) <> strFind & "/" Then
                    strText = Left(strText, lngPos - 1) & strFind & "" & strURL & Mid(strText, lngPos + Len(strFind))
                Else
                    strText = Left(strText, lngPos - 1) & strFind & "" & strRoot & Mid(strText, lngPos + Len(strFind))
                End If
            End If
        End If
        lngStart = lngPos + 1
        lngPos = 0
        lngPos = InStr(lngStart, LCase(strText), strFind, vbTextCompare)
    Loop
    Path2URL = strText
End Function

Private Sub picPost_Resize()
    On Error Resume Next
    tabPost.Move 60, 120 + lblBack.Height, picPost.ScaleWidth - 120, picPost.ScaleHeight - (lblBack.Height + 150)
    lblBack.Width = tabPost.Width
End Sub

Private Sub tabPost_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error Resume Next
    If NewTab = TAB_PREVIEW Then
        PreviewPost
        webPreview.SetFocus
    Else
        DoEvents
        ClearPreview
        frmPost.Message = ""
        If NewTab = TAB_EDITOR Then
            txtPost.SetFocus
        Else
            txtMore.SetFocus
        End If
    End If
    If OldTab = TAB_PREVIEW Then
        DeletePreview
    End If
End Sub

Private Sub txtCommand_Change()
    If WindowState = vbMinimized Then
        Unload frmAbout
        acfPost.Restore
    End If
    If FileExists(txtCommand.Text) Then
        If CanContinue() Then LoadPostFile txtCommand.Text
        txtCommand.Text = ""
    End If
End Sub

Private Sub txtExcerpt_Change()
    'Flag the change
    bolChanged = True
    'Colorize
    If gSettings.ColorizeCode And txtExcerpt.SelStart > 0 Then
        If Mid(txtExcerpt.Text, txtExcerpt.SelStart, 1) = ">" Then
            txtExcerpt.Colorize txtExcerpt.GetLineFirstCharIndex, txtExcerpt.SelStart + 1
        End If
    End If
    ' Undo
    UndoOnChangeEvent txtExcerpt, UndoExcerpt, RedoExcerpt
End Sub

Private Sub txtExcerpt_GotFocus()
    With acbMain.Bands("bndStatus").Tools("lblInsert")
       .Enabled = bolInsExpt
    End With
End Sub

Private Sub txtExcerpt_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyInsert Then
        bolInsExpt = Not bolInsExpt
        With acbMain.Bands("bndStatus").Tools("lblInsert")
           .Enabled = bolInsExpt
        End With
    End If
    KeyDownEvent txtExcerpt, KeyCode, Shift
End Sub

Private Sub txtExcerpt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    acbMain.Bands("bndPopContext").PopupMenu
End Sub

Private Sub txtExcerpt_SelChange()
    SelChangeEvent txtExcerpt, UndoExcerpt, RedoExcerpt
End Sub

Private Sub txtKeywords_Change()
    'Flag the change
    bolChanged = True
End Sub

Private Sub txtMore_Change()
    'Flag the change
    bolChanged = True
    'Colorize
    If gSettings.ColorizeCode And txtMore.SelStart > 0 Then
        If Mid(txtMore.Text, txtMore.SelStart, 1) = ">" Then
            txtMore.Colorize txtMore.GetLineFirstCharIndex, txtMore.SelStart + 1
        End If
    End If
    ' Undo
    UndoOnChangeEvent txtMore, UndoMore, RedoMore
End Sub

Private Sub txtMore_GotFocus()
    With acbMain.Bands("bndStatus").Tools("lblInsert")
       .Enabled = bolInsMore
    End With
End Sub

Private Sub txtMore_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyInsert Then
        bolInsMore = Not bolInsMore
        With acbMain.Bands("bndStatus").Tools("lblInsert")
           .Enabled = bolInsMore
        End With
    End If
    KeyDownEvent txtMore, KeyCode, Shift
End Sub

Private Sub txtMore_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    acbMain.Bands("bndPopContext").PopupMenu
End Sub

Private Sub txtMore_SelChange()
    SelChangeEvent txtMore, UndoMore, RedoMore
End Sub

Private Sub txtPost_Change()
    'Flag the change
    bolChanged = True
    'Colorize
    If gSettings.ColorizeCode And txtPost.SelStart > 0 Then
        If Mid(txtPost.Text, txtPost.SelStart, 1) = ">" Then
            txtPost.Colorize txtPost.GetLineFirstCharIndex, txtPost.SelStart + 1
        End If
    End If
    ' Undo
    UndoOnChangeEvent txtPost, UndoStack, RedoStack
End Sub

Private Sub txtPost_GotFocus()
    With acbMain.Bands("bndStatus").Tools("lblInsert")
       .Enabled = bolInsMore
    End With
End Sub

Private Sub txtPost_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If Shift = 0 And KeyCode = vbKeyInsert Then
        bolInsBody = Not bolInsBody
        With acbMain.Bands("bndStatus").Tools("lblInsert")
           .Enabled = bolInsBody
        End With
        Exit Sub
    End If
    KeyDownEvent txtPost, KeyCode, Shift
End Sub

Private Sub txtPost_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 1, 2, 7, 9, 12, 14, 16, 19, 20, 21
        KeyAscii = 0
    'Case Else
        'Debug.Print KeyAscii
    End Select
End Sub

Private Sub txtPost_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    acbMain.Bands("bndPopContext").PopupMenu
End Sub

Private Sub txtPost_SelChange()
    SelChangeEvent txtPost, UndoStack, RedoStack
End Sub

Private Sub tmrToolBar_Timer()
On Error Resume Next
Dim oControl As Control
Dim strPlay As String
    ' Enable and disable COPY/CUT/PASTE buttons
    ' according to the control with focus
    ' and Clipboard contents
    If Not Me.ActiveControl Is Nothing And WindowState <> vbMinimized Then
        acbMain.Tools("miSave").Enabled = bolChanged
        acbMain.Tools("miSaveAs").Enabled = Len(Trim(txtPostTit.Text + txtPost.Text + txtMore.Text + txtExcerpt.Text)) > 0
        With acbMain.Bands("bndStatus").Tools("lblMediaPlayer")
            strPlay = MediaPlayerInfo("Title")
            If strPlay <> strMusic Then
                strMusic = strPlay
                If Len(strMusic) Then
                    acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITNormal, imgMediaPlayer(1).Picture
                    acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITHover, imgMediaPlayer(1).Picture
                Else
                    acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITNormal, imgMediaPlayer(0).Picture
                    acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITHover, imgMediaPlayer(0).Picture
                End If
                acbMain.RecalcLayout
                If Len(strMusic) Then
                    .ToolTipText = strMusic & " - " & MediaPlayerInfo("Author")
                Else
                    .ToolTipText = GetLbl(lblNoMedia)
                End If
            End If
        End With
        Set oControl = Me.ActiveControl
        'se for algum controle que aceita Copy/Cut/Paste
        If InStr("TextBox@ActiveText@ActiveDate@HtmlEdit ", TypeName(oControl)) Then
            acbMain.Tools("miCopy").Enabled = (oControl.SelText <> "")
            acbMain.Tools("miCut").Enabled = (oControl.SelText <> "")
            acbMain.Tools("miPaste").Enabled = (Clipboard.GetText <> "")
            acbMain.Tools("miPasteText").Enabled = (Clipboard.GetText <> "")
            acbMain.Tools("miSelectAll").Enabled = (oControl.Text <> "")
        ElseIf TypeName(oControl) = "ComboBox" Then
            'O Combo s? aceita Paste no Style = 0
            acbMain.Tools("miUndo").Enabled = SendMessage(oControl.hwnd, EM_CANUNDO, 0, ByVal 0&)
            acbMain.Tools("miCopy").Enabled = (oControl.Text <> "")
            If (oControl.Style = 0) Then
                acbMain.Tools("miCut").Enabled = (oControl.Text <> "")
                acbMain.Tools("miPaste").Enabled = (Clipboard.GetText <> "")
                acbMain.Tools("miPasteText").Enabled = (Clipboard.GetText <> "")
            End If
            acbMain.Tools("miSelectAll").Enabled = (oControl.Text <> "")
        Else
            ' o controle n?o aceita Cut/Copy/Paste
            acbMain.Tools("miUndo").Enabled = False
            acbMain.Tools("miCopy").Enabled = False
            acbMain.Tools("miCut").Enabled = False
            acbMain.Tools("miPaste").Enabled = False
            acbMain.Tools("miPasteText").Enabled = False
            acbMain.Tools("miSelectAll").Enabled = False
        End If
    End If
End Sub

Private Sub txtPostTit_Change()
    'Flag the change
    bolChanged = True
End Sub

Private Sub webPreview_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If (Left(URL, 4) = "http" Or Left(URL, 4) = "ftp:") And URL <> "http:///" Then
        OpenWebPage URL
        Cancel = True
    End If
End Sub

Private Sub webPreview_StatusTextChange(ByVal Text As String)
    If tabPost.CurrTab = TAB_PREVIEW Then
        frmPost.Message = Text
    End If
End Sub

Public Property Get Changed() As Boolean
    Changed = bolChanged
End Property

Public Property Let Changed(ByVal vNewValue As Boolean)
    bolChanged = vNewValue
End Property

Public Property Let CurrentBlog(ByVal Index As Integer)
    acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex = Index
End Property

Public Property Get CurrentBlog() As Integer
    CurrentBlog = acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex
End Property

Public Property Let CurrentTab(ByVal Index As Integer)
    tabPost.CurrTab = Index
End Property

Public Property Get CurrentTab() As Integer
    CurrentTab = tabPost.CurrTab
End Property

Public Property Let Message(ByVal strText As String)
    acbMain.Bands("bndStatus").Tools("lblStatus").Caption = Replace(strText, "&", "&&")
End Property

Public Property Let Account(ByVal strText As String)
Dim strIcon As String
    acbMain.Bands("bndPopTools").Tools("miCMS").Caption = Replace(GetLbl(lblOpenSite), "%1", ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Name", App.Path & "\CMS\CMS.ini"))
    strIcon = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Icon", App.Path & "\CMS\CMS.ini")
    If FileExists(App.Path & "\CMS\" & strIcon) Then
        acbMain.Bands("bndStatus").Tools("lblAccount").Style = ddSIconText
        acbMain.Bands("bndStatus").Tools("lblAccount").SetPicture ddITNormal, LoadPicture(App.Path & "\CMS\" & strIcon)
        acbMain.Bands("bndStatus").Tools("lblAccount").SetPicture ddITHover, LoadPicture(App.Path & "\CMS\" & strIcon)
        acbMain.Bands("bndPopTools").Tools("miCMS").SetPicture ddITNormal, LoadPicture(App.Path & "\CMS\" & strIcon)
        acbMain.Bands("bndPopTools").Tools("miCMS").SetPicture ddITHover, LoadPicture(App.Path & "\CMS\" & strIcon)
    Else
        acbMain.Bands("bndStatus").Tools("lblAccount").Style = ddSStandard
        acbMain.Bands("bndStatus").Tools("lblAccount").SetPicture ddITNormal, LoadPicture()
        acbMain.Bands("bndPopTools").Tools("miCMS").SetPicture ddITNormal, LoadPicture()
        acbMain.Bands("bndPopTools").Tools("miCMS").SetPicture ddITHover, LoadPicture()
    End If
    acbMain.Bands("bndStatus").Tools("lblAccount").Caption = strText
    acbMain.RecalcLayout
End Property

Public Sub LoadPostFile(ByVal strFile As String, Optional ByVal bolTemp As Boolean)
On Error Resume Next
Dim strPost As String, strTitle As String, strFName As String
    strFile = Replace(strFile, """", "")
    If FileExists(strFile) Then
        If Left(GetNamePart(strFile), 1) = "~" Or LCase(Right(strFile, 4)) = ".tmp" Then
            bolTemp = True
        End If
        strPost = GetBinaryFile(strFile)
        If Not bolTemp Then
            strFName = " - " & objMRU.CompressFileName(strFile)
        End If
        'Just to be compatible to the old .post format
        If Left(strPost, 5) <> "<?xml" Then
            If SupportsTitle() Then
                If SupportsCategory() Then
                    If cboPostCat.ListCount > 0 Then cboPostCat.ListIndex = 0
                    Select Case gAccount.GetPostsMethod
                    Case API_B2
                        If cboPostCat.ListCount > 1 Then
                            SearchItemData cboPostCat, Val(GetCateg(strPost, False))
                        End If
                    Case API_MT
                        PostData.Categories = GetCateg(strPost, False)
                        If Len(PostData.Categories) > 0 And cboPostCat.ListCount > 1 Then
                            SearchItemData cboPostCat, Val(Split(PostData.Categories, vbTab)(0))
                        End If
                    Case API_METAWEBLOG '3.03
                        PostData.Categories = GetCateg(strPost, False)
                        If Len(PostData.Categories) > 0 And cboPostCat.ListCount > 1 Then
                            SearchComboBox cboPostCat, (Split(PostData.Categories, vbTab)(0))
                        End If
                    Case Else
                        If cboPostCat.ListCount > 1 Then
                            If cboPostCat.ItemData(0) > 0 Then
                                SearchItemData cboPostCat, Val(GetCateg(strPost, False))
                            Else
                                SearchComboBox cboPostCat, GetCateg(strPost, False)
                            End If
                        End If
                    End Select
                End If
                txtPostTit.Text = GetTitle(strPost, False)
                txtPost.Text = GetBody(strPost, False)
            Else
                strTitle = GetTitle(strPost, False)
                If Trim(strTitle) <> "" Then
                    txtPost.Text = GetTitle(strPost, False) & vbCrLf & GetBody(strPost, False)
                Else
                    txtPost.Text = strPost
                End If
            End If
        Else 'New xml .post format
            Call PostData.LoadData(strFile)
            If PostData.BlogID <> "" And PostData.BlogID <> gBlogs(acbMain.Bands("bndTools").Tools("miBlogs").CBListIndex).BlogID Then
                If MsgBox(GetMsg(msgPostFileWithID) & vbCrLf & GetMsg(msgLoadAsDraft), vbExclamation + vbYesNo) = vbYes Then
                    PostData.AccountID = -1
                    PostData.BlogID = ""
                    strFName = ""
                    bolTemp = True
                Else
                    Exit Sub
                End If
            Else
                gPostID = PostData.PostID
            End If
            If SupportsTitle() Then
                If SupportsCategory() Then
                    If cboPostCat.ListCount > 0 Then cboPostCat.ListIndex = 0
                    Select Case gAccount.GetPostsMethod
                    Case API_B2
                        If cboPostCat.ListCount > 1 Then
                            SearchItemData cboPostCat, Val(PostData.Categories)
                        End If
                    Case API_MT
                        If Len(PostData.Categories) > 0 And cboPostCat.ListCount > 1 Then
                            SearchItemData cboPostCat, Val(Split(PostData.Categories, vbTab)(0))
                        End If
                    Case API_METAWEBLOG '3.03
                        If Len(PostData.Categories) > 0 And cboPostCat.ListCount > 1 Then
                            SearchComboBox cboPostCat, (Split(PostData.Categories, vbTab)(0))
                        End If
                    Case Else
                        If cboPostCat.ListCount > 1 Then
                            If cboPostCat.ItemData(0) > 0 Then
                                SearchItemData cboPostCat, Val(PostData.Categories)
                            Else
                                SearchComboBox cboPostCat, PostData.Categories
                            End If
                        End If
                    End Select
                End If
                txtPostTit.Text = PostData.Title
                txtPost.Text = PostData.Text
                If tabPost.TabVisible(TAB_MORE) Then
                    txtMore.Text = PostData.More
                    If gAccount.PostMethod = API_MT Then
                        txtExcerpt.Text = PostData.Excerpt
                        txtKeywords.Text = PostData.Keywords
                    End If
                End If
            Else
                If Trim(PostData.Title) <> "" Then
                    txtPost.Text = PostData.Title & vbCrLf & PostData.Text
                Else
                    txtPost.Text = PostData.Text
                End If
                If tabPost.TabVisible(TAB_MORE) Then
                    txtMore.Text = PostData.More
                End If
            End If
        End If
        If gPostID <> "" Then
            If gPostID = "main" Then
                TemplateMode True
                imgStatus.Picture = acbMain.Tools("miTemplate").GetPicture(0)
                lblStatus.Caption = GetMsg(msgMainTemplate) & strFName
            ElseIf gPostID = "archiveIndex" Then
                TemplateMode True
                imgStatus.Picture = acbMain.Tools("miTemplate").GetPicture(0)
                lblStatus.Caption = GetMsg(msgArchiveTemplate) & strFName
            Else
                imgStatus.Picture = acbMain.Tools("miPosts").GetPicture(0)
                lblStatus.Caption = GetLbl(lblPost) & ": " & gPostID & strFName
            End If
        Else
            imgStatus.Picture = acbMain.Tools("miNew").GetPicture(0)
            lblStatus.Caption = GetLbl(lblPost) & ": " & GetMsg(msgNewPost) & strFName
        End If
        If gSettings.ColorizeCode Then
            txtPost.Colorize
            txtMore.Colorize
            txtExcerpt.Colorize
        End If
        If UndoStack.Count >= 1 Then UndoStack(1).Text = txtPost.TextRTF
        If Not bolTemp Then
            objMRU.AddFile strFile
            strCurrentFile = strFile
        Else
            strCurrentFile = ""
        End If
        DisplayMRU
        bolChanged = bolTemp
    End If
    If txtPostTit.Visible Then
        txtPostTit.SetFocus
    ElseIf txtPost.Visible Then
        txtPost.SetFocus
    End If
End Sub

Private Sub DisplayMRU()
Dim iFile As Integer, iCount As Integer
    For iFile = 1 To objMRU.FileCount
        If (objMRU.FileExists(iFile)) Then
            With acbMain.Bands("bndPopFile").Tools("miFMRU" & Trim$(Str(iFile)))
                .Visible = True
                .Caption = objMRU.MenuCaption(iFile)
                .TagVariant = CStr(iFile)
                If iFile = 1 Then
                    .Checked = (objMRU.file(1) = strCurrentFile)
                End If
            End With
            iCount = iCount + 1
        End If
    Next iFile
    ' Debug.Print (objMRU.FileCount > 0)
    acbMain.Bands("bndPopFile").Tools("miFMRUSep").Visible = (iCount > 0)
End Sub

Private Sub CheckSpelling()
On Error Resume Next
    Screen.MousePointer = vbHourglass
    objSpellCheck.Dictionary = App.Path & "\Spell\" & gSettings.SpellLCID & ".dic"
    Select Case tabPost.CurrTab
    Case TAB_EDITOR
        Set objSpellCheck.TextBox = txtPost
    Case TAB_MORE
        If Me.ActiveControl.Name = "txtExcerpt" Then
            Set objSpellCheck.TextBox = txtExcerpt
        Else
            Set objSpellCheck.TextBox = txtMore
        End If
    Case Else
        tabPost.CurrTab = TAB_EDITOR
        DoEvents
        Set objSpellCheck.TextBox = txtPost
    End Select
    If Err Then
        Screen.MousePointer = vbDefault
        MsgBox GetMsg(msgWaitDicLoad), vbExclamation
    Else
        objSpellCheck.SpellCheck True
    End If
End Sub

Public Sub LoadSkin()
On Error Resume Next
Dim objTool As Tool
Dim strSkin As String, strFont As String, strExt As String
    Screen.MousePointer = vbHourglass
    strSkin = gSettings.SkinFolder & "\skin.ini"
    'Set Font
    strFont = ReadINI("Skin", "Font", strSkin, "Tahoma")
    acbMain.Font.Name = strFont
    acbMain.ControlFont.Name = strFont
    acbMain.ChildBandFont.Name = strFont
    lblStatus.Font.Name = strFont
    pnlEditor.Font.Name = strFont
    tabPost.Font.Name = strFont
    'Set Colors
    With acbMain
        .BackColor = Val(ReadINI("Skin", "BackColor", strSkin, "&H8000000F&"))
        .ForeColor = Val(ReadINI("Skin", "ForeColor", strSkin, "&H80000012&"))
        .HighLightColor = Val(ReadINI("Skin", "HighLightColor", strSkin, "&H80000014&"))
        .ShadowColor = Val(ReadINI("Skin", "ShadowColor", strSkin, "&H80000010&"))
        .ThreeDDarkShadow = Val(ReadINI("Skin", "3DShadowColor", strSkin, "&H80000015&"))
        .ThreeDLight = Val(ReadINI("Skin", "3DLightColor", strSkin, "&H80000016&"))
    End With
    lblBack.BackColor = Val(ReadINI("Skin", "LabelBackColor", strSkin, "&H8000000C&"))
    lblStatus.ForeColor = Val(ReadINI("Skin", "LabelForeColor", strSkin, "&H80000009&"))
    frmPost.BackColor = acbMain.BackColor
    picPost.BackColor = acbMain.BackColor
    With tabPost
        .BackColor = acbMain.BackColor
        .BackTabColor = acbMain.BackColor
        .FrontTabColor = acbMain.BackColor
        .ForeColor = acbMain.ForeColor
        .FrontTabForeColor = acbMain.ForeColor
        .TabOutlineColor = acbMain.ShadowColor
    End With
    pnlEditor.BackColor = acbMain.BackColor
    pnlEditor.ForeColor = acbMain.ForeColor
    'Set Grab Handle Style
    acbMain.Bands("bndMenu").GrabHandleStyle = Val(ReadINI("Skin", "GrabHandleStyle", strSkin, "1"))
    acbMain.Bands("bndTools").GrabHandleStyle = Val(ReadINI("Skin", "GrabHandleStyle", strSkin, "1"))
    acbMain.Bands("bndFormat").GrabHandleStyle = Val(ReadINI("Skin", "GrabHandleStyle", strSkin, "1"))
    acbMain.Bands("bndHTML").GrabHandleStyle = Val(ReadINI("Skin", "GrabHandleStyle", strSkin, "1"))
    'Set XPLook
    If ReadINI("Skin", "XPLook", strSkin, "1") = "1" Then
        With acbMain
            .XPLook = True
            For Each objTool In .Bands("bndStatus").Tools
                objTool.LabelBevel = ddLBFlat
            Next
        End With
    Else
        With acbMain
            .XPLook = False
            For Each objTool In .Bands("bndStatus").Tools
                objTool.LabelBevel = ddLBInset
            Next
        End With
    End If
    'Load WallPaper
    If FileExists(gSettings.SkinFolder & "\" & ReadINI("Skin", "Picture", strSkin)) Then
        acbMain.Picture = LoadPicture(gSettings.SkinFolder & "\" & ReadINI("Skin", "Picture", strSkin))
    Else
        Set acbMain.Picture = Nothing
    End If
    'Load Tab Icons
    strExt = ReadINI("Skin", "IconsExt", strSkin)
    If FileExists(gSettings.SkinFolder & "\editor." & strExt) Then
        tabPost.TabPicture(TAB_EDITOR) = LoadPicture(gSettings.SkinFolder & "\editor." & strExt, vbLPSmall)
        tabPost.TabHeight = 300
    Else
        tabPost.TabPicture(TAB_EDITOR) = LoadPicture()
        tabPost.TabHeight = 270
    End If
    If FileExists(gSettings.SkinFolder & "\MoreText." & strExt) Then
        tabPost.TabPicture(TAB_MORE) = LoadPicture(gSettings.SkinFolder & "\MoreText." & strExt, vbLPSmall)
        tabPost.TabHeight = 300
    Else
        tabPost.TabPicture(TAB_MORE) = LoadPicture()
        tabPost.TabHeight = 270
    End If
    If FileExists(gSettings.SkinFolder & "\preview." & strExt) Then
        tabPost.TabPicture(TAB_PREVIEW) = LoadPicture(gSettings.SkinFolder & "\preview." & strExt, vbLPSmall)
        tabPost.TabHeight = 300
    Else
        tabPost.TabPicture(TAB_PREVIEW) = LoadPicture()
        tabPost.TabHeight = 270
    End If
    'Load Music Icons
    If FileExists(gSettings.SkinFolder & "\MusicOff." & strExt) Then
        imgMediaPlayer(0) = LoadPicture(gSettings.SkinFolder & "\MusicOff." & strExt, vbLPSmall)
    End If
    If FileExists(gSettings.SkinFolder & "\MusicOn." & strExt) Then
        imgMediaPlayer(1) = LoadPicture(gSettings.SkinFolder & "\MusicOn." & strExt, vbLPSmall)
    End If
    'Refresh Music Icon
    If Len(strMusic) Then
        acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITNormal, imgMediaPlayer(1).Picture
        acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITHover, imgMediaPlayer(1).Picture
    Else
        acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITNormal, imgMediaPlayer(0).Picture
        acbMain.Bands("bndStatus").Tools("lblMediaPlayer").SetPicture ddITHover, imgMediaPlayer(0).Picture
    End If
    'Load Toolbar Icons
    If strExt <> "" Then
        For Each objTool In acbMain.Tools
            If FileExists(gSettings.SkinFolder & "\" & Mid(objTool.Name, 3) & "." & strExt) Then
                objTool.SetPicture ddITNormal, LoadPicture(gSettings.SkinFolder & "\" & Mid(objTool.Name, 3) & "." & strExt, vbLPSmall)
                acbMain.ApplyAll objTool
                If objTool.Name = "miSettings" Then
                    acbMain.Bands("bndPopFile").Tools("miSettings").Caption = _
                    acbMain.Bands("bndPopFile").Tools("miSettings").Caption & "..."
                End If
            End If
        Next
    End If
    acbMain.RecalcLayout
    'Refresh Post Icon
    If gPostID <> "" Then
        If gPostID = "main" Then
            imgStatus.Picture = acbMain.Tools("miTemplate").GetPicture(0)
        ElseIf gPostID = "archiveIndex" Then
            imgStatus.Picture = acbMain.Tools("miTemplate").GetPicture(0)
        Else
            imgStatus.Picture = acbMain.Tools("miPosts").GetPicture(0)
        End If
    Else
        imgStatus.Picture = acbMain.Tools("miNew").GetPicture(0)
    End If
    'Tab Settings
    tabPost.Appearance = Val(ReadINI("Skin", "TabAppearance", strSkin, "2"))
    tabPost.Position = Val(ReadINI("Skin", "TabPosition", strSkin, "0"))
    tabPost.Style = Val(ReadINI("Skin", "TabStyle", strSkin, "3"))
    If tabPost.Style = tsButtons Then tabPost.TabHeight = 330
    tabPost.Refresh
    Screen.MousePointer = vbDefault
End Sub

Sub LoadLang()
On Error Resume Next
    If FileExists(App.Path & "\Lang\" & gSettings.AppLCID & ".lng") Then
        acbMain.Load "", App.Path & "\Lang\" & gSettings.AppLCID & ".lng", ddSOFile
        If Right(acbMain.Bands("bndPopFile").Tools("miSettings").Caption, 3) <> "..." Then
            acbMain.Bands("bndPopFile").Tools("miSettings").Caption = _
            acbMain.Bands("bndPopFile").Tools("miSettings").Caption & "..."
        End If
        'Translate Bands Caption
        acbMain.Bands("bndTools").Caption = acbMain.Tools("miToolsBar").Caption
        acbMain.Bands("bndFormat").Caption = acbMain.Tools("miFormatBar").Caption
        acbMain.Bands("bndHTML").Caption = acbMain.Tools("miHTMLBar").Caption
        acbMain.Bands("bndStatus").Caption = acbMain.Tools("miStatusBar").Caption
        'Get the LocaleID
        gLCID = Val(acbMain.Tools("LocaleID").Caption)
    Else
        gLCID = gSettings.AppLCID
    End If
End Sub
Private Function CreateCategArray()
On Error GoTo ErrorHandler
Dim aCateg(), aTemp, c As Integer, objStruc As xmlStruct
    Select Case gAccount.GetCategMethod
    Case API_MT
        With cboPostCat
            If .ListIndex > 0 Then
                If Len(PostData.Categories) > 0 Then
                    PostData.Categories = Mid(PostData.Categories, Len(CATEG_ID_MASK) + 2)
                End If
                If InStr(PostData.Categories, Format(.ItemData(.ListIndex), CATEG_ID_MASK)) > 0 Then
                    PostData.Categories = Replace(PostData.Categories, Format(.ItemData(.ListIndex), CATEG_ID_MASK) & vbTab, "")
                End If
                PostData.Categories = Format(.ItemData(.ListIndex), CATEG_ID_MASK) & vbTab & PostData.Categories
            Else
                PostData.Categories = ""
            End If
        End With
        If PostData.Categories <> "" Then
            aTemp = Split(PostData.Categories, vbTab)
            For c = 0 To UBound(aTemp)
                If IsNumeric(aTemp(c)) Then
                    Set objStruc = New xmlStruct
                    objStruc.Add "categoryId", CStr(Val(aTemp(c)))
                    objStruc.Add "isPrimary", CBool(c = 0)
                    ReDim Preserve aCateg(c)
                    Set aCateg(c) = objStruc
                End If
            Next
        End If
    Case API_B2
        If PostData.Categories <> "" Then
            If gAccount.UTF8 Or gAccount.UTF8OnPost Then
                aTemp = Split(UTF8_Encode(PostData.Categories), vbTab)
            Else
                aTemp = Split(PostData.Categories, vbTab)
            End If
            For c = 0 To UBound(aTemp)
                ReDim Preserve aCateg(c)
                aCateg(c) = aTemp(c)
            Next
        ElseIf cboPostCat.ListIndex > 0 Then
            ReDim aCateg(0)
            aCateg(0) = cboPostCat.ItemData(cboPostCat.ListIndex)
        End If
    Case API_METAWEBLOG
        If InStr(PostData.Categories, cboPostCat.List(cboPostCat.ListIndex)) > 0 Then
            If gAccount.UTF8 Or gAccount.UTF8OnPost Then
                aTemp = Split(UTF8_Encode(PostData.Categories), vbTab)
            Else
                aTemp = Split(PostData.Categories, vbTab)
            End If
            For c = 0 To UBound(aTemp)
                ReDim Preserve aCateg(c)
                aCateg(c) = aTemp(c)
            Next
        ElseIf cboPostCat.ListIndex > 0 Then
            ReDim aCateg(0)
            If gAccount.UTF8 Or gAccount.UTF8OnPost Then
                aCateg(0) = UTF8_Encode(cboPostCat.List(cboPostCat.ListIndex))
            Else
                aCateg(0) = cboPostCat.List(cboPostCat.ListIndex)
            End If
        End If
    Case Else
        If cboPostCat.ListIndex > 0 Then
            ReDim aCateg(0)
            If cboPostCat.ItemData(cboPostCat.ListIndex) <> 0 Then
                aCateg(0) = cboPostCat.ItemData(cboPostCat.ListIndex)
            Else
                If gAccount.UTF8 Or gAccount.UTF8OnPost Then
                    aCateg(0) = UTF8_Encode(cboPostCat.List(cboPostCat.ListIndex))
                Else
                    aCateg(0) = cboPostCat.List(cboPostCat.ListIndex)
                End If
            End If
        End If
    End Select
    CreateCategArray = aCateg
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name & ".CreateCategArray"
End Function

Private Sub KeyDownEvent(ctlText As Control, ByRef KeyCode As Integer, ByRef Shift As Integer)
Dim lngPos As Long, strSel As String, strSpc As String
On Error Resume Next
    If Shift = 0 And KeyCode = vbKeyTab Then
        If ctlText.Name = "txtPost" And acbMain.Tools("miSaveTemplate").Visible Then
            strSpc = Space(gSettings.TabSpaces)
        Else
            strSpc = String(gSettings.TabSpaces, Chr(183))
        End If
        If ctlText.SelLength > 0 Then
            lngPos = ctlText.SelStart
            strSel = strSpc & Replace(txtPost.SelText, vbCrLf, vbCrLf & strSpc)
            ctlText.SelText = strSel
            ctlText.SelStart = lngPos
            ctlText.SelLength = Len(strSel)
            ctlText.Colorize lngPos, Len(strSel)
        Else
            ctlText.SelText = strSpc
        End If
        KeyCode = 0
        Shift = 0
        DoEvents
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyReturn Then
        acbMain_ToolClick acbMain.Tools("miPost")
        KeyCode = 0
        Shift = 0
        DoEvents
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyR And _
           RedoStack.Count = 0 Then
        KeyCode = 0
        Shift = 0
        DoEvents
    ElseIf (Shift = vbCtrlMask And KeyCode = vbKeyTab) Then
        KeyCode = 0
        Shift = 0
        DoEvents
    End If
End Sub

Private Sub UndoOnChangeEvent(txtTarget As Control, colUndo As Collection, colRedo As Collection)
Dim newElement As UndoRedo   'create new undo element
Dim c%, l&
    ' Undo
    If Not trapUndo Then Exit Sub 'because trapping is disabled

    Set newElement = New UndoRedo   'create new undo element
    'remove all redo items because of the change
    For c% = 1 To colRedo.Count
        colRedo.Remove 1
    Next c%

    'set the values of the new element
    newElement.SelStart = txtTarget.SelStart
    newElement.TextLen = Len(txtTarget.TextRTF)
    newElement.Text = txtTarget.TextRTF
    'add it to the undo stack
    colUndo.Add Item:=newElement
    If colUndo.Count > MAX_UNDO Then colUndo.Remove 1
    'call the Selection Change event
    SelChangeEvent txtTarget, colUndo, colRedo
End Sub

Private Sub SelChangeEvent(txtTarget As Control, colUndo As Collection, colRedo As Collection)
    If trapUndo = False Then Exit Sub
    If colUndo.Count > 0 And trapMove Then
        colUndo.Item(colUndo.Count).SelStart = txtTarget.SelStart
    End If
    acbMain.Tools("miUndo").Enabled = colUndo.Count > 1
    acbMain.Tools("miRedo").Enabled = colRedo.Count > 0
End Sub

Private Function CanContinue() As Boolean
    CanContinue = True
    If bolChanged Then
        Select Case MsgBox(GetMsg(msgQueryUnload), vbExclamation + vbYesNoCancel)
        Case vbYes
            If Not SavePostData(strCurrentFile) Then
                CanContinue = False
            End If
        Case vbCancel
           CanContinue = False
        End Select
    End If
End Function

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
Dim i As Integer
    tabPost.TabCaption(TAB_EDITOR) = GetLbl(lblEditor)
    tabPost.TabCaption(TAB_MORE) = GetLbl(lblMore)
    tabPost.TabCaption(TAB_PREVIEW) = GetLbl(lblPreview)
    lblStatus.Caption = GetLbl(lblPost) & ":"
    txtPostTit.tag = GetLbl(lblTitle) & ":"
    cboPostCat.tag = GetLbl(lblCategory) & ":"
    lblExtEntry.Caption = GetLbl(lblExtendedEntry) & ":"
    lblExcEntry.Caption = GetLbl(lblExcerpt) & ":"
    txtKeywords.tag = GetLbl(lblKeywords) & ":"
    cmdMore.Caption = GetLbl(lblAdvanced)
    With acbMain.Bands("bndPopCustom")
        .Tools("miCustomF1").Caption = "? " & GetLbl(lblClickToEdit) & " ?"
        For i = 2 To 12
            .Tools("miCustomF" & i).Caption = .Tools("miCustomF1").Caption
        Next
    End With
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
