VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Begin VB.Form frmRecent 
   Caption         =   "Recent Posts"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Recent.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7785
   Begin SizerOneLibCtl.ElasticOne pnlRecent 
      Height          =   5490
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   7785
      _cx             =   13732
      _cy             =   9684
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
      BorderWidth     =   8
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
      GridRows        =   3
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"Recent.frx":000C
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   390
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4980
         Width           =   2745
         _cx             =   4842
         _cy             =   688
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
         Begin VB.CommandButton cmdCheckNone 
            Height          =   330
            Left            =   510
            Picture         =   "Recent.frx":0073
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Uncheck all"
            Top             =   30
            Width           =   375
         End
         Begin VB.CommandButton cmdCheckAll 
            Height          =   330
            Left            =   75
            Picture         =   "Recent.frx":03CF
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Check all"
            Top             =   30
            Width           =   375
         End
      End
      Begin VB.TextBox txtAuthor 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   2925
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4560
         Width           =   4740
      End
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         Height          =   4380
         Left            =   2925
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   4740
      End
      Begin VB.ListBox lstPostID 
         Height          =   4800
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   120
         Width           =   2745
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   390
         Left            =   3735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4980
         Width           =   3930
         _cx             =   6932
         _cy             =   688
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
         Begin VB.CommandButton cmdButton 
            Caption         =   "&Delete"
            Height          =   375
            Index           =   1
            Left            =   1365
            TabIndex        =   5
            Top             =   15
            Width           =   1230
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "&Select"
            Default         =   -1  'True
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   15
            Width           =   1230
         End
         Begin VB.CommandButton cmdButton 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   375
            Index           =   2
            Left            =   2685
            TabIndex        =   6
            Top             =   15
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "frmRecent"
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
Dim aPosts() As PostData

Public Sub ShowForm(ByVal intPosts As Integer, Optional ByVal bolOpenLast As Boolean = True)
On Error GoTo ErrorHandler
Dim varStruct(), p As Integer, strAux As String
Dim objPost As xmlStruct

    Screen.MousePointer = vbHourglass
    frmPost.Message = GetMsg(msgGettingPosts)
    DoEvents
    varStruct = GetRecentPosts(intPosts)
    If UBound(varStruct) = 0 And bolOpenLast Then
        Set objPost = varStruct(0)
        If gAccount.GetPostsMethod = API_BLOGGER2 Then 'To Wait the correct
            GetPost objPost.Member("postid").Value
        Else
            ReDim aPosts(0)
            Set aPosts(0) = FillPost(objPost)
            If gAccount.GetPostsMethod = API_MT Then
                aPosts(0).Categories = GetMTCategories(aPosts(0).PostID)
            End If
            frmPost.EditPost aPosts(0)
        End If
        frmPost.Message = ""
        Screen.MousePointer = vbDefault
        Set objPost = Nothing
        Unload Me
        Exit Sub
    End If
    lstPostID.Clear
    ReDim aPosts(0)
    For p = 0 To UBound(varStruct)
        Set objPost = varStruct(p)
        ReDim Preserve aPosts(p)
        Set aPosts(p) = FillPost(objPost)
        'Add post to the list
        If SupportsTitle() And Trim(aPosts(p).Title) <> "" Then
            lstPostID.AddItem objPost.Member("postid").Value & " - " & aPosts(p).Title
        Else
            lstPostID.AddItem objPost.Member("postid").Value & " - " & aPosts(p).DateTime
        End If
    Next
    If lstPostID.ListCount > 0 Then
        lstPostID.ListIndex = -1
        lstPostID.ListIndex = 0
    End If
    Set objPost = Nothing
    frmPost.Message = ""
    LocalizeForm
    Screen.MousePointer = vbDefault
    If Not Me.Visible Then
        Me.Icon = frmPost.Icon
        LoadFormSettings Me, , , 7905, 6000
        Me.Show vbModal, frmPost
        Unload Me
    End If
    Exit Sub
ErrorHandler:
    If Err.Number = 9 Then
        MsgBox GetMsg(msgBlogEmpty), vbExclamation
    Else
        ErrorMessage Err.Number, Err.Description, Me.Name
    End If
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim c As Integer, intDel As Integer
    If Index = 0 Then
        If gAccount.GetPostsMethod = API_BLOGGER2 Then 'To Wait the correct
            GetPost aPosts(lstPostID.ListIndex).PostID
        Else
            'If is MovableType get selected post categories
            If gAccount.GetPostsMethod = API_MT Then
                aPosts(lstPostID.ListIndex).Categories = GetMTCategories(aPosts(lstPostID.ListIndex).PostID)
            End If
            'Edit the selected post
            frmPost.EditPost aPosts(lstPostID.ListIndex)
        End If
    ElseIf Index = 1 Then 'Delete checked
        If lstPostID.SelCount > 0 Then
            If MsgBox(GetMsg(msgDelPosts), vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
            For c = 0 To lstPostID.ListCount - 1
                If lstPostID.Selected(c) Then
                    If DeletePost(aPosts(c).PostID, True, True) Then
                        intDel = intDel + 1
                    Else
                        Exit Sub
                    End If
                End If
            Next
            If lstPostID.ListCount = intDel Then
                Unload Me
            Else
                ShowForm lstPostID.ListCount - intDel, False
            End If
        End If
        Exit Sub
    End If
    Me.Hide
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdCheckAll_Click()
Dim c As Long
    If lstPostID.ListCount > 0 Then
        For c = 0 To lstPostID.ListCount - 1
            If Not lstPostID.Selected(c) Then
                lstPostID.Selected(c) = True
            End If
        Next
    End If
End Sub

Private Sub cmdCheckNone_Click()
Dim c As Long
    If lstPostID.ListCount > 0 Then
        For c = 0 To lstPostID.ListCount - 1
            If lstPostID.Selected(c) Then
                lstPostID.Selected(c) = False
            End If
        Next
    End If
End Sub

Private Sub Form_Activate()
    cmdButton(0).Enabled = (lstPostID.ListCount > 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormSettings Me
End Sub

Private Sub lstPostID_Click()
On Error Resume Next
    cmdButton(1).Enabled = lstPostID.SelCount
    If aPosts(lstPostID.ListIndex).Title <> "" Then
        txtView.Text = aPosts(lstPostID.ListIndex).Title & vbCrLf & aPosts(lstPostID.ListIndex).Text
    Else
        txtView.Text = aPosts(lstPostID.ListIndex).Text
    End If
    If aPosts(lstPostID.ListIndex).Author <> "" Then
        txtAuthor.Text = GetLbl(lblPostedBy) & " " & aPosts(lstPostID.ListIndex).Author & " " & Format(aPosts(lstPostID.ListIndex).DateTime, "General Date")
    Else
        txtAuthor.Text = Format(aPosts(lstPostID.ListIndex).DateTime, "Long Date") & " " & Format(aPosts(lstPostID.ListIndex).DateTime, "Short Time")
    End If
    lstPostID.Refresh
End Sub

Private Sub lstPostID_DblClick()
    cmdButton_Click 0
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblRecentPosts)
    cmdButton(0).Caption = GetLbl(lblSelect)
    cmdButton(1).Caption = GetLbl(lblDelete)
    cmdButton(2).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
