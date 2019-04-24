VERSION 5.00
Begin VB.Form frmMultipost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post to Many Blogs"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Multipost.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPublish 
      Appearance      =   0  'Flat
      Caption         =   "Publish"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   1575
      Value           =   1  'Checked
      Width           =   2265
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Post"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   540
      TabIndex        =   2
      Top             =   1860
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   1860
      Width           =   1155
   End
   Begin VB.ListBox lstBlogs 
      Height          =   1410
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   105
      Width           =   2850
   End
End
Attribute VB_Name = "frmMultipost"
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
Dim b As Integer, bolPosted As Boolean, strPostID As String
    If Index = 0 Then
        For b = 0 To UBound(gBlogs)
            If lstBlogs.Selected(b) Then
                frmPost.Message = GetMsg(msgPostingTo) & " " & lstBlogs.List(b) & "..."
                bolPosted = False
                Do While Not bolPosted
                    bolPosted = Post(frmPost.txtPostTit.Text, _
                                     frmPost.txtPost.Text, _
                                     frmPost.txtMore.Text, _
                                     frmPost.txtExcerpt.Text, _
                                     frmPost.txtKeywords.Text, _
                                     Array(), gBlogs(b).BlogID, _
                                     gPostID, chkPublish.Value, True, _
                                     frmPost.PostData.AllowComments, _
                                     frmPost.PostData.AllowPings, _
                                     frmPost.PostData.DateTime, _
                                     frmPost.PostData.TrackBack, _
                                     frmPost.PostData.TextFilter)
                    If Not bolPosted Then
                        Select Case MsgBox(GetMsg(msgErrorPosting) & " " & lstBlogs.List(b) & vbCrLf & _
                                           GetMsg(msgChooseAction), vbAbortRetryIgnore + vbDefaultButton2 + vbExclamation)
                        Case vbAbort
                            Unload Me
                            Exit Sub
                        Case vbRetry
                            'Just do it again
                        Case vbIgnore
                            Exit Do
                        End Select
                    ElseIf b = frmPost.CurrentBlog Then
                        strPostID = gPostID
                    End If
                    gPostID = ""
                Loop
                lstBlogs.List(b) = lstBlogs.List(b) & vbTab & Chr(149)
            End If
        Next
        If gSettings.ClearPost Then
            frmPost.NewPost
        Else
            gPostID = strPostID
            If gPostID <> "main" And _
               gPostID <> "archiveIndex" Then
                frmPost.imgStatus.Picture = frmPost.acbMain.Tools("miPosts").GetPicture(0)
                frmPost.lblStatus.Caption = GetLbl(lblPost) & ": " & gPostID
                'frmPost.SaveDraftPost
            End If
        End If
        frmPost.Message = ""
        Unload Me
    Else
        Unload Me
    End If
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub Form_Load()
Dim b As Integer
    LocalizeForm
    For b = 0 To UBound(gBlogs)
        lstBlogs.AddItem gBlogs(b).Name
    Next
End Sub

Private Sub lstBlogs_Click()
    cmdButton(0).Enabled = lstBlogs.SelCount
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblPostMany)
    cmdButton(0).Caption = frmPost.acbMain.Tools("miPost").Caption
    cmdButton(1).Caption = GetLbl(lblCancel)
    chkPublish.Caption = GetLbl(lblPublish)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
