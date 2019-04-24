VERSION 5.00
Begin VB.Form frmUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload File"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Upload.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3735
      TabIndex        =   8
      Top             =   2310
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Upload"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2505
      TabIndex        =   7
      Top             =   2310
      Width           =   1155
   End
   Begin VB.Frame fraUpload 
      Caption         =   "After the upload insert:"
      Height          =   1530
      Left            =   165
      TabIndex        =   3
      Top             =   645
      Width           =   4725
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   1185
         Left            =   75
         ScaleHeight     =   1185
         ScaleWidth      =   4575
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   4575
         Begin VB.CheckBox chkEditTag 
            Caption         =   "Edit the Tag before insert"
            Height          =   270
            Left            =   30
            TabIndex        =   6
            Top             =   840
            Width           =   4500
         End
         Begin VB.OptionButton optLink 
            Appearance      =   0  'Flat
            Caption         =   "Link to the file (<a href=""..."">filename.ext</a>)"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   30
            TabIndex        =   5
            Top             =   465
            Width           =   4545
         End
         Begin VB.OptionButton optLink 
            Appearance      =   0  'Flat
            Caption         =   "Image on the post (<img src=""..."">)"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   30
            TabIndex        =   4
            Top             =   90
            Value           =   -1  'True
            Width           =   4545
         End
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   330
      Left            =   4530
      TabIndex        =   2
      Top             =   135
      Width           =   360
   End
   Begin VB.TextBox txtFile 
      Height          =   345
      Left            =   1125
      TabIndex        =   1
      Top             =   135
      Width           =   3375
   End
   Begin VB.Image imgSize 
      Height          =   375
      Left            =   390
      Top             =   2280
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&File:"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   300
   End
End
Attribute VB_Name = "frmUpload"
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
Public HtmlTag As String

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim objFTP As clsFTP, strSize As String, strURL As String
    If Index = 0 Then 'OK
        If Not FileExists(txtFile.Text) Then
            MsgBox GetMsg(msgFileNotFound), vbExclamation
            txtFile.SetFocus
            Exit Sub
        End If
        If gBlog.APIUpload Then
            strURL = UploadMediaObject(gBlogs(frmPost.CurrentBlog).BlogID, txtFile.Text)
        Else
           strURL = UploadFTP()
        End If
        If Trim(strURL) = "" Then GoTo ExitNow
        If chkEditTag.Value = vbUnchecked Then
            If optLink(0).Value Then
                If InStr(1, txtFile.Text, ".jpg", vbTextCompare) > 0 Or _
                   InStr(1, txtFile.Text, ".gif", vbTextCompare) > 0 Then
                    On Error Resume Next
                    imgSize.Picture = LoadPicture(txtFile.Text)
                    If Err.Number = 0 Then
                        strSize = " width=""" & imgSize.Width & """ height=""" & imgSize.Height & """"
                        imgSize.Picture = LoadPicture()
                    End If
                End If
                If gSettings.XHTML Then
                    HtmlTag = "<img src=""" & strURL & """" & strSize & " alt="""" border=""0"" />"
                Else
                    HtmlTag = "<img src=""" & strURL & """" & strSize & " alt="""" border=""0"">"
                End If
            Else
                HtmlTag = "<a href=""" & strURL & """ target=""_blank"">" & GetNamePart(txtFile.Text) & "</a>"
            End If
        Else
            If optLink(0).Value Then
                frmImage.cboImage.Text = strURL
                On Error Resume Next
                imgSize.Picture = LoadPicture(txtFile.Text)
                If Err.Number = 0 Then
                    frmImage.txtWidth = imgSize.Width
                    frmImage.txtHeight = imgSize.Height
                    imgSize.Picture = LoadPicture()
                End If
                HtmlTag = "I"
            Else
                frmLink.cboURL.Text = strURL
                frmLink.txtText = GetNamePart(txtFile.Text)
                HtmlTag = "L"
            End If
        End If
        Set objXMLReg = New XMLRegistry
        objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
        objXMLReg.SaveSetting App.Title, "Settings", "UploadEditTag", Format(chkEditTag.Value)
        Set objXMLReg = Nothing
    End If
ExitNow:
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Me.Hide
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
    Resume ExitNow
End Sub

Private Sub cmdOpen_Click()
On Error Resume Next
Dim oFile As New FileDialog
    oFile.DialogTitle = cmdOpen.ToolTipText
    oFile.Filter = GetMsg(msgFileFilterAll)
    oFile.Flags = cdlFileMustExist Or cdlLongnames Or cdlHideReadOnly
    oFile.hWndParent = Me.hwnd
    oFile.ShowOpen
    If oFile.FileName <> "" Then
        txtFile.Text = oFile.FileName
    End If
    Set oFile = Nothing
End Sub

Private Sub Form_Load()
On Error Resume Next
    LocalizeForm
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    chkEditTag.Value = Val(objXMLReg.GetSetting(App.Title, "Settings", "UploadEditTag", "0"))
    Set objXMLReg = Nothing
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = frmPost.acbMain.Tools("miUpload").ToolTipText
    lblField.Caption = GetLbl(lblFile) & ":"
    fraUpload.Caption = GetLbl(lblAfterUpload) & ":"
    optLink(0).Caption = GetLbl(lblImageOnPost)
    optLink(1).Caption = GetLbl(lblLinkToFile)
    chkEditTag.Caption = GetLbl(lblEditTag)
    cmdButton(0).Caption = GetLbl(lblUpload)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub

Private Function UploadFTP() As String
On Error GoTo ErrorHandler
Dim objFTP As clsFTP, strSize As String
    Screen.MousePointer = vbHourglass
    frmPost.Message = GetMsg(msgStartUpload)
    DoEvents
    Set objFTP = New clsFTP
    objFTP.Server = gBlog.FTPHost
    objFTP.Port = gBlog.FTPPort
    If gBlog.FTPProxy Then
        objFTP.UseProxy = True
        If gAccount.UseProxy = 2 Then
            objFTP.ProxyString = gAccount.ProxyServer & ":" & gAccount.ProxyPort
        End If
    Else
        objFTP.UseProxy = False
    End If
    objFTP.UserName = gBlog.FTPUser
    objFTP.Password = gBlog.FTPPassword
    objFTP.TransferType = eftBinary
    frmPost.Message = GetMsg(msgOpenSession)
    DoEvents
    If objFTP.OpenSession() = False Then
        MsgBox GetMsg(msgNotOpenSession) & vbCrLf & objFTP.LastDLLErrorMsg, vbExclamation, App.Title & ": " & objFTP.LastDLLError
        GoTo ExitNow
    End If
    frmPost.Message = GetMsg(msgConnecting)
    DoEvents
    If objFTP.Connect() = False Then
        MsgBox GetMsg(msgNotConnect) & vbCrLf & objFTP.LastDLLErrorMsg, vbExclamation, App.Title & ": " & objFTP.LastDLLError
        GoTo ExitNow
    End If
    frmPost.Message = GetMsg(msgSendingFile)
    DoEvents
    If objFTP.PutFile(txtFile.Text, gBlog.FTPPath & GetNamePart(txtFile.Text)) = False Then
        MsgBox GetMsg(msgNotSent) & vbCrLf & objFTP.LastDLLErrorMsg, vbExclamation, App.Title & ": " & objFTP.LastDLLError
        GoTo ExitNow
    End If
    frmPost.Message = GetMsg(msgDisconnecting)
    DoEvents
    objFTP.Disconnect
    objFTP.CloseSession
    UploadFTP = gBlog.FTPLink & GetNamePart(txtFile.Text)
ExitNow:
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
    Resume ExitNow
End Function
