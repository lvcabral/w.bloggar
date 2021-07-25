VERSION 5.00
Object = "{ED442B9F-ADE2-11D4-B868-00606E3BC2C9}#1.0#0"; "ActiveCbo.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "w.bloggar"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLogin 
      Caption         =   "Login"
      Height          =   1515
      Left            =   90
      TabIndex        =   10
      Top             =   3420
      Width           =   3015
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1260
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   300
         Width           =   1635
      End
      Begin VB.CheckBox chkSavePwd 
         Appearance      =   0  'Flat
         Caption         =   "&Save Password"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         Caption         =   "&User:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   390
      End
   End
   Begin rdActiveCombo.ActiveCombo cboAccount 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   3000
      Width           =   2355
      _ExtentX        =   4154
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
   Begin VB.CommandButton cmdAccount 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2610
      Picture         =   "Login.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   5040
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1950
      TabIndex        =   9
      Top             =   5040
      Width           =   1155
   End
   Begin VB.Image imgCustom 
      Height          =   240
      Left            =   120
      Picture         =   "Login.frx":1DA7
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Account:"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   0
      Top             =   2760
      Width           =   645
   End
   Begin VB.Image imgLogin 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2565
      Left            =   90
      Picture         =   "Login.frx":2331
      Stretch         =   -1  'True
      Top             =   90
      Width           =   3015
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New Account..."
      End
   End
End
Attribute VB_Name = "frmLogin"
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
Private intSaveAcc As Integer
Private strSavePwd As String
Private intNextAcc As Integer

Private Sub cboAccount_Click()
On Error GoTo ErrorHandler
    If cboAccount.ListIndex >= 0 Then
        gAccount.Current = cboAccount.ItemData(cboAccount.ListIndex)
        LoadAccount
        txtUser.Text = gAccount.User
        txtPassword.Text = gAccount.Password
        If txtPassword.Text = "" Then
            If Me.Visible Then txtPassword.SetFocus
        End If
        chkSavePwd.Value = Abs(gAccount.SavePassword)
        mnuEdit.Enabled = True
        mnuDelete.Enabled = (gAccount.Current <> intSaveAcc)
    End If
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub cmdAccount_Click()
    PopupMenu mnuPopUp, , cmdAccount.Left, cmdAccount.Top + cmdAccount.Height
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
    If Index = 0 Then 'OK
        If Trim(txtUser.Text) = "" Then
            MsgBox GetMsg(msgEnterUser), vbInformation
            txtUser.SetFocus
            Exit Sub
        ElseIf Trim(txtPassword.Text) = "" Then
            MsgBox GetMsg(msgEnterPassword), vbInformation
            txtPassword.SetFocus
            Exit Sub
        End If
        gAccount.User = txtUser.Text
        gAccount.Password = txtPassword.Text
        gAccount.SavePassword = chkSavePwd.Value
        If gAccount.CMS <> CMS_CUSTOM Then LoadCMS
        If LoadBlogs(False) Then
            SaveAccount
            frmPost.Account = gAccount.Alias
            LoadCategories False
        Else
            gAccount.User = ""
            gAccount.Password = ""
            Exit Sub
        End If
    Else 'Cancel
        RestoreAccount
    End If
    Unload Me
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If txtUser.Text = "" Then
        txtUser.SetFocus
    ElseIf txtPassword.Text = "" Then
        txtPassword.SetFocus
    Else
        cmdButton(0).SetFocus
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim strAux As String
    LoadAccounts
    If gAccount.User <> "" Then
        'Save Current Account Data
        intSaveAcc = gAccount.Current
        strSavePwd = gAccount.Password
        'Select Current Account
        SearchItemData cboAccount, gAccount.Current
        txtPassword = strSavePwd
    Else
        intSaveAcc = -1
        strSavePwd = ""
        cboAccount.ListIndex = -1
    End If
    'Load Skin Icons and Bitmap
    strAux = ReadINI("Skin", "LoginImage", gSettings.SkinFolder & "\skin.ini")
    If FileExists(gSettings.SkinFolder & "\" & strAux) Then
        imgLogin.Picture = LoadPicture(gSettings.SkinFolder & "\" & strAux)
    End If
    strAux = ReadINI("Skin", "IconsExt", gSettings.SkinFolder & "\skin.ini")
    If FileExists(gSettings.SkinFolder & "\AccountProp." & strAux) Then
        cmdAccount.Picture = LoadPicture(gSettings.SkinFolder & "\AccountProp." & strAux, vbLPSmall)
    End If
    'Translate
    LocalizeForm
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub LoadAccounts()
On Error GoTo ErrorHandler
Dim strReg As String, strUser As String, a As Integer
Dim strIco As String, strCMS As String
Dim colIco As New Collection

    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    cboAccount.Clear
    Set cboAccount.ImageList = Nothing
    intNextAcc = -1
    For a = 0 To 99
        strReg = "Accounts/a" & Format(a, "00")
        strUser = objXMLReg.GetSetting(App.Title, strReg, "User", "*")
        If strUser = "*" Then Exit For
        If Not CBool(objXMLReg.GetSetting(App.Title, strReg, "Deleted", "0")) Then
            strCMS = objXMLReg.GetSetting(App.Title, strReg, "CMS", "0")
            strIco = App.Path & "\CMS\" & ReadINI("CMS-" & Format(strCMS, "00"), "Icon", App.Path & "\CMS\CMS.ini")
            If FileExists(strIco) Then
                On Error Resume Next
                colIco.Add colIco.Count + 1, strCMS
                If Err.Number = 0 Then
                    cboAccount.AddIcon LoadPicture(strIco)
                End If
            Else
                On Error Resume Next
                colIco.Add colIco.Count + 1, strCMS
                If Err.Number = 0 Then
                    cboAccount.AddIcon imgCustom.Picture
                End If
            End If
            Err = 0
            cboAccount.AddItem objXMLReg.GetSetting(App.Title, strReg, "Alias"), , colIco(strCMS)
            If Err.Number <> 0 Then
                cboAccount.AddItem objXMLReg.GetSetting(App.Title, strReg, "Alias")
            End If
            cboAccount.ItemData(cboAccount.NewIndex) = a
        ElseIf intNextAcc < 0 Then
            intNextAcc = a
        End If
    Next
    If intNextAcc < 0 Then intNextAcc = a
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then RestoreAccount
End Sub

Private Sub mnuDelete_Click()
    cboAccount.SetFocus
    If MsgBox(GetMsg(msgDelAccount), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        DeleteAccount
        LoadAccounts
        SearchItemData cboAccount, intSaveAcc
    End If
End Sub

Private Sub mnuEdit_Click()
    cboAccount.SetFocus
    frmAccount.Show vbModal, Me
    If cboAccount.Text <> gAccount.Alias Then
        Dim i As Long
        i = cboAccount.ListIndex
        LoadAccounts
        cboAccount.ListIndex = i
    End If
End Sub

Private Sub mnuNew_Click()
    frmAccountWiz.Show vbModal, Me
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser.Text)
End Sub

Private Sub RestoreAccount()
    'Restore Current Account Settings
    gAccount.Current = intSaveAcc
    LoadAccount
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    Call objXMLReg.SaveSetting(App.Title, "Settings", "Account", Format(gAccount.Current))
    Set objXMLReg = Nothing
    gAccount.Password = strSavePwd
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    lblField(0).Caption = GetLbl(lblUser) & ":"
    lblField(1).Caption = GetLbl(lblPassword) & ":"
    lblField(2).Caption = GetLbl(lblAccount) & ":"
    
    cmdAccount.ToolTipText = GetLbl(lblAccount)
    mnuEdit.Caption = GetLbl(lblConnection) & "..."
    mnuDelete.Caption = GetLbl(lblDeleteAccount)
    mnuNew.Caption = GetLbl(lblNew) & "..."
    
    chkSavePwd.Caption = GetLbl(lblSavePassword)
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub

