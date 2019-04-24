VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Begin VB.Form frmBlgLocal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "w.bloggar Localizer :: v4.x ::"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BlgLocal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLCID 
      Height          =   315
      Left            =   5775
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   105
      Width           =   2310
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 barTranslate 
      Height          =   465
      Left            =   8415
      TabIndex        =   14
      Top             =   4065
      Visible         =   0   'False
      Width           =   960
      _LayoutVersion  =   1
      _ExtentX        =   1693
      _ExtentY        =   820
      _DataPath       =   ""
      Bands           =   "BlgLocal.frx":058A
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 barEnglish 
      Height          =   465
      Left            =   8415
      TabIndex        =   13
      Top             =   3555
      Visible         =   0   'False
      Width           =   960
      _LayoutVersion  =   1
      _ExtentX        =   1693
      _ExtentY        =   820
      _DataPath       =   ""
      Bands           =   "BlgLocal.frx":0752
   End
   Begin VB.ListBox lstTranslate 
      Height          =   4545
      Left            =   4140
      TabIndex        =   0
      Top             =   495
      Width           =   3975
   End
   Begin VB.CommandButton cmdAction 
      Height          =   330
      Index           =   1
      Left            =   7755
      Picture         =   "BlgLocal.frx":091A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Reload English Entry"
      Top             =   5130
      Width           =   345
   End
   Begin VB.CommandButton cmdAction 
      Height          =   330
      Index           =   0
      Left            =   7380
      Picture         =   "BlgLocal.frx":0EA4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Translate"
      Top             =   5130
      Width           =   345
   End
   Begin VB.TextBox txtTranslate 
      Height          =   330
      Left            =   4140
      TabIndex        =   1
      Top             =   5130
      Width           =   3195
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Help"
      Height          =   405
      Index           =   2
      Left            =   8235
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Close"
      Height          =   405
      Index           =   3
      Left            =   8250
      TabIndex        =   7
      Top             =   1950
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Reload English"
      Height          =   405
      Index           =   1
      Left            =   8235
      TabIndex        =   5
      Top             =   675
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Save Language"
      Height          =   405
      Index           =   0
      Left            =   8235
      TabIndex        =   4
      Top             =   135
      Width           =   1335
   End
   Begin VB.ListBox lstEnglish 
      Enabled         =   0   'False
      Height          =   4545
      Left            =   105
      TabIndex        =   8
      Top             =   495
      Width           =   3975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localized Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   4275
      TabIndex        =   12
      Top             =   135
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000C&
      Height          =   375
      Left            =   4140
      TabIndex        =   11
      Top             =   75
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "English Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   225
      TabIndex        =   10
      Top             =   135
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Height          =   375
      Left            =   105
      TabIndex        =   9
      Top             =   75
      Width           =   3975
   End
End
Attribute VB_Name = "frmBlgLocal"
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

Private Sub cboLCID_Click()
On Error Resume Next
Dim tooLoop As Tool
Dim strTool As String, lngLast As Long
    If Not Me.Visible Then
        Exit Sub
    ElseIf cboLCID.ListIndex <= 0 Then
        MsgBox "You must select a Locale ID to this translation file!" & vbCrLf & _
               "Windows reports your current Locale ID as : " & GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE) & " - " & LCID, vbInformation
        cboLCID.SetFocus
        cmdButton(0).Enabled = False
        cmdButton(1).Enabled = False
        lstTranslate.Enabled = False
        txtTranslate.Enabled = False
        cmdAction(0).Enabled = False
        cmdAction(1).Enabled = False
        Exit Sub
    End If
    'Load Localization File
    If FileExists(App.Path & "\" & Format(cboLCID.ItemData(cboLCID.ListIndex)) & ".lng") Then
        barTranslate.Load "", App.Path & "\" & Format(cboLCID.ItemData(cboLCID.ListIndex)) & ".lng", ddSOFile
        lstTranslate.Clear
        For Each tooLoop In barEnglish.Tools
            Err = 0
            tooLoop.Enabled = False
            If Val(tooLoop.ID) < 30000 Then
                If Val(tooLoop.ID) > 10000 Then
                    strTool = barTranslate.Tools(Val(tooLoop.ID)).Caption
                    If Err = 0 Then
                        lstTranslate.AddItem strTool
                    Else
                        lstTranslate.AddItem ">>"
                    End If
                    If tooLoop.ToolTipText <> "" Then
                        strTool = barTranslate.Tools(Val(tooLoop.ID)).ToolTipText
                        If Err = 0 Then
                            lstTranslate.AddItem strTool
                        Else
                            lstTranslate.AddItem ">>"
                        End If
                    End If
                    lngLast = tooLoop.ID
                End If
            ElseIf tooLoop.Description <> "" Then
                strTool = barTranslate.Tools(Val(tooLoop.ID)).Description
                If Err = 0 Then
                    lstTranslate.AddItem strTool
                Else
                    lstTranslate.AddItem ">>"
                End If
            Else
                strTool = barTranslate.Tools(Val(tooLoop.ID)).Caption
                If Err = 0 Then
                    lstTranslate.AddItem strTool
                Else
                    lstTranslate.AddItem ">>"
                End If
            End If
        Next
        SearchItemData cboLCID, barTranslate.Tools("LocaleID").Caption
        barTranslate.Load "", barEnglish.Save("", "", ddSOByteArray), ddSOByteArray
    Else
        barTranslate.Load "", barEnglish.Save("", "", ddSOByteArray), ddSOByteArray
        lstTranslate.Clear
        For Each tooLoop In barTranslate.Tools
            tooLoop.Enabled = False
            If Val(tooLoop.ID) < 30000 Then
                If Val(tooLoop.ID) > 10000 Then
                    lstTranslate.AddItem tooLoop.Caption
                    If tooLoop.ToolTipText <> "" Then
                        lstTranslate.AddItem tooLoop.ToolTipText
                    End If
                End If
            ElseIf tooLoop.Description <> "" Then
                lstTranslate.AddItem tooLoop.Description
            Else
                lstTranslate.AddItem tooLoop.Caption
            End If
        Next
    End If
    cmdButton(0).Enabled = True
    cmdButton(1).Enabled = True
    lstTranslate.Enabled = True
    txtTranslate.Enabled = True
    cmdAction(0).Enabled = True
    cmdAction(1).Enabled = True
    If lstTranslate.ListCount > 0 Then
        lstTranslate.ListIndex = 0
    End If
End Sub

Private Sub cmdAction_Click(Index As Integer)
    If Index = 0 Then
        If txtTranslate.Text = "" Then
            MsgBox "Please, enter the translated String!", vbInformation
            Exit Sub
        End If
        lstTranslate.List(lstTranslate.ListIndex) = txtTranslate.Text
        cmdButton(0).Enabled = True
    Else
        lstTranslate.List(lstTranslate.ListIndex) = lstEnglish.List(lstTranslate.ListIndex)
        txtTranslate.Text = lstEnglish.List(lstTranslate.ListIndex)
    End If
    If lstTranslate.ListIndex < (lstTranslate.ListCount - 1) Then
        lstTranslate.ListIndex = lstTranslate.ListIndex + 1
        lstTranslate.SetFocus
        txtTranslate.SetFocus
    End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim bndLoop As Band, tooLoop As Tool, tooBand As Tool
Dim i As Integer, strHelp As String
    Select Case Index
    Case 0
        If cboLCID.ListIndex <= 0 Then
            MsgBox "You must select a Locale ID to this translation file!" & vbCrLf & _
                   "Windows reports your current Locale ID as : " & GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE) & " - " & LCID, vbInformation
            cboLCID.SetFocus
            Exit Sub
        End If
        i = 0
        For Each tooLoop In barTranslate.Tools
            tooLoop.Enabled = True
            If Val(tooLoop.ID) < 30000 Then
                If Val(tooLoop.ID) > 10000 Then
                    tooLoop.Caption = lstTranslate.List(i)
                    i = i + 1
                    If tooLoop.ToolTipText <> "" Then
                        tooLoop.ToolTipText = lstTranslate.List(i)
                        i = i + 1
                    End If
                End If
            ElseIf tooLoop.Description <> "" Then
                tooLoop.Description = lstTranslate.List(i)
                i = i + 1
            Else
                tooLoop.Caption = lstTranslate.List(i)
                i = i + 1
            End If
            tooLoop.SetPicture ddITNormal, LoadPicture()
            For Each bndLoop In barTranslate.Bands
                For Each tooBand In bndLoop.Tools
                    If tooBand.ID = tooLoop.ID Then
                        tooBand.Caption = tooLoop.Caption
                        tooBand.ToolTipText = tooLoop.ToolTipText
                        tooBand.Description = tooLoop.Description
                        If bndLoop.Name <> "bndPopPosts" And _
                           bndLoop.Name <> "bndPopView" And _
                           Left(tooBand.Name, 6) <> "miFMRU" Then
                            tooBand.SetPicture ddITNormal, LoadPicture()
                        End If
                    End If
                Next
            Next
        Next
        barTranslate.Tools("LocaleID").Caption = Format(cboLCID.ItemData(cboLCID.ListIndex))
        If FileExists(App.Path & "\" & Format(cboLCID.ItemData(cboLCID.ListIndex)) & ".lng") Then
            If MsgBox("This language file already exists, do you wan't to overwrite it?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
        barTranslate.Save "", App.Path & "\" & Format(cboLCID.ItemData(cboLCID.ListIndex)) & ".lng", ddSOFile
'        If FileExists(App.Path & "\wbloggar.chg") Then
'            Call Kill(App.Path & "\wbloggar.chg")
'        End If
        cmdButton(0).Enabled = False
        MsgBox "The " & Format(cboLCID.ItemData(cboLCID.ListIndex)) & ".lng file was saved successfully!", vbInformation
    Case 1
        If MsgBox("Warning! This will undo all translations made, are you sure?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
            lstTranslate.Clear
            For Each tooLoop In barEnglish.Tools
                If Val(tooLoop.ID) < 30000 Then
                    If Val(tooLoop.ID) > 10000 Then
                        lstTranslate.AddItem tooLoop.Caption
                        If tooLoop.ToolTipText <> "" Then
                            lstTranslate.AddItem tooLoop.ToolTipText
                        End If
                    End If
                ElseIf tooLoop.Description <> "" Then
                    lstTranslate.AddItem tooLoop.Description
                Else
                    lstTranslate.AddItem tooLoop.Caption
                End If
            Next
            If lstTranslate.ListCount > 0 Then
                lstTranslate.ListIndex = 0
            End If
            cmdButton(0).Enabled = True
        End If
    Case 2
        strHelp = "Select an entry and translate it in the TextBox, " & _
                  "press Enter or click on Check Button to confirm." & vbCrLf & vbCrLf & _
                  "You can undo to english one entry or full list with the " & _
                  "correspondent buttons." & vbCrLf & vbCrLf & _
                  "When you click Save Language button a file named [Locale ID].lng is saved on disk," & vbCrLf & _
                  "you can make intermediate saves." & vbCrLf & vbCrLf & _
                  "To test the generated file, select the Locale ID " & vbCrLf & _
                  "at w.bloggar Options window and restart it."
        MsgBox strHelp, vbInformation
    Case 3
        Unload Me
    End Select
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err & " - " & Error$, vbCritical
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtTranslate.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim tooLoop As Tool
Dim i As Integer, t As Integer
Dim strTool As String, lngLast As Long
    'get the user's current default ID
    LCID = GetUserDefaultLCID()
    'Populate LCID ComboBox
    cboLCID.AddItem ">> Select the Locale ID"
    Call EnumSystemLocales(AddressOf EnumSystemLocalesProc, LCID_SUPPORTED)
    cboLCID.ListIndex = 0
    'Read English Text
    barEnglish.Load "", App.Path + "\wbloggar.len", ddSOFile
    For Each tooLoop In barEnglish.Tools
        tooLoop.Enabled = False
        If Val(tooLoop.ID) < 30000 Then
            If Val(tooLoop.ID) > 10000 Then
                lstEnglish.AddItem tooLoop.Caption
                lstEnglish.ItemData(lstEnglish.NewIndex) = tooLoop.ID
                If tooLoop.ToolTipText <> "" Then
                    lstEnglish.AddItem tooLoop.ToolTipText
                    lstEnglish.ItemData(lstEnglish.NewIndex) = tooLoop.ID
                End If
            End If
        ElseIf tooLoop.Description <> "" Then
            lstEnglish.AddItem tooLoop.Description
            lstEnglish.ItemData(lstEnglish.NewIndex) = tooLoop.ID
        Else
            lstEnglish.AddItem tooLoop.Caption
            lstEnglish.ItemData(lstEnglish.NewIndex) = tooLoop.ID
        End If
    Next
    cmdButton(0).Enabled = False
    cmdButton(1).Enabled = False
    lstTranslate.Enabled = False
    txtTranslate.Enabled = False
    cmdAction(0).Enabled = False
    cmdAction(1).Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <= vbFormCode And cmdButton(0).Enabled Then
        Select Case MsgBox("There are unsaved changes on the translation file, do you want to save before closing?", vbExclamation + vbYesNoCancel)
        Case vbYes
            cmdButton_Click 0
        Case vbNo
            Cancel = False
        Case vbCancel
            Cancel = True
        End Select
    End If
End Sub

Private Sub lstTranslate_Click()
    lstEnglish.TopIndex = lstTranslate.TopIndex
    lstEnglish.ListIndex = lstTranslate.ListIndex
    If lstTranslate.Text = ">>" Then
        txtTranslate.Text = lstEnglish.Text
    Else
        txtTranslate.Text = lstTranslate.Text
    End If
End Sub

Private Sub txtTranslate_GotFocus()
    txtTranslate.SelStart = 0
    txtTranslate.SelLength = Len(txtTranslate.Text)
End Sub

Private Sub txtTranslate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       cmdAction_Click 0
       KeyAscii = 0
    End If
End Sub
