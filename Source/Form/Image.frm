VERSION 5.00
Begin VB.Form frmImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Image"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Image.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHSpace 
      Height          =   315
      Left            =   4860
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1410
      Width           =   525
   End
   Begin VB.TextBox txtVSpace 
      Height          =   315
      Left            =   6285
      MaxLength       =   5
      TabIndex        =   17
      Top             =   1410
      Width           =   525
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Top             =   1380
      Width           =   2175
   End
   Begin VB.TextBox txtBorder 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3540
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "1"
      Top             =   960
      Width           =   330
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   6285
      MaxLength       =   5
      TabIndex        =   11
      Top             =   960
      Width           =   525
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   4860
      MaxLength       =   5
      TabIndex        =   9
      Top             =   960
      Width           =   525
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   5685
      TabIndex        =   19
      Top             =   1860
      Width           =   1155
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4455
      TabIndex        =   18
      Top             =   1860
      Width           =   1155
   End
   Begin VB.TextBox txtToolTip 
      Height          =   315
      Left            =   1200
      MaxLength       =   255
      TabIndex        =   3
      Top             =   540
      Width           =   5610
   End
   Begin VB.ComboBox cboAlign 
      Height          =   315
      ItemData        =   "Image.frx":000C
      Left            =   1200
      List            =   "Image.frx":0025
      TabIndex        =   5
      Top             =   960
      Width           =   1350
   End
   Begin VB.ComboBox cboImage 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5610
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "H&Space:"
      Height          =   195
      Index           =   8
      Left            =   4140
      TabIndex        =   14
      Top             =   1470
      Width           =   600
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&VSpace:"
      Height          =   195
      Index           =   7
      Left            =   5640
      TabIndex        =   16
      Top             =   1485
      Width           =   585
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Cl&ass:"
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   12
      Top             =   1425
      Width           =   435
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Border:"
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   6
      Top             =   1020
      Width           =   540
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Height:"
      Height          =   195
      Index           =   4
      Left            =   5700
      TabIndex        =   10
      Top             =   1020
      Width           =   525
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&Width:"
      Height          =   195
      Index           =   3
      Left            =   4260
      TabIndex        =   8
      Top             =   1020
      Width           =   480
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Align:"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   1005
      Width           =   405
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Title:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   360
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Image:"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   510
   End
End
Attribute VB_Name = "frmImage"
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
Private aAlign() As String

Private Sub cboAlign_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboAlign, KeyAscii
End Sub

Private Sub cboImage_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboImage, KeyAscii
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim i As Integer
    If Index = 0 Then
        If Trim(cboImage.Text) = "" Or Trim(cboImage.Text) = "http://" Then
            MsgBox GetMsg(msgEnterImage), vbInformation
            cboImage.SetFocus
            cboImage.SelStart = Len(cboImage.Text)
            Exit Sub
        End If
        HtmlTag = "<img src=""" & Trim(cboImage.Text) & """"
        If Trim(cboClass.Text) <> "" Then
            HtmlTag = HtmlTag & "class=""" & cboClass.Text & """ "
        End If
        If Trim(cboAlign.Text) <> "" Then
            If cboAlign.ListIndex >= 0 Then
                HtmlTag = HtmlTag & " align=""" & Trim(LCase(aAlign(cboAlign.ListIndex))) & """"
            Else
                HtmlTag = HtmlTag & " align=""" & Trim(LCase(cboAlign.Text)) & """"
            End If
        End If
        If gSettings.XHTML Then
            HtmlTag = HtmlTag & " alt="""" title=""" & txtToolTip.Text & """"
        Else
            HtmlTag = HtmlTag & " title=""" & txtToolTip.Text & """"
        End If
        If Trim(txtWidth.Text) <> "" Then
            HtmlTag = HtmlTag & " width=""" & txtWidth.Text & """"
        End If
        If Trim(txtHeight.Text) <> "" Then
            HtmlTag = HtmlTag & " height=""" & txtHeight.Text & """"
        End If
        If Trim(txtBorder.Text) <> "" Then
            HtmlTag = HtmlTag & " border=""" & txtBorder.Text & """"
        End If
        If Trim(txtHSpace.Text) <> "" Then
            HtmlTag = HtmlTag & " hspace=""" & txtHSpace.Text & """"
        End If
        If Trim(txtVSpace.Text) <> "" Then
            HtmlTag = HtmlTag & " vspace=""" & txtVSpace.Text & """"
        End If
        If gSettings.XHTML Then
            HtmlTag = HtmlTag & " />"
        Else
            HtmlTag = HtmlTag & ">"
        End If
        Set objXMLReg = New XMLRegistry
        objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
        SaveCombo cboImage, "Images", "Image"
        If Trim(cboClass.Text) <> "" Then
            SaveCombo cboClass, "Classes", "Class"
        End If
        objXMLReg.SaveSetting App.Title, "Images", "Border", txtBorder.Text
        objXMLReg.SaveSetting App.Title, "Images", "HSpace", txtHSpace.Text
        objXMLReg.SaveSetting App.Title, "Images", "VSpace", txtVSpace.Text
        Set objXMLReg = Nothing
    Else
        HtmlTag = ""
    End If
    Me.Hide
End Sub

Private Sub Form_Activate()
    cboImage.SelStart = Len(cboImage.Text)
End Sub

Private Sub Form_Load()
Dim i As Integer, strAux As String
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    'Save Align Options
    ReDim aAlign(cboAlign.ListCount - 1)
    For i = 0 To UBound(aAlign)
        aAlign(i) = cboAlign.List(i)
    Next
    LocalizeForm
    HtmlTag = ""
    'Load Images
    For i = 0 To 30
        strAux = objXMLReg.GetSetting(App.Title, "Images", "Image" & Format(i, "00"), "**")
        If strAux = "**" Then
            Exit For
        Else
            cboImage.AddItem strAux
        End If
    Next
    'Load Classes
    For i = 0 To 30
        strAux = objXMLReg.GetSetting(App.Title, "Classes", "Class" & Format(i, "00"), "**")
        If strAux = "**" Then
            Exit For
        Else
            cboClass.AddItem strAux
        End If
    Next
    txtBorder.Text = objXMLReg.GetSetting(App.Title, "Images", "Border", "0")
    txtHSpace.Text = objXMLReg.GetSetting(App.Title, "Images", "HSpace", "")
    txtVSpace.Text = objXMLReg.GetSetting(App.Title, "Images", "VSpace", "")
    Set objXMLReg = Nothing
End Sub

Private Sub txtBorder_GotFocus()
    AutoSelect txtBorder
End Sub

Private Sub txtBorder_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumbers(KeyAscii)
End Sub

Private Sub txtHeight_GotFocus()
    AutoSelect txtHeight
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumbers(KeyAscii)
End Sub

Private Sub txtHSpace_GotFocus()
    AutoSelect txtHSpace
End Sub

Private Sub txtHSpace_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumbers(KeyAscii)
End Sub

Private Sub txtVSpace_GotFocus()
    AutoSelect txtVSpace
End Sub

Private Sub txtVSpace_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumbers(KeyAscii)
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelect txtWidth
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumbers(KeyAscii)
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = GetLbl(lblInsertImage)
    lblField(0).Caption = GetLbl(lblImage) & ":"
    lblField(1).Caption = GetLbl(lblTitle) & ":"
    lblField(2).Caption = GetLbl(lblAlign) & ":"
    lblField(3).Caption = GetLbl(lblWidth) & ":"
    lblField(4).Caption = GetLbl(lblHeight) & ":"
    lblField(5).Caption = GetLbl(lblBorder) & ":"
    lblField(6).Caption = GetLbl(lblClass) & ":"
    'lblField(7).Caption = GetLbl(lblHSpace) & ":"
    'lblField(6).Caption = GetLbl(lblVSpace) & ":"
    cmdButton(0).Caption = GetLbl(lblOK)
    cmdButton(1).Caption = GetLbl(lblCancel)
    'Translate Align Combo
    cboAlign.List(0) = GetLbl(lblLeft)
    cboAlign.List(1) = GetLbl(lblRight)
    cboAlign.List(2) = GetLbl(lblMiddle)
    cboAlign.List(3) = GetLbl(lblAbsMiddle)
    cboAlign.List(4) = GetLbl(lblTop)
    cboAlign.List(5) = GetLbl(lblBottom)
    cboAlign.List(6) = GetLbl(lblCenter)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub

