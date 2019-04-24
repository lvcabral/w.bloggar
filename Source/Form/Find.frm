VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboReplaceStr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   585
      Width           =   3435
   End
   Begin VB.ComboBox cboSearchStr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3435
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4785
      TabIndex        =   13
      Top             =   1635
      Width           =   1500
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   4785
      TabIndex        =   12
      Top             =   1125
      Width           =   1500
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   4785
      TabIndex        =   11
      Top             =   630
      Width           =   1500
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   4785
      TabIndex        =   10
      Top             =   120
      Width           =   1500
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   975
      Left            =   2070
      TabIndex        =   5
      Top             =   1035
      Width           =   2565
      Begin VB.CheckBox chkMatchWord 
         Appearance      =   0  'Flat
         Caption         =   "&Whole Word Only"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox chkMatchCase 
         Appearance      =   0  'Flat
         Caption         =   "&Match Case"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   2295
      End
   End
   Begin VB.Frame fraScope 
      Caption         =   "Scope"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1035
      Width           =   1830
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   720
         Left            =   45
         ScaleHeight     =   720
         ScaleWidth      =   1725
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   1725
         Begin VB.OptionButton optWholeText 
            Appearance      =   0  'Flat
            Caption         =   "Al&l Text"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   6
            Top             =   120
            Width           =   1665
         End
         Begin VB.OptionButton optSelected 
            Appearance      =   0  'Flat
            Caption         =   "&Selected"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   7
            Top             =   420
            Width           =   1665
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Replace:"
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   645
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find:"
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   180
      Width           =   360
   End
End
Attribute VB_Name = "frmFind"
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

' Search options
Private MatchCase As Boolean, MatchWord As Boolean

' Indicates whether the form is in Find-only state or
' Find/Replace
'
Public Replace As Boolean
Public Silent As Boolean

Private bolLocked As Boolean
Private strLastSearch As String
Private txtTarget As Control

Private Const WM_SYSCOMMAND = &H112
Private Const MOUSE_MOVE = &HF012

' These are positional constants for shuffling things
' around when user switches to Replace mode
'
Const REPLACE_HEIGHT = 2535
Const REPLACE_OPT_TOP = 1035
Const REPLACE_CLOSE_TOP = 1635

Const FIND_HEIGHT = 2205
Const FIND_OPT_TOP = 700
Const FIND_CLOSE_TOP = 1320

Private Wrapped As Boolean
Private LastLine As Long

Private Sub cboReplaceStr_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboSearchStr, KeyAscii
End Sub

Private Sub cboSearchStr_KeyPress(KeyAscii As Integer)
    ActiveComboBox cboSearchStr, KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Replace = False
    Silent = False
End Sub

Private Sub cboSearchStr_Change()
    LastLine = -1
End Sub

Private Sub chkMatchCase_Click()
    MatchCase = IIf(chkMatchCase.Value = 1, True, False)
End Sub

Private Sub chkMatchWord_Click()
    MatchWord = IIf(chkMatchWord.Value = 1, True, False)
End Sub

Private Sub cmdClose_Click()
    txtTarget.SetFocus
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
    DoFind
    UpdateLists
    strLastSearch = cboSearchStr.Text
End Sub

Private Sub cmdReplace_Click()
On Error GoTo ErrorHandler
Dim i%
    ' Switch to Replace mode if not in it, then exit sub
    If Not Replace Then
        Replace = True
        UpdateReplaceStatus
        Exit Sub
    End If
    
    If Len(cboSearchStr.Text) = 0 Then Exit Sub
    
    ' Replace next occurrence
    i = DoFind
    If i = -1 Then Exit Sub
    
    txtTarget.SelText = cboReplaceStr.Text
    txtTarget.SelStart = i
    txtTarget.SelLength = Len(cboReplaceStr.Text)
    
    UpdateLists
    strLastSearch = cboSearchStr.Text
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub cmdReplaceAll_Click()
On Error GoTo ErrorHandler
' Replace all occurrences
Dim Result%, c%, ST$, Compensate As Integer
Dim StartL As Long, EndL As Long
    If Len(txtTarget.Text) = 0 Then
        MsgBox GetMsg(msgNoTextSearch), vbExclamation
        Exit Sub
    End If

    If optSelected.Value = True Then
        StartL = txtTarget.SelStart
        EndL = StartL + txtTarget.SelLength
        
        If Len(cboReplaceStr.Text) > Len(cboSearchStr.Text) Then
            Compensate = 1
        ElseIf Len(cboSearchStr.Text) > Len(cboReplaceStr.Text) Then
            Compensate = 2
        End If
    Else
        StartL = 0
        EndL = Len(txtTarget.Text)
        Compensate = 0
    End If
    
    While Result <> -1
        Result = txtTarget.Find(cboSearchStr.Text, StartL, EndL, CreateFlags)
        
        If Result <> -1 Then
            txtTarget.SelText = ""
            txtTarget.SelText = cboReplaceStr.Text
            c = c + 1
            ' move past string for next search
            StartL = Result + Len(cboReplaceStr.Text)
            
            ' If the search string is longer or shorter than the replacement
            ' string, the ending character index will have to be changed each time
            ' a replacement is made.
            '
            If Compensate = 1 Then
                EndL = EndL + (Len(cboReplaceStr.Text) - Len(cboSearchStr.Text))
            ElseIf Compensate = 2 Then
                EndL = EndL - (Len(cboSearchStr.Text) - Len(cboReplaceStr.Text))
            End If
        End If
    Wend
    
    If c = 0 Then
        MsgBox GetMsg(msgNoMatches), vbInformation
    Else
        MsgBox c & " " & GetMsg(msgReplMade), vbInformation
    End If

    UpdateLists
    strLastSearch = cboSearchStr.Text
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next
    LocalizeForm
    'Get the pointer to the target Text Control
    Select Case frmPost.ActiveControl.Name
    Case "txtPostTit", "txtPost"
        Set txtTarget = frmPost.txtPost
    Case "txtMore", "txtKeywords"
        Set txtTarget = frmPost.txtMore
    Case "txtExcerpt"
        Set txtTarget = frmPost.txtExcerpt
    Case Else
        If frmPost.CurrentTab = TAB_EDITOR Then
            Set txtTarget = frmPost.txtPost
        Else
            Set txtTarget = frmPost.txtMore
        End If
    End Select
    cboSearchStr.Text = strLastSearch
    If Silent And cboSearchStr.Text <> "" Then Exit Sub
    UpdateReplaceStatus
    If txtTarget.SelLength > 0 Then
        If Not txtTarget.SelText Like "*" & vbCrLf & "*" Then
            cboSearchStr.Text = txtTarget.SelText
            optWholeText.Value = True
        Else
            optSelected.Value = True
        End If
    Else
        optWholeText.Value = True
        optSelected.Enabled = False
    End If
End Sub

' Performs a basic find based on selected options and
' returns the result (matching position or -1)
'
Public Function DoFind() As Integer
On Error GoTo ErrorHandler
Dim Result As Long
Dim l As Long, r$
Dim StartL As Long, EndL As Long
    If Len(txtTarget.Text) = 0 Then
        MsgBox GetMsg(msgNoTextSearch), vbCritical
        Exit Function
    End If
    
'    Set rtbText = txtTarget
    StartL = txtTarget.SelStart
        
    If Wrapped Then
        ' Start at the top and go down
        bolLocked = True
        Result = txtTarget.Find(cboSearchStr.Text, 0, , CreateFlags)
        If Result = -1 Then
            MsgBox GetMsg(msgNoMatches), vbInformation
        Else
            l = txtTarget.GetLineFromChar(Result)
            If LastLine = l Then
                r$ = "Only m"
            Else
                r$ = "M"
            End If
            Debug.Print r$ & "atch found on line " & l
            LastLine = l
        End If
        Wrapped = False
        DoFind = Result
        bolLocked = False
    Else
        If optSelected.Value = True Then
            EndL = StartL + txtTarget.SelLength
        Else
            StartL = StartL + 1
            EndL = Len(txtTarget.Text)
        End If
        
        ' Go down
        bolLocked = True
        Result = txtTarget.Find(cboSearchStr.Text, StartL, EndL, CreateFlags)
        If Result = -1 Then
            ' If only searching selected text then
            ' call it quits
            If optSelected.Value = True Then
                MsgBox GetMsg(msgNoMatchesSel), vbInformation
                DoFind = Result
            ' Otherwise wrap around to beginning
            Else
                Wrapped = True
                DoFind = DoFind()    ' Recursively call to search again from top
            End If
        Else
            DoFind = Result
            l = txtTarget.GetLineFromChar(Result)
            If LastLine = l Then
                r$ = "Another m"
            Else
                r$ = "M"
            End If
            Debug.Print r$ & "atch found on line " & txtTarget.GetLineFromChar(Result)
            LastLine = l
            bolLocked = False
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Function

Private Function CreateFlags() As Integer
Dim FindFlags%
    FindFlags = 0
    If MatchCase And MatchWord Then
        FindFlags = rtfMatchCase Or rtfWholeWord
    ElseIf MatchWord Then
        FindFlags = rtfWholeWord
    ElseIf MatchCase Then
        FindFlags = rtfMatchCase
    End If
    CreateFlags = FindFlags
End Function

' Adds new items to Find and Replace lists for use in
' AutoComplete.  Called whenever a new search or replace
' is performed.
Private Sub UpdateLists()
On Error GoTo ErrorHandler
Dim i As Integer
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    If Not InList(cboSearchStr, cboSearchStr.Text) And _
        Len(cboSearchStr.Text) > 0 Then
        objXMLReg.SaveSetting App.Title, "Search", "Find00", cboSearchStr.Text
        If cboSearchStr.ListCount > 0 Then
            For i = 1 To Min(cboSearchStr.ListCount, 10)
                objXMLReg.SaveSetting App.Title, "Search", "Find" & Format(i, "00"), cboSearchStr.List(i - 1)
            Next
        End If
        cboSearchStr.AddItem cboSearchStr.Text, 0
        cboSearchStr.ListIndex = 0
    End If
    
    If Not Replace Then Exit Sub
    
    If Not InList(cboReplaceStr, cboReplaceStr.Text) And _
        Len(cboReplaceStr.Text) > 0 Then
        objXMLReg.SaveSetting App.Title, "Search", "Replace00", cboReplaceStr.Text
        If cboReplaceStr.ListCount > 0 Then
            For i = 1 To Min(cboReplaceStr.ListCount, 10)
                objXMLReg.SaveSetting App.Title, "Search", "Replace" & Format(i, "00"), cboReplaceStr.List(i - 1)
            Next
        End If
        cboReplaceStr.AddItem cboReplaceStr.Text, 0
        cboReplaceStr.ListIndex = 0
    End If
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

' Shuffle things around depending on whether this is
' a Find or a Find/Replace dialog
'
Private Sub UpdateReplaceStatus()
On Error GoTo ErrorHandler
Dim i As Integer, strAux As String
    'Load Find Combo
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    For i = 0 To 10
        strAux = objXMLReg.GetSetting(App.Title, "Search", "Find" & Format(i, "00"), "**")
        If strAux = "**" Then
            Exit For
        Else
            cboSearchStr.AddItem strAux
        End If
    Next
    If Not Replace Then
        cmdReplace.Caption = "&" & GetLbl(lblReplace) & "..."
        cmdReplaceAll.Visible = False
        cboReplaceStr.Visible = False
        Label2.Visible = False
        
        cmdClose.Top = FIND_CLOSE_TOP
        fraOptions.Top = FIND_OPT_TOP
        fraScope.Top = FIND_OPT_TOP
        Me.Height = FIND_HEIGHT
    Else
        'Load Replace Combo
        For i = 0 To 10
            strAux = objXMLReg.GetSetting(App.Title, "Search", "Replace" & Format(i, "00"), "**")
            If strAux = "**" Then
                Exit For
            Else
                cboReplaceStr.AddItem strAux
            End If
        Next
        cmdReplace.Caption = "&" & GetLbl(lblReplace)
        frmFind.Caption = GetLbl(lblReplace)
        cmdReplaceAll.Visible = True
        cboReplaceStr.Visible = True
        Label2.Visible = True
        
        cmdClose.Top = REPLACE_CLOSE_TOP
        fraOptions.Top = REPLACE_OPT_TOP
        Me.Height = REPLACE_HEIGHT
        fraScope.Top = REPLACE_OPT_TOP
    End If
    Set objXMLReg = Nothing
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, Me.Name
End Sub

' Checks for the existence of a string in a list
Private Function InList(l As Variant, s As String) As Boolean
Dim i%, c As New Collection, t$
    
    t$ = TypeName(l)
    
    ' If we have a combobox or listbox here, copy it over
    ' to a collection. This way we can handle all kinds
    ' of requests.
    '
    If t$ = "ComboBox" Or t$ = "ListBox" Then
        If l.ListCount = 0 Then
            InList = False
            Exit Function
        End If
        
        For i = 0 To l.ListCount - 1
            c.Add l.List(i)
        Next i
    ElseIf t$ = "Collection" Then
        Set c = l
    Else
        InList = False
        Exit Function
    End If
    
    If c.Count = 0 Then
        InList = False
        Exit Function
    Else
        ' Do a non-case-sensitive search
        For i = 1 To c.Count
            If (s = c(i)) Or ((Len(s) = Len(c(i))) And _
                InStr(1, s, c(i), 1)) Then
                    InList = True
                    Exit Function
            End If
        Next i
    End If
    
    InList = False
End Function

Private Sub LocalizeForm()
On Error GoTo ErrorHandler:
    Caption = GetLbl(lblFind)
    Label1.Caption = GetLbl(lblFind) & ":"
    Label2.Caption = GetLbl(lblReplace) & ":"
    fraScope.Caption = GetLbl(lblScope)
    optWholeText.Caption = GetLbl(lblAllText)
    optSelected.Caption = GetLbl(lblSelected)
    fraOptions.Caption = GetLbl(lblOptions)
    chkMatchCase.Caption = GetLbl(lblMatchCase)
    chkMatchWord.Caption = GetLbl(lblWholeWord)
    cmdFindNext.Caption = GetLbl(lblFindNext)
    cmdReplace.Caption = "&" & GetLbl(lblReplace)
    cmdReplaceAll.Caption = GetLbl(lblReplaceAll)
    cmdClose.Caption = GetLbl(lblClose)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
