VERSION 5.00
Begin VB.Form frmCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spell Check"
   ClientHeight    =   3150
   ClientLeft      =   3345
   ClientTop       =   5070
   ClientWidth     =   6300
   Icon            =   "Check.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdIgnoreAll 
      Caption         =   "Ignore A&ll"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4665
      TabIndex        =   8
      Top             =   1215
      Width           =   1500
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4665
      TabIndex        =   9
      Top             =   1665
      Width           =   1500
   End
   Begin VB.TextBox txtChangeTo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   945
      Width           =   4320
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4665
      TabIndex        =   10
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4665
      TabIndex        =   7
      Top             =   765
      Width           =   1500
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "C&hange"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4665
      TabIndex        =   6
      Top             =   315
      Width           =   1500
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   330
      Width           =   4320
   End
   Begin VB.ListBox lstSuggestions 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      TabIndex        =   5
      Top             =   1590
      Width           =   4320
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Change &To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   720
      Width           =   840
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Not In Dictionary:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "&Suggested words:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   4
      Top             =   1350
      Width           =   1305
   End
End
Attribute VB_Name = "frmCheck"
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F

Private txtText As Object ' The textbox to spell check

Private bNoClickEvent As Boolean
Public bDictChanged As Boolean
Private nStart As Long
Private nEnd As Long

Public Event BeforeChange(bCancel As Boolean)
Public Event BeforeAdd(bCancel As Boolean)
Public Event BeforeSave()
Public Event CompleteSpellCheck()
Private bStartSentance As Boolean
' Returns the position of next word divider
Function getNextWordDivider(ByVal strText As String, ByVal lStartPos As Long) As Long
    Dim lLength As Long
    Dim sChar As String
    Dim bOpenedTag As Boolean
    Static bWasPeriod As Boolean
    Static lLastVal As Long
    lLength = Len(strText)
    If lStartPos = 0 Then lStartPos = 1
    If Mid(strText, lStartPos, 1) = "<" Then
        '// Start a HTML Tag
        bOpenedTag = True
    End If
    For lStartPos = lStartPos To lLength
        sChar = Mid$(strText, lStartPos, 1)
        Select Case sChar
        '// any letter or number
        '--changed by Marcelo Cabral (09/06/2002)
        '--Added special characters
        Case "A" To "Z", "a" To "z", "À" To "Ö", "Ø" To "ö", _
             "ø" To "ÿ", 0 To 9, "-", "&", "/"
            DoEvents
        '// Close a HTML Tag
        Case "<"
            If Not bOpenedTag Then
                getNextWordDivider = lStartPos - 1
                Exit Function
            End If
        '// Close a HTML Tag
        Case ">"
            getNextWordDivider = lStartPos
            Exit Function
        '// other punctuation (ie new word)
        Case Else
          If Not bOpenedTag Then
            '// check that it is a word in quotes (ie 'hello'), rather than
            '// an Apostrophe (ie wasn't)
            On Error Resume Next
            If Not Mid$(strText, lStartPos - 1, 3) Like "[A-z]'[A-z]" Then
                If Mid$(strText, lStartPos - 1, 4) Like "[A-z].[A-z]." Then
                    lStartPos = lStartPos + 2
                Else
                    If Mid$(strText, lLastVal - 1, 2) <> ". " Then
                        bStartSentance = False
                    End If
                    If bWasPeriod Then
                        bStartSentance = True
                        bWasPeriod = False
                    End If
                    If sChar = "." Then
                        bWasPeriod = True
                    End If
                    lLastVal = lStartPos
                    getNextWordDivider = lStartPos - 1
                    Exit Function
                End If
            End If
          End If
        End Select
    Next
    getNextWordDivider = Len(strText) + 1
End Function

' Loops until it finds a misspelled word
' Returns true if there was a spelling mistake
Private Function SpellCheckFrom(StartPos As Long, Optional EndPos As Long = -1) As Boolean
    Dim DocumentLength As Long
    Dim WordLength As Integer
    Dim sWord As String
    Dim SpellingNotOK As Boolean
    Dim sText As String
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    DoEvents
    If EndPos = -1 Then
        EndPos = Len(txtText.Text)
    End If
    sText = txtText.Text
    If StartPos <> 0 Then
        '// if we aren't at the start of the document
        '// find the next word
        txtText.SelStart = getNextWordDivider(sText, StartPos)
        If Mid(txtText.Text, txtText.SelStart, 1) = ">" Then
            StartPos = txtText.SelStart
        Else
            StartPos = txtText.SelStart + 1
        End If
    Else
        '// if there are word dividers before the first word
        '// then go there
        If getNextWordDivider(sText, StartPos) = 1 Then
            StartPos = 1
        End If
    End If
    Do While StartPos <= EndPos
        WordLength = getNextWordDivider(sText, StartPos + 1) - StartPos
        If WordLength = -1 Then Exit Do
        sWord = (Mid$(sText, StartPos + 1, WordLength))
        ''''''''
        ' Highlight Word
        If WordLength <> 0 And sWord <> "" And Val(sWord) = 0 Then
            
            If IsCorrectWord(sWord) = False Then
                txtText.SelStart = StartPos
                txtText.SelLength = WordLength
                ListAlternates (sWord)
                cmdAdd.Visible = True
                cmdIgnoreAll.Visible = True
                Show , txtText.Parent
                SpellCheckFrom = True
                Screen.MousePointer = vbDefault
                Exit Function
            Else
                'If Mid$(sText, IIf(StartPos < 2, 2, StartPos) - 1, 2) Like ".*" Or StartPos = 0 Then
                If bStartSentance Or StartPos = 0 Then
                    '// start of new sentance
                    '// check if letter is upper case
                    If Mid$(sText, StartPos + 1, 1) <> UCase(Mid$(sText, StartPos + 1, 1)) Then
                        '// not correct case
                        txtText.SelStart = StartPos
                        txtText.SelLength = WordLength
                        txtWord = sWord
                        lstSuggestions.Clear
                        lstSuggestions.AddItem StrConv(sWord, vbProperCase)
                        lstSuggestions.ListIndex = 0
                        cmdAdd.Visible = False
                        cmdIgnoreAll.Visible = False
                        Show , txtText.Parent
                        SpellCheckFrom = True
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If
                End If
            End If
        End If
        '--changed by Marcelo Cabral (17/07/2002)
        '--Very if the last word is a tag
        If Right(sWord, 1) = ">" Then
            StartPos = StartPos + WordLength
        Else
            StartPos = StartPos + WordLength + 1
        End If
    Loop
    Screen.MousePointer = vbDefault
    Exit Function
ErrHandler:
    MsgBox "Error: " & Err & " - " & Error, vbExclamation
End Function
Function IsCorrectWord(sWord As String) As Boolean
    Dim Col As Collection
    Dim strSoundExCode As String
    Dim WordLength As Long
On Error GoTo ErrHandler
    sWord = Trim$(sWord)
    WordLength = Len(sWord)
    If WordLength = 0 Then Exit Function
    If IsNumeric(sWord) Or Right(sWord, 1) = ">" Then
        IsCorrectWord = True
        Exit Function
    End If
    strSoundExCode = Soundex(sWord) ' Generate the code for the word
    If Len(strSoundExCode) > 0 Then
        Set Col = WordTree.GetCodeNode(strSoundExCode) ' Get all words with this code
        If Col Is Nothing Then
            ' Code not found
            IsCorrectWord = False
            Exit Function
        Else
            Dim i As Integer
            Dim LettersToCompare As Integer
            For i = 1 To Col.Count
                '// if it is upper case in the dictionary, then
                '// bring up error. If it is lowercase, don't
                '// worry about the case
                If Left$(Col.Item(i), 1) = UCase$(Left$(Col.Item(i), 1)) Then
                    If Col.Item(i) = sWord Then
                        IsCorrectWord = True
                        Exit Function
                    End If
                ElseIf LCase$(Col.Item(i)) = LCase$(sWord) Then
                    ' The word is in the dictionary.
                    ' No need to bother user.
                    IsCorrectWord = True
                    Exit Function
                End If
            Next
        End If
    End If
    Exit Function
ErrHandler:
    MsgBox "Error: " & Err & " - " & Error, vbExclamation
End Function
'// Returns true if mis-spelled
Function ListAlternates(sWord As String) As Boolean
    Dim Col As Collection
    Dim strSoundExCode As String
    Dim WordLength As Long
    Dim LettersThatMatch As Long
    Dim MostLikelyWord As Long
    Dim J As Long
On Error GoTo ErrHandler
    txtWord = sWord
    sWord = LCase$(Trim$(sWord))
    WordLength = Len(sWord)
    If WordLength = 0 Then Exit Function
    
    lstSuggestions.Clear
    If bFileNotFound Then
        lstSuggestions.AddItem GetLbl(lblDicNotFound)
        lstSuggestions.Enabled = False
        ListAlternates = True
        Exit Function
    End If
    strSoundExCode = Soundex(sWord) ' Generate the code for the word
    If Len(strSoundExCode) > 0 Then
        Set Col = WordTree.GetCodeNode(strSoundExCode) ' Get all words with this code
        
        If Col Is Nothing Then
            ' Code not found
            lstSuggestions.AddItem GetLbl(lblNoSuggestions)
            txtChangeTo = txtWord
            lstSuggestions.Enabled = False
            ListAlternates = True
        Else
            lstSuggestions.Enabled = True
            Dim i As Integer
            Dim LettersToCompare As Integer
            For i = 1 To Col.Count
                ListAlternates = True
                lstSuggestions.AddItem Col.Item(i) ' Add suggestion
                
                ' Select the word from the suggestions box
                ' that has the most characters in common with
                ' with the word being checked.
                LettersToCompare = Len(Col.Item(i))
                If sWord = LCase$(Col.Item(i)) Then
                    lstSuggestions.Clear
                    lstSuggestions.AddItem Col.Item(i)
                    MostLikelyWord = 1
                    
                    Exit For
                Else
                    If LettersToCompare > WordLength Then LettersToCompare = WordLength
                    
                    For J = 1 To LettersToCompare
                        If Mid$(Col.Item(i), J, 1) = Mid$(sWord, J, 1) Then
                            If J > LettersThatMatch Then
                                LettersThatMatch = J
                                MostLikelyWord = i
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    ' End of finding closest match
                End If
                
            Next
            lstSuggestions.ListIndex = MostLikelyWord - 1 ' Select the most likely word
        End If
    End If
    Exit Function
ErrHandler:
    MsgBox "Error: " & Err & " - " & Error, vbExclamation
End Function

Private Sub cmdAdd_Click()
    Dim bCancel As Boolean
    RaiseEvent BeforeAdd(bCancel)
    If bCancel Then Exit Sub
    AddItem txtWord, False
    cmdIgnore_Click
End Sub

Public Function AddItem(strWord As String, bShowMessage As Boolean) As Boolean
    Dim strSoundExCode As String
    Dim WordCollection As Collection
    Dim i As Long
    
    AddItem = True
    strSoundExCode = Soundex(strWord)
    On Error Resume Next
    Set WordCollection = WordTree.Root.Item(strSoundExCode)
    If Err = 0 Then
        For i = 1 To WordCollection.Count
            If WordCollection.Item(i) = strWord Then
                MsgBox GetMsg(msgItemExists), vbExclamation
                Exit Function
            End If
        Next
    End If
    If Len(strSoundExCode) > 0 Then
        WordTree.Add strSoundExCode, strWord
        ReDim Preserve WordsToAdd(0 To UBound(WordsToAdd) + 1)
        WordsToAdd(UBound(WordsToAdd)) = strWord
        bDictChanged = True
    End If
End Function
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdIgnoreAll_Click()
    '// add for this session only
    Dim strSoundExCode As String
    strSoundExCode = Soundex(txtWord)
    If Len(strSoundExCode) > 0 Then
        WordTree.Add strSoundExCode, txtWord
    End If
    cmdIgnore_Click
End Sub

Private Sub cmdChange_Click()
    Dim bCancel As Boolean
    RaiseEvent BeforeChange(bCancel)
    If bCancel Then Exit Sub
    txtText.SelText = txtChangeTo.Text
    '// then, check the word we have changed it to.
    txtText.SelStart = txtText.SelStart - Len(txtChangeTo.Text)
    cmdIgnore_Click
End Sub

Public Sub SpellCheck(txtBox As Object)
    Set txtText = txtBox
    nStart = txtText.SelStart - 1
    nEnd = -1
    cmdIgnore_Click
End Sub
Private Sub cmdIgnore_Click()
    If Not SpellCheckFrom(txtText.SelStart + txtText.SelLength, nEnd) Then
        If nStart <> -1 Then
            txtText.SelStart = 0
            nEnd = nStart
            nStart = -1
            If SpellCheckFrom(txtText.SelStart, nEnd) = False Then
                Complete
            End If
        Else
            Complete
        End If
    Else
        txtWord.Text = txtText.SelText
    End If
End Sub
Private Sub Complete()
    Hide
    RaiseEvent CompleteSpellCheck
    MsgBox GetMsg(msgSpellComplete), vbOKOnly + vbInformation, Caption
    cmdClose_Click
End Sub

Private Sub Form_Load()
    RestoreWindowPos
    LocalizeForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    txtText.SelLength = 0
    SaveWindowPos
    Hide
    RaiseEvent BeforeSave
End Sub
Private Sub RestoreWindowPos()
On Error Resume Next
    LoadFormSettings Me
End Sub
Private Sub SaveWindowPos()
On Error Resume Next
    SaveFormSettings Me
End Sub
Private Sub lstSuggestions_Click()
    If Not bNoClickEvent Then
        txtChangeTo.Text = lstSuggestions.Text
    End If
End Sub
Private Sub txtChangeTo_Change()
    bNoClickEvent = True
    lstSuggestions.ListIndex = SendMessage(lstSuggestions.hwnd, LB_FINDSTRING, -1, ByVal txtChangeTo.Text)
    bNoClickEvent = False
End Sub
Private Sub txtWord_DblClick()
    txtChangeTo = txtWord
End Sub

Private Sub LocalizeForm()
On Error GoTo ErrorHandler
    Me.Caption = frmPost.acbMain.Tools("miTSpelling").ToolTipText & ": " & GetLocaleName(gSettings.SpellLCID)
    lblField(0).Caption = GetLbl(lblNotInDic) & ":"
    lblField(1).Caption = GetLbl(lblChangeTo) & ":"
    lblField(2).Caption = GetLbl(lblSuggestedWords) & ":"
    cmdChange.Caption = GetLbl(lblChange)
    cmdIgnore.Caption = GetLbl(lblIgnore)
    cmdIgnoreAll.Caption = GetLbl(lblIgnoreAll)
    cmdAdd.Caption = GetLbl(lblAdd)
    cmdClose.Caption = GetLbl(lblCancel)
    Exit Sub
ErrorHandler:
    Debug.Print "Error on Localize: " & Me.Name & " - " & "(" & Err & ") " & Err.Description
    Resume Next
End Sub
