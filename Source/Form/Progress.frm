VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.ElasticOne pnlProgress 
      Height          =   225
      Left            =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   405
      Width           =   4275
      _cx             =   7541
      _cy             =   397
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
      Appearance      =   4
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
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   2205
      TabIndex        =   1
      Top             =   60
      Width           =   195
   End
End
Attribute VB_Name = "frmProgress"
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
Private lngWalk As Long

Private Sub Form_Load()
    pnlProgress.FloodDirection = fdRight
    pnlProgress.FloodStyle = fsBlocks
End Sub

Public Sub RunJob(ByVal strJob As String, Optional ByVal strParam As String)
    lngWalk = 0
    Select Case strJob
    Case "ConvertSettings"
        Screen.MousePointer = vbHourglass
        lblMessage.Caption = GetMsg(msgConvSettingsXML)
        ConvertSettings
        Screen.MousePointer = vbDefault
        If MsgBox(GetMsg(msgKeepSettings), vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Uninstall
        End If
    Case "LoadWords"
        lblMessage.Caption = GetMsg(msgLoadingDict) & ": " & GetLocaleName(gSettings.SpellLCID) & "..."
        LoadWords strParam
    End Select
    Unload Me
End Sub

Private Sub WalkOnProgress(ByVal lngTotal As Long, Optional ByVal lngStep As Long)
On Error Resume Next
    lngWalk = lngWalk + 1
    If lngStep > 0 Then
        pnlProgress.FloodPercent = lngStep * 100 / lngTotal
    Else
        pnlProgress.FloodPercent = lngWalk * 100 / lngTotal
    End If
    DoEvents
End Sub

Private Sub ConvertSettings()
Dim aAccounts() As String, aBlogs() As String
Dim strRoot As String, f As Integer, intSteps As Integer
    Set objXMLReg = New XMLRegistry
    If Not objXMLReg.OpenXMLFile(gAppDataPath & XML_SETTINGS, True) Then
        MsgBox GetMsg(msgErrOpenSettings), vbCritical
        End
    End If
    Show
    DoEvents
    strRoot = "Software\VB and VBA Program Settings\" & REGISTRY_KEY & "\"
    'Calculate Steps
    intSteps = 10
    aAccounts = EnumKeys(HKEY_CURRENT_USER, strRoot & "Accounts")
    aBlogs = EnumKeys(HKEY_CURRENT_USER, strRoot & "Blogs")
    If Not IsArrayEmpty(aAccounts) Then
        intSteps = intSteps + UBound(aAccounts)
    End If
    If Not IsArrayEmpty(aBlogs) Then
        intSteps = intSteps + UBound(aBlogs)
    End If
    'Convert Accounts
    WalkOnProgress intSteps
    If Not IsArrayEmpty(aAccounts) Then
        For f = 0 To UBound(aAccounts)
            RegToXML "Accounts\" & aAccounts(f), objXMLReg
            WalkOnProgress intSteps
        Next
    End If
    'Convert Blogs
    If Not IsArrayEmpty(aBlogs) Then
        For f = 0 To UBound(aBlogs)
            RegToXML "Blogs\" & aBlogs(f), objXMLReg
            WalkOnProgress intSteps
        Next
    End If
    'Convert Other Settings
    RegToXML "Settings", objXMLReg
    WalkOnProgress intSteps
    RegToXML "Colors", objXMLReg
    WalkOnProgress intSteps
    RegToXML "Forms", objXMLReg
    WalkOnProgress intSteps
    RegToXML "Images", objXMLReg
    WalkOnProgress intSteps
    RegToXML "Links", objXMLReg
    WalkOnProgress intSteps
    RegToXML "MRU", objXMLReg
    WalkOnProgress intSteps
    RegToXML "Search", objXMLReg
    WalkOnProgress intSteps
    RegToXML "Table", objXMLReg
    WalkOnProgress intSteps
    Set objXMLReg = Nothing
End Sub

Private Sub RegToXML(ByVal strSection As String, objXMLRegistry As XMLRegistry)
Dim aItems() As String, i As Integer, strItem As String
    aItems = EnumValues(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\" & REGISTRY_KEY & "\" & strSection & "\")
    If IsArrayEmpty(aItems) Then Exit Sub
    For i = 0 To UBound(aItems)
        strItem = GetSetting(REGISTRY_KEY, strSection, aItems(i))
        If LCase(Left(aItems(i), 9)) = "customtag" Then
            objXMLRegistry.SaveSetting App.Title, "CustomTags", aItems(i), strItem
        ElseIf LCase(Left(aItems(i), 5)) = "class" And LCase(strSection) = "links" Then
            objXMLRegistry.SaveSetting App.Title, "Classes", aItems(i), strItem
        ElseIf LCase(aItems(i)) = "skinfolder" Then
            objXMLRegistry.SaveSetting App.Title, strSection, "Skin", GetNamePart(strItem, True)
        ElseIf strItem <> "" Then
            objXMLRegistry.SaveSetting App.Title, Replace(strSection, "\", "/" & Left(strSection, 1)), aItems(i), strItem
        End If
    Next
End Sub

Private Sub LoadWords(strDictionaryFile As String)
    Dim iFreeFile As Integer
    Dim strLine As String
    Dim strSoundExCode As String
'    Dim fT As Double

    Dim lRead As Long, lLen As Long, lThisRead As Long, lLastRead As Long
    Dim sBuf As String, iNextPos As Long, iLastPos As Long
    Dim sWord As String, sRemain As String
    
'    fT = Timer
    bFileNotFound = False
    iFreeFile = FreeFile
    Set WordTree = New clsWordTree
    On Error GoTo DiskErr
    Show
    DoEvents
    ' Open file for binary access:
    Open strDictionaryFile For Binary Access Read Lock Write As #iFreeFile
    lLen = LOF(iFreeFile)
    
    ' Loop through the file, loading it up in chunks of 64k:
    Do While lRead < lLen
        lThisRead = 65536
        If lThisRead + lRead > lLen Then
            lThisRead = lLen - lRead
        End If
        If Not lThisRead = lLastRead Then
            sBuf = Space$(lThisRead)
        End If
        
        Get #iFreeFile, , sBuf
        lRead = lRead + lThisRead
        
        ' Extract elements from string:
        iLastPos = 1
        Do
            iNextPos = InStr(iLastPos, sBuf, vbCrLf)
            If iNextPos = 0 Then
                If iLastPos < lThisRead Then
                   sRemain = Mid$(sBuf, iLastPos)
                End If
            Else
                sWord = Mid$(sBuf, iLastPos, iNextPos - iLastPos)
                If Len(sRemain) > 0 Then
                   sWord = sRemain & sWord
                   sRemain = ""
                End If
                strSoundExCode = Soundex(sWord)
                WordTree.Add strSoundExCode, sWord
                iLastPos = iNextPos + 2
            End If
        Loop While Not (iNextPos = 0)
        WalkOnProgress lLen, lRead
    Loop
    
    ' any remainder needs to be added as a word:
    WordTree.Add Soundex(sWord), sWord
    Close iFreeFile
    ' Time taken
    'Debug.Print Timer - fT, WordTree.Count
        
    Exit Sub
DiskErr:
    If Err = 53 Then
        bFileNotFound = True
    End If
    Err.Raise Err.Number, "vbwSpellCheck", Err.Description
End Sub

