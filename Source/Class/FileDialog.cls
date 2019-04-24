VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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

' Class       : CFileDialog
' Description : Class for displaying the File Open/Save Common Dialog.
'
Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private m_strDefaultExt As String
Private m_strDialogTitle As String
Private m_strFileName As String
Private m_strFileTitle As String
Private m_strInitialDir As String
Private m_strFilter As String
Private m_intFilterIndex As Integer
Private m_eFlags As EnumFilFlags
Private m_intMaxFileSize As Integer
Private m_lnghWndParent As Long

Private Const cintMaxFileLength As Integer = 260

Public Enum EnumFilFlags
        cdlReadOnly = &H1
        cdlOverWritePrompt = &H2
        cdlHideReadOnly = &H4
        cdlNoChangeDir = &H8
        cdlShowHelp = &H10
        cdlEnableHook = &H20
        cdlEnableTemplate = &H40
        cdlEnableTemplateHandle = &H80
        cdlNoValidate = &H100
        cdlAllowMultiselect = &H200
        cdlExtensionDifferent = &H400
        cdlPathMustExist = &H800
        cdlFileMustExist = &H1000
        cdlCreatePrompt = &H2000
        cdlShareAware = &H4000
        cdlNoReadOnlyReturn = &H8000
        cdlNoTestFileCreate = &H10000
        cdlNoNetworkButton = &H20000
        cdlExplorer = &H80000
        cdlLongnames = &H200000
End Enum

Public Property Get DefaultExt() As String
    ' Returns: The default extension

    DefaultExt = m_strDefaultExt

End Property

Public Property Let DefaultExt(ByVal strValue As String)
    ' strValue: Set the default extension used for a filename

    m_strDefaultExt = strValue

End Property

Public Property Get DialogTitle() As String
    ' Returns: The title displayed in the dialog
    

    DialogTitle = m_strDialogTitle

End Property

Public Property Let DialogTitle(ByVal strValue As String)
    ' strValue: Set the title displayed in the dialog
    

    m_strDialogTitle = strValue

End Property

Public Property Get FileName() As String
    ' Returns: The path and filename
    

    FileName = m_strFileName

End Property

Public Property Let FileName(ByVal strValue As String)
    ' strValue: Set the filename
    

    m_strFileName = strValue

End Property

Public Property Get FileTitle() As String
    ' Returns: The filename without the path
    

    FileTitle = m_strFileTitle

End Property

Public Property Let FileTitle(ByVal strValue As String)
    ' strValue: Set the file title
    

    m_strFileTitle = strValue

End Property

Public Property Get Filter() As String
    ' Returns: The filter string
    

    Filter = m_strFilter

End Property

Public Property Let Filter(ByVal strValue As String)
    ' strValue: Set the filter string
    

    m_strFilter = strValue

End Property

Public Property Get FilterIndex() As Integer
    ' Returns: The index of the filter to display
    

    FilterIndex = m_intFilterIndex

End Property

Public Property Let FilterIndex(ByVal intValue As Integer)
    ' Set the index of the filter to display
    

    m_intFilterIndex = intValue

End Property

Public Property Get Flags() As EnumFilFlags
    ' Returns: The flags
    

    Flags = m_eFlags

End Property

Public Property Let Flags(ByVal eValue As EnumFilFlags)
    ' eValue: Set the flags
    

    m_eFlags = eValue

End Property

Public Property Get hWndParent() As Long
    ' Returns: The parent hwnd
    

    hWndParent = m_lnghWndParent

End Property

Public Property Let hWndParent(ByVal lngValue As Long)
    ' lngValue: Set the parent hwnd
    

    m_lnghWndParent = lngValue

End Property

Public Property Get InitialDir() As String
        ' Returns: The current value of InitialDir
        

        InitialDir = m_strInitialDir

End Property

Public Property Let InitialDir(ByVal strValue As String)
        ' strValue: Set to the path where the dialog should open
        

        m_strInitialDir = strValue

End Property

Public Property Get MaxFileSize() As Integer
    ' Returns: The maximum length of the filename
    

    MaxFileSize = m_intMaxFileSize

End Property

Public Property Let MaxFileSize(ByVal intValue As Integer)
    ' intValue: Set the maximum length of the filename
    

    m_intMaxFileSize = intValue

End Property

Public Function ShowOpen() As Boolean
    ShowOpen = Show(True)
End Function

Public Function ShowSave() As Boolean
    ShowSave = Show(False)
End Function

Private Function Show(fOpen As Boolean) As Boolean
        ' Comments  : This procedure displays the file Open/Save common dialog
        ' Parameters: fOpen - Determines if the Open or Save dialog is displayed.
        ' Returns   : False if cancel selected, true otherwise.
        
        Dim of As OPENFILENAME
        Dim strChar As String * 1
        Dim intCounter As Integer
        Dim strTemp As String
        
        On Error GoTo PROC_ERR

        ' Initialize the OPENFILENAME type
        of.lpstrTitle = m_strDialogTitle & ""
        of.Flags = m_eFlags
        of.lpstrDefExt = m_strDefaultExt & ""
        of.lStructSize = LenB(of)
        of.lpstrFilter = m_strFilter & "||"
        of.nFilterIndex = m_intFilterIndex
        
        ' To make Windows-style filter, replace pipes with nulls
        For intCounter = 1 To Len(m_strFilter)
                strChar = Mid$(m_strFilter, intCounter, 1)
                If strChar = "|" Then
                        strTemp = strTemp & vbNullChar
                Else
                        strTemp = strTemp & strChar
                End If
        Next
        ' Put double null at end
        strTemp = strTemp & vbNullChar & vbNullChar
        of.lpstrFilter = strTemp
        
        ' Pad file and file title buffers to maximum path length
        strTemp = m_strFileName & String$(cintMaxFileLength - Len(m_strFileName), 0)
        of.lpstrFile = strTemp
        of.nMaxFile = cintMaxFileLength
        
        strTemp = m_strFileTitle & String$(cintMaxFileLength - Len(m_strFileTitle), 0)
        of.lpstrFileTitle = strTemp
        of.lpstrInitialDir = m_strInitialDir
        of.nMaxFileTitle = cintMaxFileLength
        of.hwndOwner = m_lnghWndParent
        
        ' If fOpen is true, show the Open file dialog, otherwise show the Save dialog
        If fOpen Then
                If GetOpenFileName(of) Then
                        Show = True
                        ' Assign property variables to appropriate values
                        m_strFileName = TrimNulls(of.lpstrFile)
                        m_strFileTitle = TrimNulls(of.lpstrFileTitle)
                        m_intFilterIndex = of.nFilterIndex
                Else
                        Show = False
                End If
        Else
                If GetSaveFileName(of) Then
                        Show = True
                        ' Assign property variables to appropriate values
                        m_strFileName = TrimNulls(of.lpstrFile)
                        m_strFileTitle = TrimNulls(of.lpstrFileTitle)
                        m_intFilterIndex = of.nFilterIndex
                Else
                        Show = False
                End If
        End If
        
PROC_EXIT:
        Exit Function

PROC_ERR:
        MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
                "Show"
        Resume PROC_EXIT

End Function

Private Function TrimNulls(ByVal strIn As String) As String
    ' Comments  : Returns the passed string terminated at the first double null
    ' Parameters: strIn - Value to parse
    ' Returns   : Parsed string
    
    '
    Dim intPos As Integer
    
    On Error GoTo PROC_ERR
        
    intPos = InStr(strIn, vbNullChar & vbNullChar)
    
    If intPos = 0 Then
        ' No nulls in the string, just return it as is
        TrimNulls = strIn
    Else
        If intPos = 1 Then
            ' If the null character is at the first position, the
            ' entire string is a null string, so return a zero-length string
            TrimNulls = ""
        Else
            ' Not at the first position, so return the contents up
            ' to the occurrence of the double null character
            TrimNulls = Left$(strIn, intPos - 1)
        End If
    End If
        
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "TrimNulls"
    Resume PROC_EXIT
        
End Function

