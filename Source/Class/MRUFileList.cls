VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMRUFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private m_lMaxCount As Long
Private m_lFileCount As Long
Private m_sFiles() As String

Public Property Get MaxFileCount() As Long
    MaxFileCount = m_lMaxCount
End Property
Public Property Let MaxFileCount(ByVal lMax As Long)
    If (m_lMaxCount <> lMax) Then
        ReDim Preserve m_sFiles(1 To lMax) As String
        m_lMaxCount = lMax
    End If
End Property
Public Property Get FileCount() As Long
    FileCount = m_lFileCount
End Property
Public Property Get file(ByVal iIndex As Long) As String
    file = m_sFiles(iIndex)
End Property
Public Property Get FileExists(ByVal iIndex As Long) As Boolean
    FileExists = (m_sFiles(iIndex) <> "")
End Property
Public Property Get MenuCaption(ByVal iIndex As Long) As String
Dim sOut As String

    sOut = CompressFileName(m_sFiles(iIndex))
    MenuCaption = "&" & iIndex & "  " & sOut
    
End Property

Public Function CompressFileName(ByVal sFileName As String) As String
Dim lLen As Long
Dim iPos As Long
Dim sOut As String
Dim iLastPos As Long
    ' Easier way using API probably, but here goes:
    lLen = Len(sFileName)
    sOut = sFileName
    If (lLen > 32) Then
        For iPos = lLen To 1 Step -1
            If (Mid$(sFileName, iPos, 1) = "\") Then
                If (lLen - iPos) > 32 Then
                    If (iLastPos = 0) Then
                        sOut = "..." & Mid$(sFileName, iPos)
                    Else
                        sOut = "..." & Mid$(sFileName, iLastPos)
                    End If
                Else
                    iLastPos = iPos
                End If
            End If
        Next iPos
    End If
    CompressFileName = sOut
End Function
Public Sub AddFile(ByVal sFileName As String)
Dim iFile As Long
Dim iExistIndex As Long
Dim iSwap As Long
    'If Dir(sFileName) = "" Then Exit Sub
    If Not basBloggar.FileExists(sFileName) Then Exit Sub
    ' does the file already exist?
    For iFile = 1 To m_lFileCount
        If (UCase$(m_sFiles(iFile)) = UCase$(sFileName)) Then
            iExistIndex = iFile
            Exit For
        End If
    Next iFile
    
    ' If it exists swap all the entries above it down
    ' by one, otherwise swap all entries down by one:
    If (iExistIndex = 0) Then
        If (m_lFileCount < MaxFileCount) Then
            m_lFileCount = m_lFileCount + 1
        End If
        iSwap = m_lFileCount
    Else
        iSwap = iExistIndex
    End If
    
    For iFile = iSwap - 1 To 1 Step -1
        m_sFiles(iFile + 1) = m_sFiles(iFile)
    Next iFile
    
    ' Add this entry:
    m_sFiles(1) = sFileName
    
End Sub

Public Function Save() As Boolean
On Error GoTo SaveError

Dim iFile As Long

    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    
    Call objXMLReg.SaveSetting(App.Title, "MRU", "MaxCount", Format(MaxFileCount))
    Call objXMLReg.SaveSetting(App.Title, "MRU", "Count", Format(m_lFileCount))

    For iFile = 1 To m_lFileCount
        Call objXMLReg.SaveSetting(App.Title, "MRU", "File" & iFile, m_sFiles(iFile))
    Next iFile
    Set objXMLReg = Nothing
    Save = True
    Exit Function
SaveError:
    Exit Function
End Function

Public Function Load() As Boolean
On Error GoTo LoadError
    
Dim iFile As Long
Dim iFileNo As Long
Dim sFile As String
    
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    MaxFileCount = Val(objXMLReg.GetSetting(App.Title, "MRU", "MaxCount", "-1"))
    If MaxFileCount > -1 Then
        If (MaxFileCount > 0) Then
            m_lFileCount = Val(objXMLReg.GetSetting(App.Title, "MRU", "Count", "0"))
            For iFile = 1 To m_lFileCount
                sFile = objXMLReg.GetSetting(App.Title, "MRU", "File" & iFile, "")
                If Dir$(sFile) <> "" And sFile <> "" Then
                    iFileNo = iFileNo + 1
                    m_sFiles(iFile) = sFile
                Else
                    'Commented for 4.00 - DeleteSetting REGISTRY_KEY, "MRU", "File" & iFile
                    objXMLReg.SaveSetting App.Title, "MRU", "File" & iFile, ""
                End If
            Next iFile
            m_lFileCount = iFileNo
        End If
    End If
    Set objXMLReg = Nothing
    Load = True
    Exit Function
LoadError:
    Exit Function
End Function

