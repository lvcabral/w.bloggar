VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"No"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
' Class       : CColorDialog
' Description : Class for displaying the Choose Color Common Dialog

'
Private Declare Function ChooseColor_API Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Enum EnumColFlags
    colEnableHook = &H10
    colEnableTemplate = &H20
    colEnableTemplateHandle = &H40
    colFullOpen = &H2
    colPreventFullOpen = &H4
    colRgbinit = &H1
    colShowHelp = &H8
End Enum

Private m_lngColor As Long
Private m_eFlags As EnumColFlags
Private m_lnghWndParent As Long
Private m_alngColorRef(15) As Long

Private Sub Class_Initialize()
    

    m_eFlags = colRgbinit

End Sub

Public Property Get Color() As Long
    ' Returns: The color

    Color = m_lngColor

End Property

Public Property Let Color(ByVal lngValue As Long)
    ' lngValue: Sets the color

    m_lngColor = lngValue

End Property

Public Property Get Flags() As EnumColFlags
    ' Returns: The flags

    Flags = m_eFlags

End Property

Public Property Let Flags(ByVal eValue As EnumColFlags)
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

Public Function ShowColor() As Boolean
    ' Comments  : This procedure displays the choose color dialog common dialog
    ' Parameters: None
    ' Returns   : True if OK was selected, False otherwise.
    '
    '
    Dim cc As CHOOSECOLOR
    On Error GoTo PROC_ERR
    ' Initialize the CHOOSECOLOR type
    cc.lStructSize = LenB(cc)
    cc.hwndOwner = m_lnghWndParent
    cc.lpCustColors = VarPtr(m_alngColorRef(0))
    cc.rgbResult = m_lngColor
    cc.Flags = m_eFlags
   
    ' Display the dialog
    If ChooseColor_API(cc) <> 0 Then
        ShowColor = True
        m_lngColor = cc.rgbResult
        
    Else
        ShowColor = False
    End If
    
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "Show"
    Resume PROC_EXIT
    
End Function

Public Property Get CustomColors(ByVal Index As Long) As Long
    CustomColors = m_alngColorRef(Index)
End Property

Public Property Let CustomColors(ByVal Index As Long, ByVal vNewValue As Long)
    m_alngColorRef(Index) = vNewValue
End Property
