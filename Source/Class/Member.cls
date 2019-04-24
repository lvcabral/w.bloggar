VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*
'*  CLASS INFORMATION
'*
'*  Product       : XML-RPC implementation for COM/MTS
'*  File name     : Member.cls
'*  Class name    : Member
'*
'*  Date          : 11 Jun 2000
'*  Programmer    : Jan G.P. Sijm
'*
'*  Description   :
'*    This class represents a <member> element in a <struct> element.
'*
'*  Comments      :
'*
'*

'*-------------
'*  Implements
'*-------------

'*-------------
'*  Enums
'*-------------

'*-------------
'*  Private constants
'*-------------

'*-------------
'*  User defined types
'*-------------

'*-------------
'*  Member variables
'*-------------
' The value of the structure member
Private m_vValue As Variant
' The name of the structure member
Private m_sName As String

'*-------------
'*  Properties
'*-------------

'*-------------
'*  Events
'*-------------

'*-------------
'*  External declarations (DLL etc.)
'*-------------

Option Explicit

'*-------------
'*  The value of the structure member
'*
'*  Input    : None
'*  Returns  : The value of the Value property
'*-------------
Public Property Get Value() As Variant
Attribute Value.VB_Description = "The value of the structure member"

  If VarType(m_vValue) = vbObject Then
    Set Value = m_vValue
  Else
    Value = m_vValue
  End If

End Property

'*-------------
'*  Set the value of the structure member
'*
'*  Input    : vValue, The value of the Value property
'*-------------
Friend Property Let Value(ByVal vValue As Variant)

  m_vValue = vValue

End Property

'*-------------
'*  Set the value of the structure member
'*
'*  Input    : vValue, The value of the Value property
'*-------------
Friend Property Set Value(ByVal vValue As Variant)

  Set m_vValue = vValue

End Property

'*-------------
'*  The name of the structure member
'*
'*  Input    : None
'*  Returns  : The value of the Name property
'*-------------
Public Property Get Name() As String
Attribute Name.VB_Description = "The name of the structure member"

  Name = m_sName

End Property

'*-------------
'*  Set the name of the structure member
'*
'*  Input    : sName, The value of the Name property
'*-------------
Friend Property Let Name(ByVal sName As String)

  m_sName = sName

End Property

