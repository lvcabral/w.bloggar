VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*
'*  CLASS INFORMATION
'*
'*  Product       : XML-RPC implementation for COM/MTS
'*  File name     : Struct.cls
'*  Class name    : Struct
'*
'*  Date          : 11 Jun 2000
'*  Programmer    : Jan G.P. Sijm
'*
'*  Description   :
'*    This class represents a <struct> element.
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
Private m_Collection As New Collection

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
'*  This method will create a new member object and add this object to the collection.
'*
'*  Input    : sName , The name of the member
'*             vValue, The value of the member
'*  Returns  : A reference to the new member object
'*-------------
Public Function Add(ByVal sName As String, ByVal vValue As Variant) As xmlMember
Attribute Add.VB_Description = "Add a new member to the structure"

  Dim Member As New xmlMember
  
  '
  ' Initialize the new member
  '
  Member.Name = sName
  
  If VarType(vValue) = vbObject Then
    Set Member.Value = vValue
  Else
    Member.Value = vValue
  End If
  
  '
  ' Add the new member to the collection
  '
  m_Collection.Add Member, sName

End Function

'*-------------
'*  This procedure will get the value of the 'Count' property.
'*  This property contains the number of items in the collection.
'*
'*  Input    : None
'*  Return   : Number of items
'*-------------
Public Property Get Count() As Long
Attribute Count.VB_Description = "Get the number of members in the structure."

  '
  ' Return the number of items in the collection
  '
  Count = m_Collection.Count

End Property

'*-------------
'*  This function will get the member identified by the specified key
'*  from the collection. Raises an error is the specified member is
'*  not found in the collection.
'*
'*  Input    : vKey, Key value identifying the requested object
'*  Return   : Requested member
'*-------------
Public Function Member(ByVal vKey As Variant) As xmlMember
Attribute Member.VB_Description = "Get the specified member of the structure"

  '
  ' Return the requested member from the collection
  '
  
  Set Member = m_Collection.Item(vKey)
End Function
