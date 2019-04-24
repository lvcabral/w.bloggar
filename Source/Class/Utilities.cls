VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*
'*  CLASS INFORMATION
'*
'*  Product       : XML-RPC implementation for COM/MTS
'*  File name     : Utilities.cls
'*  Class name    : Utilities
'*
'*  Date          : 11 Jun 2000
'*  Programmer    : Jan G.P. Sijm
'*
'*  Description   :
'*    This class implements a number of helper methods.
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
'*  This method will be called whenever an error occurs. It will construct a
'*  <fault> XML-RPC response.
'*
'*  Input    : lNumber     , The error number
'*             sDescription, A detailed error description
'*  Returns  : The <fault> structure
'*-------------
Public Function FormatFault(ByVal lNumber As Long, ByVal sDescription As String) As String
Attribute FormatFault.VB_Description = "Create an XML-RPC <fault> element"

  Dim sXML As String

  '
  ' Setup the request header
  '
  sXML = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & "?>" & vbCrLf
  sXML = sXML & "<methodResponse>" & vbCrLf
  sXML = sXML & "<fault>" & vbCrLf
  sXML = sXML & "<value>" & vbCrLf
  sXML = sXML & "<struct>" & vbCrLf

  '
  ' Add the <faultCode> element
  '
  sXML = sXML & "<member>" & vbCrLf
  sXML = sXML & "<name>faultCode</name>" & vbCrLf
  sXML = sXML & "<value><int>" & CStr(lNumber) & "</int></value>" & vbCrLf
  sXML = sXML & "</member>" & vbCrLf
  
  '
  ' Add the <faultString> element
  '
  sXML = sXML & "<member>" & vbCrLf
  sXML = sXML & "<name>faultString</name>" & vbCrLf
  sXML = sXML & "<value><string>" & Escape(sDescription) & "</string></value>" & vbCrLf
  sXML = sXML & "</member>" & vbCrLf
  
  sXML = sXML & "</struct>" & vbCrLf
  sXML = sXML & "</value>" & vbCrLf
  sXML = sXML & "</fault>" & vbCrLf
  sXML = sXML & "</methodResponse>"

  FormatFault = sXML

End Function

'*-------------
'*  This function will replace characters that have special meaning by the
'*  corresponding escape characters.
'*
'*  Input    : sValue, String value to be escaped
'*  Returns  : String value containing escape characters
'*-------------
Public Function Escape(ByVal sValue As String) As String
Attribute Escape.VB_Description = "Translate special characters to escape sequence"

'--changed by Marcelo Cabral - 04/01/2002
  sValue = Replace(sValue, "&", "&amp;")
  sValue = Replace(sValue, "<", "&lt;")
  sValue = Replace(sValue, ">", "&gt;")
  sValue = Replace(sValue, "'", "&apos;")
  sValue = Replace(sValue, Chr$(34), "&quot;")
'--
  Escape = sValue

End Function

'*-------------
'*  This function will replace escape characters by the corresponding values.
'*
'*  Input    : sValue, String value to be un-escaped
'*  Returns  : String value without escape characters
'*-------------
Public Function Unescape(ByVal sValue As String) As String
Attribute Unescape.VB_Description = "Translate escape sequence to special character"

'--changed by Marcelo Cabral - 04/01/2002
  sValue = Replace(sValue, "&lt;", "<")
  sValue = Replace(sValue, "&gt;", ">")
  sValue = Replace(sValue, "&amp;", "&")
  sValue = Replace(sValue, "&apos;", "'")
  sValue = Replace(sValue, "&quot;", Chr$(34))
'--
  Unescape = sValue

End Function

'--commented by Marcelo Cabral - 04/01/2002
'*-------------
'*  This function will replace all occurences of a string by a replacement string.
'*  The replacement string may have a different length than the string to be replaced.
'*
'*  Input    : sValue    , Source string
'*             sToReplace, String to replace
'*             sReplaceBy, String to replace by
'*  Return   : Modified string
'*-------------
'Private Function StringReplace(ByVal sValue As String, ByVal sToReplace As String, ByVal sReplaceBy As String) As String
'
'  Dim iNdx As Integer, iR As Integer, iB As Integer
'  Dim sLeft As String, sRight As String
'
'  '
'  ' Initialize the replace loop
'  '
'  iR = Len(sToReplace)
'  iB = Len(sReplaceBy)
'
'  '
'  ' Find first occurence of string to replace
'  '
'  iNdx = InStr(sValue, sToReplace)
'
'  Do While iNdx <> 0
'
'    '
'    ' Replace string
'    '
'    sLeft = Left$(sValue, iNdx - 1)
'    sRight = Mid$(sValue, iNdx + iR)
'    sValue = sLeft + sReplaceBy + sRight
'    iNdx = iNdx + iB
'
'    '
'    ' Find next occurence of string to replace
'    '
'    iNdx = InStr(iNdx, sValue, sToReplace)
'
'  Loop
'
'  StringReplace = sValue
'
'End Function

