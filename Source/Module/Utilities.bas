Attribute VB_Name = "basUtilities"
'*
'*  MODULE INFORMATION
'*
'*  Product       : XML-RPC implementation for COM/MTS
'*  File name     : basUtilities.bas
'*
'*  Date          : 10 May 2000
'*  Programmer    : Jan G.P. Sijm
'*
'*  Description   :
'*    This module contains several generic routine used by the XML-RPC
'*    client or server components.
'*
Option Explicit
Private Const INTERNET_ERROR_BASE = 12000
Private Const ERROR_INTERNET_OUT_OF_HANDLES As Long = (INTERNET_ERROR_BASE + 1)
Private Const ERROR_INTERNET_TIMEOUT As Long = (INTERNET_ERROR_BASE + 2)
Private Const ERROR_INTERNET_EXTENDED_ERROR As Long = (INTERNET_ERROR_BASE + 3)
Private Const ERROR_INTERNET_INTERNAL_ERROR As Long = (INTERNET_ERROR_BASE + 4)
Private Const ERROR_INTERNET_INVALID_URL As Long = (INTERNET_ERROR_BASE + 5)
Private Const ERROR_INTERNET_UNRECOGNIZED_SCHEME As Long = (INTERNET_ERROR_BASE + 6)
Private Const ERROR_INTERNET_NAME_NOT_RESOLVED As Long = (INTERNET_ERROR_BASE + 7)
Private Const ERROR_INTERNET_PROTOCOL_NOT_FOUND As Long = (INTERNET_ERROR_BASE + 8)
Private Const ERROR_INTERNET_INVALID_OPTION As Long = (INTERNET_ERROR_BASE + 9)
Private Const ERROR_INTERNET_BAD_OPTION_LENGTH As Long = (INTERNET_ERROR_BASE + 10)
Private Const ERROR_INTERNET_OPTION_NOT_SETTABLE As Long = (INTERNET_ERROR_BASE + 11)
Private Const ERROR_INTERNET_SHUTDOWN As Long = (INTERNET_ERROR_BASE + 12)
Private Const ERROR_INTERNET_INCORRECT_USER_NAME As Long = (INTERNET_ERROR_BASE + 13)
Private Const ERROR_INTERNET_INCORRECT_PASSWORD As Long = (INTERNET_ERROR_BASE + 14)
Private Const ERROR_INTERNET_LOGIN_FAILURE As Long = (INTERNET_ERROR_BASE + 15)
Private Const ERROR_INTERNET_INVALID_OPERATION As Long = (INTERNET_ERROR_BASE + 16)
Private Const ERROR_INTERNET_OPERATION_CANCELLED As Long = (INTERNET_ERROR_BASE + 17)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE As Long = (INTERNET_ERROR_BASE + 18)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_STATE As Long = (INTERNET_ERROR_BASE + 19)
Private Const ERROR_INTERNET_NOT_PROXY_REQUEST As Long = (INTERNET_ERROR_BASE + 20)
Private Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND As Long = (INTERNET_ERROR_BASE + 21)
Private Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER As Long = (INTERNET_ERROR_BASE + 22)
Private Const ERROR_INTERNET_NO_DIRECT_ACCESS As Long = (INTERNET_ERROR_BASE + 23)
Private Const ERROR_INTERNET_NO_CONTEXT As Long = (INTERNET_ERROR_BASE + 24)
Private Const ERROR_INTERNET_NO_CALLBACK As Long = (INTERNET_ERROR_BASE + 25)
Private Const ERROR_INTERNET_REQUEST_PENDING As Long = (INTERNET_ERROR_BASE + 26)
Private Const ERROR_INTERNET_INCORRECT_FORMAT As Long = (INTERNET_ERROR_BASE + 27)
Private Const ERROR_INTERNET_ITEM_NOT_FOUND As Long = (INTERNET_ERROR_BASE + 28)
Private Const ERROR_INTERNET_CANNOT_CONNECT As Long = (INTERNET_ERROR_BASE + 29)
Private Const ERROR_INTERNET_CONNECTION_ABORTED As Long = (INTERNET_ERROR_BASE + 30)
Private Const ERROR_INTERNET_CONNECTION_RESET As Long = (INTERNET_ERROR_BASE + 31)
Private Const ERROR_INTERNET_FORCE_RETRY As Long = (INTERNET_ERROR_BASE + 32)
Private Const ERROR_INTERNET_INVALID_PROXY_REQUEST As Long = (INTERNET_ERROR_BASE + 33)
Private Const ERROR_INTERNET_NEED_UI As Long = (INTERNET_ERROR_BASE + 34)
Private Const ERROR_INTERNET_HANDLE_EXISTS As Long = (INTERNET_ERROR_BASE + 36)
Private Const ERROR_INTERNET_SEC_CERT_DATE_INVALID As Long = (INTERNET_ERROR_BASE + 37)
Private Const ERROR_INTERNET_SEC_CERT_CN_INVALID As Long = (INTERNET_ERROR_BASE + 38)
Private Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR As Long = (INTERNET_ERROR_BASE + 39)
Private Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR As Long = (INTERNET_ERROR_BASE + 40)
Private Const ERROR_INTERNET_MIXED_SECURITY As Long = (INTERNET_ERROR_BASE + 41)
Private Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE As Long = (INTERNET_ERROR_BASE + 42)
Private Const ERROR_INTERNET_POST_IS_NON_SECURE As Long = (INTERNET_ERROR_BASE + 43)
Private Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED As Long = (INTERNET_ERROR_BASE + 44)
Private Const ERROR_INTERNET_INVALID_CA As Long = (INTERNET_ERROR_BASE + 45)
Private Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP As Long = (INTERNET_ERROR_BASE + 46)
Private Const ERROR_INTERNET_ASYNC_THREAD_FAILED As Long = (INTERNET_ERROR_BASE + 47)
Private Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE As Long = (INTERNET_ERROR_BASE + 48)
Private Const ERROR_INTERNET_DIALOG_PENDING As Long = (INTERNET_ERROR_BASE + 49)
Private Const ERROR_INTERNET_RETRY_DIALOG As Long = (INTERNET_ERROR_BASE + 50)
Private Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR As Long = (INTERNET_ERROR_BASE + 52)
Private Const ERROR_INTERNET_INSERT_CDROM As Long = (INTERNET_ERROR_BASE + 53)

'*-------------
'*  This function will format the specified parameter into a standard XML-RPC
'*  fragment. The function is called recursive if the passed parameter is
'*  an object or array.
'*
'*  Input    : vParam, The parameter to be formatted
'*  Returns  : The formatted XML-RPC parameter
'*-------------
Public Function FormatParameter(ByVal vParam As Variant) As String

  Dim sValue As String
  Dim Base64Codec As xmlBase64Codec
  Dim Utilities As New xmlUtilities

  '
  ' Format the parameter
  '
  Select Case VarType(vParam)

    '
    ' Format the supported basic data types
    '
    Case vbEmpty
      sValue = "<string></string>"

    Case vbNull
      sValue = "<string></string>"

    Case vbInteger
      sValue = "<int>" & CStr(vParam) & "</int>"

    Case vbLong
      sValue = "<int>" & CStr(vParam) & "</int>"

    Case vbSingle
      sValue = "<double>" & CStr(vParam) & "</double>"

    Case vbDouble
      sValue = "<double>" & Str(vParam) & "</double>"

    Case vbCurrency
      sValue = "<double>" & CStr(vParam) & "</double>"

    Case vbDate
      sValue = "<dateTime.iso8601>" & Format(CDate(vParam), "YYYYMMDD") & "T" & Format(CDate(vParam), "HH:MM:SS") & "</dateTime.iso8601>"

    Case vbString
      sValue = "<string>" & Utilities.Escape(vParam) & "</string>"

    Case vbBoolean
      sValue = "<boolean>" & CStr(Abs(CInt(vParam))) & "</boolean>"

    Case vbByte
      sValue = "<int>" & CStr(vParam) & "</int>"

    Case vbObject
    
      If TypeName(vParam) = "xmlStruct" Then

          '
          ' Format the collection as a <struct> element
          '
          sValue = FormatStruct(vParam)

      Else

          '
          ' Unsupported object type
          '
          Err.Raise 13, , "An unsupported object type was passed as an XML-RPC parameter."

      End If

    Case Else

      If VarType(vParam) > vbArray Then

        If VarType(vParam) - vbArray = vbByte Then

          '
          ' Format the array as a <base64> element
          '
          Set Base64Codec = New xmlBase64Codec
          sValue = "<base64>" & Base64Codec.EncodeArray((vParam)) & "</base64>"

        Else

          '
          ' Format the array as an <array> element
          '
          sValue = FormatArray(vParam)

        End If

      Else

        '
        ' Unsupported data type
        '
        Err.Raise 13, , "An unsupported data type was passed as an XML-RPC parameter."

      End If

  End Select

  FormatParameter = "<value>" & vbCrLf & sValue & vbCrLf & "</value>"

End Function

'*-------------
'*  This function will format the specified parameter into a standard XML-RPC
'*  <array> element.
'*
'*  Input    : vParam, The parameter to be formatted
'*  Returns  : The formatted <array> element
'*-------------
Public Function FormatArray(ByVal vParam As Variant) As String

  Dim sValue As String
  Dim lIndex As Long, lUBound As Long

  '
  ' Format an array
  '
  lUBound = UBound(vParam)
  sValue = "<array>" & vbCrLf & "<data>" & vbCrLf

  '
  ' Format all items of the array
  '
  For lIndex = 0 To lUBound
    sValue = sValue & FormatParameter(vParam(lIndex)) & vbCrLf
  Next

  sValue = sValue & "</data>" & vbCrLf & "</array>"

  FormatArray = sValue

End Function

'*-------------
'*  This function will format the specified parameter into a standard XML-RPC
'*  <struct> element.
'*
'*  Input    : Struct, The structure to be formatted
'*  Returns  : The formatted <struct> element
'*-------------
Public Function FormatStruct(ByVal Struct As xmlStruct) As String

  Dim sValue As String
  Dim lIndex As Long, lUBound As Long

  '
  ' Format a struct
  '
  sValue = "<struct>" & vbCrLf

  '
  ' Format all items of the collection
  '
  For lIndex = 1 To Struct.Count

    '
    ' Format this item into a <member>
    '
    sValue = sValue & "<member>" & vbCrLf
    sValue = sValue & "<name>" & Struct.Member(lIndex).Name & "</name>" & vbCrLf
    sValue = sValue & FormatParameter(Struct.Member(lIndex).Value) & vbCrLf
    sValue = sValue & "</member>" & vbCrLf

  Next

  sValue = sValue & "</struct>"

  FormatStruct = sValue

End Function

'*-------------
'*  Format the properties of an IXMLDOMParseError object into a string that
'*  can be displayed or logged.
'*
'*  Input    : ParseError, The error object returned from a DOMDocument
'*  Returns  : A formatted error message
'*-------------
Public Function ParseErrorString(ByVal ParseError As IXMLDOMParseError) As String

  Dim sError As String

  '
  ' Format the error information
  '
  If Not (ParseError Is Nothing) Then

    sError = vbCrLf
    sError = sError & "Microsoft XML Parser error object (IXMLDOMParseError) : " & vbCrLf
    sError = sError & "errorCode = " & Hex$(ParseError.errorCode) & vbCrLf
    sError = sError & "filepos = " & CStr(ParseError.filepos) & vbCrLf
    sError = sError & "line = " & CStr(ParseError.Line) & vbCrLf
    sError = sError & "linepos = " & CStr(ParseError.linepos) & vbCrLf
    sError = sError & "reason = " & ParseError.reason & vbCrLf
    sError = sError & "srcText = " & ParseError.srcText & vbCrLf
    sError = sError & "url = " & ParseError.URL & vbCrLf

  End If

  ParseErrorString = sError

End Function

'*-------------
'*  This function will parse the specified Node containing a <value></value> element pair.
'*  The function is called recursive if the passed Node contains an object or an array.
'*
'*  Input    : Node  , A node containing a <value> element.
'*             vValue, Target for the value
'*-------------
Public Sub ParseValue(ByVal Node As IXMLDOMNode, ByRef vValue As Variant)

  Dim vArray() As Variant
  Dim dtDateTime As Date
  Dim Value As IXMLDOMNode
  Dim sName As String
  Dim xBytes() As Byte
  Dim Base64Codec As xmlBase64Codec
  Dim Utilities As New xmlUtilities
  
  If Not Node.hasChildNodes Then
  
    '
    ' An empty <value> element
    '
    sName = "string"
  
  Else
  
    If Node.childNodes(0).nodeType = NODE_TEXT Then
    
      '
      ' An element without data type tags
      '
      sName = "string"
      Set Value = Node.childNodes(0)
    
    Else
    
      '
      ' An element with data type tags
      '
      sName = Node.childNodes(0).nodeName
      Set Value = Node.childNodes(0).childNodes(0)
    
    End If
  
  End If
  
  Select Case sName

    '
    ' Parse the supported basic data types
    '
    Case "string"
      If Not (Value Is Nothing) Then
        
        '--changed by Marcelo Cabral - 21/12/2001
        'vValue = Utilities.Unescape(CStr(Value.nodeValue))
        
        '--changed by Marcelo Cabral - 13/02/2002
        If IsNull(Value.nodeValue) Then
            vValue = ""
        Else
            vValue = CStr(Value.nodeValue)
        End If
        '--
      Else
        vValue = ""
      End If

    Case "int", "i4"
      '--changed by Marcelo Cabral - 19/02/2003
      If Not (Value Is Nothing) Then
          vValue = CLng(Value.nodeValue)
      Else
          vValue = CLng(0)
      End If
      '--
    Case "double"
      '--changed by Marcelo Cabral - 19/02/2003
      If Not (Value Is Nothing) Then
          vValue = CDbl(Val(Value.nodeValue))
      Else
          vValue = CDbl(0)
      End If
      '--
    Case "dateTime.iso8601"
      '--changed by Marcelo Cabral - 19/02/2003
      If Not (Value Is Nothing) Then
          vValue = CStr(Value.nodeValue)
      Else
          vValue = ""
      End If
      '--
      '--changed by Marcelo Cabral - 13/02/2002
      If Len(vValue) = 17 Then
        dtDateTime = DateSerial(Mid$(vValue, 1, 4), Mid$(vValue, 5, 2), Mid$(vValue, 7, 2))
        dtDateTime = dtDateTime + TimeSerial(Mid$(vValue, 10, 2), Mid$(vValue, 13, 2), Mid$(vValue, 16, 2))
        vValue = dtDateTime
      '--added by Marcelo Cabral - 05/07/2004 - changed 23/10/2007
      ElseIf Len(vValue) = 19 Or Len(vValue) = 20 Then
        dtDateTime = DateSerial(Mid$(vValue, 1, 4), Mid$(vValue, 6, 2), Mid$(vValue, 9, 2))
        dtDateTime = dtDateTime + TimeSerial(Mid$(vValue, 12, 2), Mid$(vValue, 15, 2), Mid$(vValue, 18, 2))
        vValue = dtDateTime
      Else
        vValue = CDate(0)
      End If
      '--

    Case "boolean"
      '--changed by Marcelo Cabral - 19/02/2003
      If Not (Value Is Nothing) Then
         vValue = CBool(CInt(Value.nodeValue) = 1)
      Else
         vValue = False
      End If
      '--
    Case "struct"
      Set vValue = ParseStruct(Node)

    Case "array"
      ParseArray Node, vArray
      vValue = vArray

    Case "base64"
      Set Base64Codec = New xmlBase64Codec
      xBytes = Base64Codec.DecodeArray(Value.nodeValue)
      vValue = xBytes

    Case Else
      '
      ' Unsupported data type
      '
      Err.Raise 13, , "The <value> element contains an unsupported data type element."

  End Select

End Sub

'*-------------
'*  This function will parse the specified Node containing a <value></value> element pair
'*  and determine the data type of the resulting value.
'*
'*  Input    : Node, A node containing a <value> element.
'*  Returns  : An integer containing the data type
'*-------------
Public Function ParseValueType(ByVal Node As IXMLDOMNode) As Integer

  Dim iDataType As Integer

  If Not Node.hasChildNodes Then

    '
    ' The values does not contain data type tags
    '
    Err.Raise vbObjectError + 512, , "The data type element of the <value> element is missing."

  Else

    '
    ' Parse the value
    '
    Select Case Node.childNodes(0).nodeName

      '
      ' Parse the supported basic data types
      '
      Case "string"
        iDataType = vbString

      Case "int", "i4"
        iDataType = vbInteger

      Case "double"
        iDataType = vbDouble

      Case "dateTime.iso8601"
        iDataType = vbDate

      Case "boolean"
        iDataType = vbBoolean

      Case "struct"
        iDataType = vbObject

      Case "array"
        iDataType = vbVariant + vbArray

      Case "base64"
        iDataType = vbByte + vbArray

      Case Else
        '
        ' Unsupported data type
        '
        Err.Raise 13, , "The <value> element contains an unsupported data type element."

    End Select

  End If

  ParseValueType = iDataType

End Function

'*-------------
'*  This function will parse the specified Node containing a <value></value> element pair,
'*  which contains an <array> element. The value will be added to an array of Variant's
'*
'*  Input    : Node  , A node containing a <value> element.
'*             vArray, The target array of Variant's
'*-------------
Public Sub ParseArray(ByVal Node As IXMLDOMNode, ByRef vArray() As Variant)

  Dim NodeList As IXMLDOMNodeList
  Dim Value As IXMLDOMNode, DataNode As IXMLDOMNode
  Dim lIndex As Long, lItem As Long

  '
  ' Parse an array
  '
  Set DataNode = Node.selectSingleNode(".//data")

  If DataNode Is Nothing Then
    Err.Raise vbObjectError + 512, , "The <data> element of the <array> element is missing."
  Else

    Set NodeList = Node.selectNodes(".//value")

    If NodeList Is Nothing Then
      Err.Raise vbObjectError + 512, , "The <value> elements of the <array> element are missing."
    Else

      '
      ' Parse all <value> elements
      '
      lItem = 0
      For lIndex = 0 To NodeList.length - 1

        Set Value = NodeList.Item(lIndex)

        If Value.parentNode Is DataNode Then

          '
          ' Store the value in the Variant array
          '
          ReDim Preserve vArray(lItem) As Variant
          ParseValue Value, vArray(lItem)
          lItem = lItem + 1

        End If

      Next

    End If

  End If

End Sub

'*-------------
'*  This function will parse the specified Node containing a <value></value> element pair,
'*  which contains a <struct> element. The name - value pairs of the structure will be
'*  added to a Collection.
'*
'*  Input    : Node, A node containing a <value> element.
'*  Returns  : A Collection object containing the actual values
'*-------------
Public Function ParseStruct(ByVal Node As IXMLDOMNode) As xmlStruct

  Dim vValue As Variant, vArray() As Variant
  Dim NodeList As IXMLDOMNodeList
  Dim Value As IXMLDOMNode, Name As IXMLDOMNode, Member As IXMLDOMNode
  Dim Struct As New xmlStruct
  Dim lIndex As Long
  Dim sName As String

  '
  ' Parse the <struct> element
  '
  Set NodeList = Node.selectNodes(".//member")

  If NodeList Is Nothing Then
    Err.Raise vbObjectError + 512, , "The <member> elements of the <struct> element are missing."
  Else

    '
    ' Parse all <member> elements
    '
    For lIndex = 0 To NodeList.length - 1

      Set Member = NodeList.Item(lIndex)

      If Member.parentNode.parentNode Is Node Then

        '
        ' Parse the <name> element
        '
        Set Name = Member.selectSingleNode(".//name")
  
        If Name Is Nothing Then
          Err.Raise vbObjectError + 512, , "The <name> element of the <member> element is missing."
        Else
          sName = CStr(Name.childNodes(0).nodeValue)
        End If
  
        '
        ' Parse the <value> element
        '
        Set Value = Member.selectSingleNode(".//value")
  
        If Value Is Nothing Then
          Err.Raise vbObjectError + 512, , "The <value> element of the <member> element is missing."
        Else
          ParseValue Value, vValue
        End If
  
        '
        ' Store the value in the Dictionary
        '
        Struct.Add sName, vValue
      
      End If

    Next

  End If

  Set ParseStruct = Struct

End Function

Public Function TranslateWinsockError(ByVal lngError As Long) As String
    Select Case lngError
    Case ERROR_INTERNET_TIMEOUT
        TranslateWinsockError = "The request has timed out."
    Case ERROR_INTERNET_INTERNAL_ERROR
        TranslateWinsockError = "An internal error has occurred."
    Case ERROR_INTERNET_INVALID_URL
        TranslateWinsockError = "The URL is invalid."
    Case ERROR_INTERNET_UNRECOGNIZED_SCHEME
        TranslateWinsockError = "The URL scheme could not be recognized, or is not supported."
    Case ERROR_INTERNET_NAME_NOT_RESOLVED
        TranslateWinsockError = "The server name could not be resolved."
    Case ERROR_INTERNET_PROTOCOL_NOT_FOUND
        TranslateWinsockError = "The requested protocol could not be located."
    Case ERROR_INTERNET_SHUTDOWN
        TranslateWinsockError = "The Win32 Internet function support is being shut down or unloaded."
    Case ERROR_INTERNET_INCORRECT_USER_NAME
        TranslateWinsockError = "The supplied user name is incorrect."
    Case ERROR_INTERNET_INCORRECT_PASSWORD
        TranslateWinsockError = "The supplied password is incorrect."
    Case ERROR_INTERNET_LOGIN_FAILURE
        TranslateWinsockError = "The request to connect and log on to an FTP server failed."
    Case ERROR_INTERNET_INVALID_OPERATION
        TranslateWinsockError = "The requested operation is invalid."
    Case ERROR_INTERNET_OPERATION_CANCELLED
        TranslateWinsockError = "The operation was canceled."
    Case ERROR_INTERNET_NOT_PROXY_REQUEST
        TranslateWinsockError = "The request cannot be made via a proxy."
    Case ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND
        TranslateWinsockError = "A required registry value could not be located."
    Case ERROR_INTERNET_BAD_REGISTRY_PARAMETER
        TranslateWinsockError = "A required registry value is an incorrect type or has an invalid value."
    Case ERROR_INTERNET_NO_DIRECT_ACCESS
        TranslateWinsockError = "Direct network access cannot be made at this time."
    Case ERROR_INTERNET_REQUEST_PENDING
        TranslateWinsockError = "One or more requests are pending."
    Case ERROR_INTERNET_INCORRECT_FORMAT
        TranslateWinsockError = "The format of the request is invalid."
    Case ERROR_INTERNET_ITEM_NOT_FOUND
        TranslateWinsockError = "The requested item could not be located."
    Case ERROR_INTERNET_CANNOT_CONNECT
        TranslateWinsockError = "The attempt to connect to the server failed."
    Case ERROR_INTERNET_CONNECTION_ABORTED
        TranslateWinsockError = "The connection with the server has been terminated."
    Case ERROR_INTERNET_CONNECTION_RESET
        TranslateWinsockError = "The connection with the server has been reset."
    Case ERROR_INTERNET_FORCE_RETRY
        TranslateWinsockError = "The Win32 Internet function needs to redo the request."
    Case ERROR_INTERNET_INVALID_PROXY_REQUEST
        TranslateWinsockError = "The request to the proxy was invalid."
    Case Else
        TranslateWinsockError = ""
    End Select
End Function

'--added by Marcelo Cabral - 23/12/2003
Public Function UInt2SInt(ByVal Value As Variant) As Integer
Const IntOffset = 65536
Const MaxInt = 32767
On Error GoTo Error:
    If Value >= IntOffset Then Value = IntOffset - 1
    If Value < (0 - MaxInt) - 1 Then Value = (0 - MaxInt) - 1
    If Value <= MaxInt Then
        UInt2SInt = Value
    Else
        UInt2SInt = Value - IntOffset
    End If
Exit Function
Error:
    UInt2SInt = Value
End Function

'--added by Marcelo Cabral - 23/12/2003
Public Function SInt2UInt(ByVal Value As Variant) As Long
Const IntOffset = 65536
Const MaxInt = 32767
On Error GoTo Error:
    If Value > MaxInt Then Value = MaxInt
    If Value < 0 Then
        SInt2UInt = Value + IntOffset
    Else
        SInt2UInt = Value
    End If
Exit Function
Error:
    SInt2UInt = Value
End Function

