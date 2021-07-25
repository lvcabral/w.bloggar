VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*
'*  CLASS INFORMATION
'*
'*  Product       : XML-RPC implementation for COM/MTS
'*  File name     : Client.cls
'*  Class name    : Client
'*
'*  Date          : 10 May 2000
'*  Programmer    : Jan G.P. Sijm
'*
'*  Description   :
'*    This class implements a standard XML-RPC client.
'*
'*  Comments      :
'*    This class will parse the response of an XML-RPC call in order to retrieve
'*    the necessary values (e.g. the faultCode, faultString and the values of param's).
'*    It will only validate the response to make sure it can actually retrieve the
'*    necessary values. It will not perform a complete syntax validation for the
'*    required XML-RPC compliance.
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
'--added by Marcelo Cabral - 18/12/2001
'Default timeout
Private Const HTTP_TIMEOUT = 30000 'milliseconds
'--

'*-------------
'*  User defined types
'*-------------

'*-------------
'*  Member variables
'*-------------
' Use SSL (e.g. HTTPS) to communicate with the XML-RPC server.
Private m_bSecure As Boolean
' The port used to communicate with the XML-RPC server.
Private m_lPort As Long
' The user ID to use for the HTTP request.
Private m_sUserID As String
' The password to use for the HTTP request.
Private m_sPassword As String
'--added by Marcelo Cabral - 08/02/2002
Private m_lTimeout As Long
'--added by Marcelo Cabral - 21/09/2002
' Use Proxy - Optional.
Private m_bUseProxy As Boolean
'--added by Marcelo Cabral - 02/01/2002
' The Proxy server - Optional.
Private m_sProxyServer As String
' The port used to communicate with the Proxy server - Optional.
Private m_lProxyPort As Long
'--
'--added by Marcelo Cabral - 23/06/2003
' The user ID to use for Proxy Authentication.
Private m_sProxyUserID As String
' The password to use for Proxy Authentication.
Private m_sProxyPassword As String
'--

'*-------------
'*  Properties
'*-------------

'*-------------
'*  Events
'*-------------
'
' The Error event will be raised if an error occurs.
' The lNumber and sDescription arguments will contains Err.Number and Err.Description
' The ParseError argument contains a reference to a IXMLDOMParseError object containing
' detailed error information from the XML parser. This argument can also be Nothing
' if the error was not related to the parsing process.
'
Public Event Error(ByVal lNumber As Long, ByVal sDescription As String, ByVal ParseError As IXMLDOMParseError)

'*-------------
'*  External declarations (DLL etc.)
'*-------------

Option Explicit

'*-------------
'*  Use SSL (e.g. HTTPS) to communicate with the XML-RPC server.
'*
'*  Input    : bSecure
'*-------------
Public Property Let Secure(ByVal bSecure As Boolean)
Attribute Secure.VB_Description = "Use SSL (e.g. HTTPS) to communicate with the XML-RPC server."

  m_bSecure = bSecure

End Property

'*-------------
'*  The port used to communicate with the XML-RPC server.
'*
'*  Input    : lPort
'*-------------
Public Property Let Port(ByVal lPort As Long)
Attribute Port.VB_Description = "The port used to communicate with the XML-RPC server."

    '--added by Marcelo Cabral - 17/12/2001
    If lPort > 0 Then
    '--
        m_lPort = lPort
        
    End If

End Property

'--added by Marcelo Cabral - 08/02/2002
'*-------------
'*  The Timeout used to communicate with the XML-RPC server.
'*
'*  Input    : lTimeout
'*-------------
Public Property Let Timeout(ByVal lTimeout As Long)

    m_lTimeout = lTimeout
        
End Property

'*-------------
'*  The user ID to use for the HTTP request.
'*
'*  Input    : sUserID
'*-------------
Public Property Let UserID(ByVal sUserID As String)
Attribute UserID.VB_Description = "The user ID to use for the HTTP request."

  m_sUserID = sUserID

End Property

'*-------------
'*  The password to use for the HTTP request.
'*
'*  Input    : sPassword
'*-------------
Public Property Let Password(ByVal sPassword As String)
Attribute Password.VB_Description = "The password to use for the HTTP request."

  m_sPassword = sPassword

End Property

'--added by Marcelo Cabral - 21/09/2002
'*-------------
'*  A flag to tell if will use a Proxy Server.
'*
'*  Input    : bUseProxy
'*-------------
Public Property Let UseProxy(ByVal bUseProxy As Boolean)

    m_bUseProxy = bUseProxy

End Property
'--

'--added by Marcelo Cabral - 02/01/2002
'*-------------
'*  The proxy server
'*
'*  Input    : sProxyServer
'*-------------
Public Property Let ProxyServer(ByVal sProxyServer As String)

    m_sProxyServer = sProxyServer
    If sProxyServer <> "" Then m_bUseProxy = True
    
End Property
'--

'--added by Marcelo Cabral - 02/01/2002
'*-------------
'*  The port used to communicate with the Proxy server.
'*
'*  Input    : lPort
'*-------------
Public Property Let ProxyPort(ByVal lProxyPort As Long)

    If lProxyPort > 0 Then
        m_lProxyPort = lProxyPort
    End If

End Property

'--added by Marcelo Cabral - 23/06/2003
'*-------------
'*  The user ID to use for the Proxy Authentication.
'*
'*  Input    : sProxyUserID
'*-------------
Public Property Let ProxyUserID(ByVal sProxyUserID As String)

  m_sProxyUserID = sProxyUserID

End Property

'--added by Marcelo Cabral - 23/06/2003
'*-------------
'*  The password to use for the Proxy Authentication.
'*
'*  Input    : sProxyPassword
'*-------------
Public Property Let ProxyPassword(ByVal sProxyPassword As String)

  m_sProxyPassword = sProxyPassword

End Property
'--

'*-------------
'*  This method will execute the specified method using an XML-RPC call.
'*  It will format the request into a standard XML-RPC request and the result
'*  of the call is parsed into a DOMDocument object.
'*
'*  Input    : sHost  , XML-RPC server that will process the XML-RPC call
'*             sPage  , Page that will receive the XML-RPC request
'*             sMethod, The nam of the method to call
'*             vParams, The parameters of the method
'*  Returns  : A DOMDocument object containing the resulting XML if successful
'*-------------
Public Function Execute(ByVal sHost As String, ByVal sPage As String, ByVal sMethod As String, ParamArray vParams() As Variant) As DOMDocument
Attribute Execute.VB_Description = "Execute the specified method using an XML-RPC call and pass the specified parameters."

  Dim DOMDocument As DOMDocument
  Dim sXML As String
  Dim lIndex As Long, lCount As Long
  Dim vParameters() As Variant

  On Error GoTo ErrorHandler

  '
  ' Do a basic validation of the method name
  '
  ValidateMethod sMethod

  '
  ' Format the method and parameters into a standard XML-RPC request
  '
  lCount = UBound(vParams) + 1

  If lCount <> 0 Then
    ReDim vParameters(lCount - 1) As Variant
    For lIndex = 0 To lCount - 1
      If VarType(vParams(lIndex)) = vbObject Then
        Set vParameters(lIndex) = vParams(lIndex)
      Else
        vParameters(lIndex) = vParams(lIndex)
      End If
    Next
  End If

  sXML = FormatRequest(sMethod, vParameters, lCount)

  '
  ' Call the specified XML-RPC on the specified URL
  '
  Set DOMDocument = Post(sHost, sPage, sXML)
  Set Execute = DOMDocument

  '
  ' Parse the result and raise an error if a fault was returned <> 0
  '
  ParseFault DOMDocument

  Exit Function

ErrorHandler:

  Dim lNumber As Long
  Dim sDescription As String
  Dim ParseError As IXMLDOMParseError

  '
  ' Retain and format information about the error
  '
  If Not (DOMDocument Is Nothing) Then
    If DOMDocument.ParseError.errorCode <> 0 Then
      Set ParseError = DOMDocument.ParseError
    End If
  End If
  
  lNumber = Err.Number
  sDescription = Err.Description & ParseErrorString(ParseError)

  '
  ' Inform the client of the error that occured
  '
  RaiseEvent Error(lNumber, sDescription, ParseError)
  Err.Raise lNumber, TypeName(Me) & ".Execute", sDescription

End Function

'*-------------
'*  This function will convert the response of an XML-RPC call (e.g. the Execute method)
'*  to an array of Variant values.
'*
'*  Input    : DOMDocument, The response of an XML-RPC call.
'*             vResponse  , The target for a Variant value containing the values from
'*                          the response.
'*-------------
Public Sub ResponseToVariant(ByVal DOMDocument As DOMDocument, ByRef vResponse As Variant)
Attribute ResponseToVariant.VB_Description = "The response of the XML-RPC call converted into a Variant value."

  Dim vValue As Variant
  Dim Param As IXMLDOMNode
  Dim Value As IXMLDOMNode

  On Error GoTo ErrorHandler

  '
  ' Convert the response into a variant array
  '
  Set Param = DOMDocument.selectSingleNode(".//param")
  
  If Param Is Nothing Then
    '
    ' The response does not contain a <param> tag
    '
    '--added by Marcelo Cabral - 17/07/2002
    DOMDocument.Save gAppDataPath & "\lasterror.xml"
    '--
    Err.Raise vbObjectError + 512, , "The <param> element is missing."
        
  Else

    '
    ' Get the <value> element
    '
    Set Value = Param.selectSingleNode(".//value")

    If Value Is Nothing Then
      '--added by Marcelo Cabral - 17/07/2002
      DOMDocument.Save gAppDataPath & "\lasterror.xml"
      '--
      Err.Raise vbObjectError + 512, , "The <value> element of the <param> element is missing."
    Else

      '
      ' Parse the <value> element into a Variant
      '
      ParseValue Value, vResponse
    End If

  End If

  Exit Sub

ErrorHandler:

  Dim lNumber As Long
  Dim sDescription As String

  '
  ' Retain and format information about the error
  '
  lNumber = Err.Number
  sDescription = Err.Description

  '
  ' Inform the client of the error that occured
  '
  RaiseEvent Error(lNumber, sDescription, Nothing)
  Err.Raise lNumber, TypeName(Me) & ".Response", sDescription

End Sub

'*-------------
'*  This function will post a standard XML-RPC request to the specified URL
'*  and will retrieve, parse and return the reponse of the call.
'*
'*  Input    : sHost  , XML-RPC server that will process the XML-RPC call
'*             sPage  , Page that will receive the XML-RPC request
'*             sXML   , The actual XML-RPC request
'*-------------
Private Function Post(ByVal sHost As String, ByVal sPage As String, ByVal sXML As String) As DOMDocument

  Dim HttpRequest As New xmlHttpRequest
  Dim DOMDocument As New DOMDocument
  Dim sResponse As String

  '
  ' Open a connection to the specified URL and send the XML
  '
  HttpRequest.Content = sXML
  
  '--added by Marcelo Cabral - 18/12/2001
  HttpRequest.Timeout = m_lTimeout
  '--
  '--changed by Marcelo Cabral - 02/01/2002
  '--changed by Marcelo Cabral - 23/06/2003
  HttpRequest.Post sHost, m_lPort, m_bSecure, sPage, m_sUserID, m_sPassword, m_bUseProxy, m_sProxyServer, m_lProxyPort, m_sProxyUserID, m_sProxyPassword
  '--
  '
  ' Parse the XML response of the XML-RPC call
  '
  sResponse = HttpRequest.Response
  '--added by Marcelo Cabral - 12/12/2001
  '--changed by Marcelo Cabral - 10/06/2002
  '--changed by Marcelo Cabral - 08/10/2002
  ' If the XML parser is older then 3.0 eliminate encoding param
  If Not DOMDocument.implementation.hasFeature("XSLT", "1.0") Then
      sResponse = "<?xml version=""1.0""" & Mid(sResponse, InStr(1, sResponse, "?>", vbTextCompare))
  ' If there's garbage before the xml header, ignore it
  ElseIf Left(sResponse, 2) <> "<?" And InStr(1, sResponse, "<?", vbTextCompare) > 0 Then
      sResponse = Mid(sResponse, InStr(1, sResponse, "<?", vbTextCompare))
  End If
  '--
  DOMDocument.loadXML sResponse
  Set Post = DOMDocument

  If DOMDocument.ParseError.errorCode <> 0 Then

    '--added by Marcelo Cabral - 17/07/2002
    SaveBinaryFile gAppDataPath & "\lasterror.xml", sResponse
    '--
    '
    ' Parsing failed, raise an error to the caller
    '
    Err.Raise DOMDocument.ParseError.errorCode, , "Unable to parse the XML response. Parser Reason:" & DOMDocument.ParseError.reason

  End If

End Function

'*-------------
'*  This function will format the specified method name and method parameters
'*  into a standard XML-RPC request.
'*
'*  Input    : sMethod  , The name of the method to call
'*             vParams(), The parameters of the method
'*             lCount   , Number of parameters
'*  Returns  : A standard XML-RPC request
'*-------------
Private Function FormatRequest(ByVal sMethod As String, ByRef vParams() As Variant, ByVal lCount As Long) As String

  Dim sXML As String
  Dim lIndex As Long

  '
  ' Setup the request header
  '
  sXML = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & "?>" & vbCrLf
  sXML = sXML & "<methodCall>" & vbCrLf
  sXML = sXML & "<methodName>" & sMethod & "</methodName>" & vbCrLf
  sXML = sXML & "<params>" & vbCrLf
  
  '
  ' Add the parameters to the request
  '
  If lCount <> 0 Then

    For lIndex = 0 To lCount - 1
      sXML = sXML & "<param>" & vbCrLf
      sXML = sXML & FormatParameter(vParams(lIndex)) & vbCrLf
      sXML = sXML & "</param>" & vbCrLf
    Next

  End If

  sXML = sXML & "</params>" & vbCrLf
  sXML = sXML & "</methodCall>"

  FormatRequest = sXML

End Function

'*-------------
'*  This procedure will parse the result of the XML-RPC call for a <fault> element.
'*  If a <fault> element is found, an error will be raised.
'*
'*  Input    : DOMDocument, The XML result of the XML-RPC call
'*-------------
Private Sub ParseFault(ByVal DOMDocument As DOMDocument)
  
  '--changed by Marcelo Cabral - 02/08/2002
  Dim vCode As Variant
  '--
  Dim sDescription As String
  Dim Fault As IXMLDOMNode, FaultCode As IXMLDOMNode, FaultString As IXMLDOMNode
  Dim Value As IXMLDOMNode

  '
  ' Check for a <fault> element
  '
  Set Fault = DOMDocument.selectSingleNode(".//fault")

  If Not (Fault Is Nothing) Then

    '
    ' Get the <faultCode> and <faultString> elements
    '
    Set FaultCode = Fault.selectSingleNode(".//name[. = ""faultCode""]").parentNode
    Set FaultString = Fault.selectSingleNode(".//name[. = ""faultString""]").parentNode

    If FaultCode Is Nothing Or FaultString Is Nothing Then
      '--added by Marcelo Cabral - 17/07/2002
      DOMDocument.Save gAppDataPath & "\lasterror.xml"
      '--
      Err.Raise vbObjectError + 512, , "The <fault> element is invalid. The <faultCode> or <faultString> element is missing."
    Else

      '
      ' Get the faultCode value
      '
      Set Value = FaultCode.selectSingleNode(".//value")

      If Value Is Nothing Then
        '--added by Marcelo Cabral - 17/07/2002
        DOMDocument.Save gAppDataPath & "\lasterror.xml"
        '--
        Err.Raise vbObjectError + 512, , "The <value> element of the <faultCode> element is missing."
      Else
        ParseValue Value, vCode
      End If

      '
      ' Get the faultString value
      '
      Set Value = FaultString.selectSingleNode(".//value")

      If Value Is Nothing Then
        '--added by Marcelo Cabral - 17/07/2002
        DOMDocument.Save gAppDataPath & "\lasterror.xml"
        '--
        Err.Raise vbObjectError + 512, , "The <value> element of the <faultString> element is missing."
      Else
        ParseValue Value, sDescription
      End If

      '
      ' Raise an error using the faultCode and faultString
      '
      '--changed by Marcelo Cabral - 17/07/2002
      If vCode <> "" Or sDescription <> "" Then
        DOMDocument.Save gAppDataPath & "\lasterror.xml"
        If IsNumeric(vCode) Then
            If vCode = 0 Then vCode = vbObjectError + 512
            Err.Raise vCode, , "Your blog server returned the following error message:" & vbCrLf & vbCrLf & _
                               Trim(Replace(sDescription, "java.lang.Exception:", ""))
        Else
            Err.Raise vbObjectError + 512, , "Your blog server returned the following error message:" & vbCrLf & vbCrLf & _
                               Trim(Replace(sDescription, "java.lang.Exception:", ""))
        End If
      End If
      '--

    End If

  End If

End Sub

'*-------------
'*  This procedure will do a basic validation of the specified method name.
'*
'*  Input    : sMethod, Name of the method
'*-------------
Private Sub ValidateMethod(ByVal sMethod As String)

  Dim bFound As Boolean
  Dim iLength As Integer, iCharacter As Integer

  Const Valid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_.:/"

  '
  ' Do a basic validation of the method name
  '
  bFound = False
  iLength = Len(sMethod)
  iCharacter = 0

  Do While bFound = False And iCharacter < iLength
    If InStr(Valid, Mid$(sMethod, iCharacter + 1, 1)) = 0 Then
      bFound = True
    End If
    iCharacter = iCharacter + 1
  Loop

  If bFound = True Then

    '
    ' An invalid character was found, raise an error
    '
    Err.Raise 380, , "The specified method name contains invalid characters. Only a-z, A-Z, 0-9, underscore, dot, colon and slash are allowed."

  End If

End Sub

Private Sub Class_Initialize()

  '
  ' Set the default values of the properties
  '
  m_lPort = 80
  m_bSecure = False
  m_lTimeout = HTTP_TIMEOUT

End Sub
