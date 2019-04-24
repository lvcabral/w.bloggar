VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlHttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*
'*  CLASS INFORMATION
'*
'*  Product       : XML-RPC implementation for COM/MTS
'*  File name     : HttpRequest.cls
'*  Class name    : HttpRequest
'*
'*  Date          : 12 Oct 2000
'*  Programmer    : Jan G.P. Sijm
'*
'*  Description   :
'*    The HttpRequest class can be used to execute an HTTP or HTTPS
'*    request. It uses the WinInet.dll component to do this.
'*
'*  Comments      :
'*    At this moment, the class only supports an HTTP POST.
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
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
'--added by Marcelo Cabral - 02/02/2002
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_PROXY = 3
'--
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_SECURE = &H800000
Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Private Const INTERNET_OPTION_SEND_TIMEOUT = 5
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Private Const INTERNET_OPTION_PROXY_USERNAME = 43
Private Const INTERNET_OPTION_PROXY_PASSWORD = 44
Private Const HTTP_QUERY_STATUS_CODE = 19

'*-------------
'*  User defined types
'*-------------

'*-------------
'*  Member variables
'*-------------
' The timeout for 'connect', 'send' and 'receive' communications.
' Specify the timeout in miliseconds.
Private m_lTimeout As Long
' The response returned by the server.
Private m_sResponse As String
' The content type of the content.
Private m_sContentType As String
' The content of the request.
Private m_sContent As String

'*-------------
'*  Properties
'*-------------

'*-------------
'*  Events
'*-------------

'*-------------
'*  External declarations (DLL etc.)
'*-------------
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Long, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, lpBuffer As Any, ByVal dwBufferLength As Long) As Integer
Private Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'--commented by Marcelo Cabral - 17/12/2001
'--changed to Err.LastDLLError
'Private Declare Function GetLastError Lib "kernel32" () As Long
'--
Option Explicit

'*-------------
'*  The timeout for 'connect', 'send' and 'receive' communications.
'*  Specify the timeout in miliseconds.
'*
'*  Input    : lTimeout
'*-------------
Public Property Let Timeout(ByVal lTimeout As Long)
Attribute Timeout.VB_Description = "The timeout for 'connect', 'send' and 'receive' communications.\r\nSpecify the timeout in miliseconds."

  m_lTimeout = lTimeout

End Property

'*-------------
'*  The response returned by the server.
'*
'*  Input    : None
'*  Returns  : The value of the Response property
'*-------------
Public Property Get Response() As String
Attribute Response.VB_Description = "The response returned by the server."

  Response = m_sResponse

End Property

'*-------------
'*  The content type of the content.
'*  Set this property if you want to POST content.
'*
'*  Input    : sContentType
'*-------------
Public Property Let ContentType(ByVal sContentType As String)

  m_sContentType = sContentType

End Property

'*-------------
'*  The content type of the content.
'*  Get this property after a GET.
'*
'*  Input    : None
'*  Returns  : The value of the ContentType property
'*-------------
Public Property Get ContentType() As String
Attribute ContentType.VB_Description = "The content type of the content."

  ContentType = m_sContentType

End Property

'*-------------
'*  The content of the request.
'*  Set this property if you want to POST content.
'*
'*  Input    : sContent
'*-------------
Public Property Let Content(ByVal sContent As String)

  m_sContent = sContent

End Property

'*-------------
'*  The content of the request.
'*  Get this property after a GET.
'*
'*  Input    : None
'*  Returns  : The value of the Content property
'*-------------
Public Property Get Content() As String
Attribute Content.VB_Description = "The content of the request."

  Content = m_sContent

End Property

'*-------------
'*  This method will post the value of the Content property to the
'*  specified host. If a user ID and password have been specified, they
'*  will be passed to the host. If the content type of the actual content
'*  is different than text/html, you must set the ContentType property.
'*
'*  Input    : sHost    , The name of the host (without http etc.)
'*             lPort    , The port number (80 = HTTP, 443 = HTTPS)
'*             bSecure  , Use HTTPS/SSL True/False
'*             sPage    , Post the content to this page
'*             sUserID  , The optional user ID
'*             sPassword, The optional password
'*-------------
Public Sub Post(ByVal sHost As String, _
                ByVal lPort As Long, _
                ByVal bSecure As Boolean, _
                ByVal sPage As String, _
                Optional ByVal sUserID As String, _
                Optional ByVal sPassword As String, _
                Optional ByVal bUseProxy As Boolean, _
                Optional ByVal sProxyServer As String, _
                Optional ByVal lProxyPort As Long, _
                Optional ByVal sProxyUserID As String, _
                Optional ByVal sProxyPassword As String)
  Dim lSession As Long, lConnection As Long, lRequest As Long
  Dim lFlags As Long
  Dim iResult As Integer
  Dim sBuffer As String * 256
  Dim lError As Long
  Dim sError As String
  Dim sHeaders As String
  
  On Error GoTo ErrorHandler
  
  '
  ' Intialize a session with WinInet
  '
  '--changed by Marcelo Cabral - 21/09/2002
  '
  If Not bUseProxy Then
     lSession = InternetOpen(App.ProductName, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
  ElseIf sProxyServer <> "" Then
     lSession = InternetOpen("Mozilla/4.0 (compatible; MSIE 5.0; Windows 98)", INTERNET_OPEN_TYPE_PROXY, sProxyServer & ":" & Format(lProxyPort), "<local>", 0)
  Else
     lSession = InternetOpen("Mozilla/4.0 (compatible; MSIE 5.0; Windows 98)", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  End If
  '--
  If lSession = 0 Then
  
    '
    ' Unable to start a session
    '
    Err.Raise Err.LastDLLError, , "InternetOpen failed"
    
  Else
  
    '
    ' Setup the timeout values
    '
    InternetSetOption lSession, INTERNET_OPTION_CONNECT_TIMEOUT, m_lTimeout, Len(m_lTimeout)
    InternetSetOption lSession, INTERNET_OPTION_RECEIVE_TIMEOUT, m_lTimeout, Len(m_lTimeout)
    InternetSetOption lSession, INTERNET_OPTION_SEND_TIMEOUT, m_lTimeout, Len(m_lTimeout)
  
    '
    ' Make a HTTP connection with the specified host
    '
    If sUserID <> "" Then
      lConnection = InternetConnect(lSession, sHost, lPort, sUserID, sPassword, INTERNET_SERVICE_HTTP, 0, 0)
    Else
      lConnection = InternetConnect(lSession, sHost, lPort, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    End If
      
    If lConnection = 0 Then
    
      '
      ' Unable to start a session
      '
      Err.Raise Err.LastDLLError, , "InternetConnect failed"
    
    Else
    
      '
      ' Open a HTTP request
      '
      If bSecure = True Then
        lFlags = INTERNET_FLAG_SECURE
      End If
      
      lRequest = HttpOpenRequest(lConnection, "POST", sPage, "HTTP/1.0", vbNullString, 0, lFlags + INTERNET_FLAG_RELOAD, 0)
    
      If lRequest = 0 Then
      
        '
        ' Unable to open an HTTP request
        '
        lError = Err.LastDLLError
        Err.Raise lError, , "HttpOpenRequest failed"
        
      Else
        '
        ' Proxy Authentication
        '
        '--added by Marcelo Cabral - 23/06/2003
        '
        If bUseProxy And sProxyUserID <> "" Then
            iResult = InternetSetOptionStr(lRequest, INTERNET_OPTION_PROXY_USERNAME, _
                                           sProxyUserID, Len(sProxyUserID) + 1)
            iResult = InternetSetOptionStr(lRequest, INTERNET_OPTION_PROXY_PASSWORD, _
                                           sProxyPassword, Len(sProxyPassword) + 1)
            sHeaders = "Content-Type: " & m_sContentType
'            sHeaders = "Content-Type: " & m_sContentType & vbCrLf
'            Dim Base64Codec As xmlBase64Codec
'            Set Base64Codec = New xmlBase64Codec
'            sHeaders = sHeaders & "Proxy-Authorization: Basic " & Base64Codec.EncodeArray(sProxyUserID & ":" & sProxyPassword) & vbCrLf
'            Set Base64Codec = Nothing
        Else
            sHeaders = "Content-Type: " & m_sContentType
        End If
        '--
        '
        ' Send the HTTP request
        '
        iResult = HttpSendRequest(lRequest, sHeaders, -1, ByVal m_sContent, Len(m_sContent))
      
        If iResult = False Then
                   
          '
          ' Unable to send the HTTP request
          '
          '--changed by Marcelo Cabral - 18/03/2002
          '
          lError = Err.LastDLLError
          Select Case lError
          Case 12002
              Err.Raise lError, , GetMsg(msgErrTimeout)
          Case 12005
              Err.Raise lError, , GetMsg(msgErrInvalidURL)
          Case 12007
              Err.Raise lError, , GetMsg(msgErrServerName)
          Case Else
              sError = TranslateWinsockError(lError)
              If sError <> "" Then
                Err.Raise lError, , sError
              Else
                Err.Raise lError, , "HttpSendRequest failed"
              End If
          End Select
          '--
        Else
        
          '
          ' Retrieve the status code
          '
          iResult = HttpQueryInfo(lRequest, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, Len(sBuffer), 0)
          
          If iResult = False Then
                    
            '
            ' Unable to send the HTTP request
            '
            Err.Raise Err.LastDLLError, , "HttpQueryInfo failed"
        
          Else
          
            sBuffer = ZeroTerminatedToVB(sBuffer)
            iResult = Val(sBuffer)
            
            If iResult <> 200 Then
            
              '
              ' HTTP request failed
              '
              Err.Raise CLng(iResult), , "HTTP POST failed. The error number is the returned HTTP status code."
              
            Else
            
              '
              ' Get the XML result if available
              '
              GetResponse lRequest
            
            End If
          
          End If
          
          InternetCloseHandle lRequest
       
        End If
      
      End If
    
      InternetCloseHandle lConnection
    
    End If
  
    InternetCloseHandle lSession
  
  End If
  
  Exit Sub
  
ErrorHandler:
  
  Dim lNumber As Long
  Dim sDescription As String

  '
  ' An error occured...
  '
  lNumber = Err.Number
  sDescription = Err.Description

  Resume Cleanup

Cleanup:

  '
  ' Cleanup after error...
  '
  On Error Resume Next

  InternetCloseHandle lRequest
  InternetCloseHandle lConnection
  InternetCloseHandle lSession

  On Error GoTo 0

  Err.Raise lNumber, TypeName(Me) & ".Post", sDescription

End Sub

'*-------------
'*  This method will retrieve the complete response from the server.
'*
'*  Input    : lRequest, The HINTERNET handle
'*  Return   : None
'*-------------
Private Sub GetResponse(ByVal lRequest As Long)

  Dim lBytes As Long
  Dim sBuffer As String * 1024
  Dim bEnd As Boolean, bRc As Boolean

  '
  ' Retrieve the complete response from the server
  '
  m_sResponse = ""
  
  Do While Not bEnd
   
    '
    ' Get the next chunk
    '
    sBuffer = vbNullString
    bRc = CBool(InternetReadFile(lRequest, sBuffer, Len(sBuffer), lBytes))
  
    If bRc = False Then
    
      '
      ' Unable to read response
      '
      Err.Raise Err.LastDLLError, , "InternetReadFile failed"
    
    Else
    
      '
      ' Append this chunk to the XML
      '
      m_sResponse = m_sResponse & Left$(sBuffer, lBytes)
      bEnd = (lBytes = 0)
    

    End If
    '--added by Marcelo Cabral - 16/12/2001
    DoEvents
    '--
  Loop

End Sub

'*-------------
'*  This function will convert a zero terminated string
'*  to a normal Visual Basic string.
'*
'*  Input    : sz, Zero terminated string
'*  Modifies : None
'*  Return   : Visual Basic string with correct length
'*-------------
Private Function ZeroTerminatedToVB(ByVal sz As String) As String

  Dim iZero As Integer

  iZero = InStr(sz, Chr$(0))

  If iZero <> 0 Then
    ZeroTerminatedToVB = Left$(sz, iZero - 1)
  Else
    ZeroTerminatedToVB = sz
  End If

End Function

Private Sub Class_Initialize()

  '
  ' Initialize the class
  '
  m_sContentType = "text/xml"

End Sub

