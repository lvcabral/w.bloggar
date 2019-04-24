VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
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
' Class      : clsFTP
' Description: Code for working with FTP servers

Private Const MAX_PATH As Long = 260
Private Const NO_ERROR As Long = 0
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Private Const FILE_ATTRIBUTE_OFFLINE As Long = &H1000

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Const ERROR_NO_MORE_FILES As Long = 18

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
        
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hftpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hftpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hftpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hftpSession As Long, ByVal lpszDirectory As String) As Boolean
    
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hftpSession As Long, ByVal lpszCurrentDirectory As String, lpdword As Long) As Boolean
    
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hftpSession As Long, ByVal lpszDirectory As String) As Boolean
        
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT As Long = 1
Private Const INTERNET_OPEN_TYPE_PROXY As Long = 3
Private Const INTERNET_INVALID_PORT_NUMBER As Long = 0
Private Const FTP_TRANSFER_TYPE_ASCII As Long = &H1
Private Const FTP_TRANSFER_TYPE_BINARY As Long = &H2
Private Const INTERNET_FLAG_PASSIVE As Long = &H8000000

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Long, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
                                
Private Const INTERNET_ERROR_BASE = 12000
Private Const ERROR_INTERNET_EXTENDED_ERROR As Long = (INTERNET_ERROR_BASE + 3)

' FTP API Errors
Private Const ERROR_FTP_TRANSFER_IN_PROGRESS As Long = (INTERNET_ERROR_BASE + 110)
Private Const ERROR_FTP_DROPPED  As Long = (INTERNET_ERROR_BASE + 111)
Private Const ERROR_FTP_NO_PASSIVE_MODE  As Long = (INTERNET_ERROR_BASE + 112)

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Private Const INTERNET_DEFAULT_FTP_PORT As Long = 21
Private Const INTERNET_DEFAULT_GOPHER_PORT As Long = 70
Private Const INTERNET_DEFAULT_HTTP_PORT As Long = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT As Long = 443
Private Const INTERNET_DEFAULT_SOCKS_PORT As Long = 1080

Private Const INTERNET_OPTION_CONNECT_TIMEOUT As Long = 2
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT As Long = 6
Private Const INTERNET_OPTION_SEND_TIMEOUT As Long = 5

Private Const INTERNET_OPTION_USERNAME As Long = 28
Private Const INTERNET_OPTION_PASSWORD As Long = 29
Private Const INTERNET_OPTION_PROXY_USERNAME As Long = 43
Private Const INTERNET_OPTION_PROXY_PASSWORD As Long = 44

Private Const INTERNET_SERVICE_FTP As Long = 1
Private Const INTERNET_SERVICE_GOPHER As Long = 2
Private Const INTERNET_SERVICE_HTTP As Long = 3

Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION As Long = &H400000
Private Const INTERNET_FLAG_MULTIPART As Long = &H200000
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000

Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hftpSession As Long, ByVal sFileName As String, ByVal lAccess As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
        
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hftpSession As Long, ByVal lpszFileName As String) As Boolean

Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hftpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
    
' Private variables to support class properties
Private m_blnUseProxy As Boolean
Private m_strProxyString As String
Private m_lngSessionHandle As Long
Private m_lngConnectionHandle As Long
Private m_fSessionActive As Boolean
Private m_strUserName As String
Private m_strPassword As String
Private m_strServer As String
Private m_lngPort As Long
Private m_lngTransferType  As Long
Private m_strLastDLLError As String
Private m_strLastDLLErrorMsg As String

Public Enum EnumFTPTransferType
    eftASCII = 1
    eftBinary = 2
End Enum

' Collections to hold directory information
Public mcolDirEntries As New Collection
'Private mcolDirEntries As New Collection
Private mcolDirAttributes As New Collection
Private mcolDirSizes As New Collection

Private Sub Class_Initialize()
    ' Comments  : Code run when the class is initialized
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    
    ' Set reasonable values
    m_fSessionActive = False
    m_lngSessionHandle = 0
    m_lngConnectionHandle = 0
        
    m_lngTransferType = eftBinary
    
    ' Assume the FTP site allows anonymous logins
    m_strUserName = "anonymous"
    m_strPassword = ""
    m_lngPort = INTERNET_DEFAULT_FTP_PORT
    
ExitNow:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Sub

Public Property Get ConnectionHandle() As Long
    ' Returns: The handle to the Internet connection
    '
    '
    ConnectionHandle = m_lngConnectionHandle
    
End Property

Public Property Get FTPCurDir() As String
    ' Returns: The current directory as set on the FTP server
    '
    '
    Dim strBuf As String * 512
    Dim fRetval As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Clear the error slot
    ClearDLLError
    
    ' Get the current directory
    fRetval = FtpGetCurrentDirectory(m_lngConnectionHandle, strBuf, Len(strBuf))
    
    ' Save the error
    GetDLLError
    
    If fRetval Then
        ' Massage the return value
        FTPCurDir = TrimNulls(strBuf)
    Else
        FTPCurDir = ""
    End If
    
ExitNow:
    Exit Property
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Property

Public Property Get FTPDirAttrNumber(intIndex As Integer) As Long
    ' Returns   : The file attributes of the specified item in the
    '             current directory as a long integer
    ' Parameters: intIndex - index position
    '
    '
    FTPDirAttrNumber = mcolDirAttributes(intIndex)
    
ExitNow:
    Exit Property
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Property

Public Property Get FTPDirAttrString(intIndex As Integer) As String
    ' Returns   : The file attributes of the specified item in the
    '             current directory as a readable string
    ' Parameters: intIndex - index position
    '
    '
    FTPDirAttrString = GetEnglishAttributes(mcolDirAttributes(intIndex))
    
ExitNow:
    Exit Property
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Property

Public Property Get FTPDirCount() As Integer
    ' Returns: The count of entries in the current directory
    '
    '
    FTPDirCount = mcolDirEntries.Count
    
ExitNow:
    Exit Property
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
            
End Property

Public Property Get FTPDirEntry(intIndex As Integer) As String
    ' Returns   : The name of the specified file or sub-directory in the
    '             current directory
    ' Parameters: intIndex - index position
    '
    '
    FTPDirEntry = mcolDirEntries(intIndex)
    
ExitNow:
    Exit Property
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Property

Public Property Get FTPDirSize(intIndex As Integer) As Long
    ' Returns   : The file size of the specified item in the
    '             current directory as a long integer
    ' Parameters: intIndex - index position
    '
    '
    FTPDirSize = mcolDirSizes(intIndex)
    
ExitNow:
    Exit Property
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Property

Public Property Get LastDLLError() As String
    ' Returns: The string of the last error raised by a call to a DLL function
    '
    '
    LastDLLError = m_strLastDLLError
    
End Property

Public Property Get LastDLLErrorMsg() As String
    ' Returns: The string of the last error raised by a call to a DLL function
    '
    '
    LastDLLErrorMsg = m_strLastDLLErrorMsg
    
End Property

Public Property Get Password() As String
    ' Returns: The current login password
    '
    '
    Password = m_strPassword
    
End Property

Public Property Let Password(strValue As String)
    ' strValue: The setting for the login password. Leave this property value
    '           blank for anonymous FTP logins.
    '
    '
    m_strPassword = strValue
    
End Property

Public Property Get UseProxy() As Boolean
    ' Returns: True if uses a proxy server
    '
    '
    UseProxy = m_blnUseProxy
    
End Property

Public Property Let UseProxy(blnValue As Boolean)
    ' blnValue: True if uses a proxy server
    '
    '
    m_blnUseProxy = blnValue
    
End Property

Public Property Get ProxyString() As String
    ' Returns: The current proxy server string
    '
    '
    ProxyString = m_strProxyString
    
End Property

Public Property Let ProxyString(strValue As String)
    ' strValue: The proxy server setting to use
    '
    '
    m_strProxyString = strValue
    If strValue <> "" Then m_blnUseProxy = True
End Property

Public Property Get Server() As String
    ' Returns: The URL of the current FTP servers
    '
    '
    Server = m_strServer
    
End Property

Public Property Let Server(strValue As String)
    ' strValue: Name of the FTP server to connect to in the form:
    '           ftp.domain.suffix (e.g. ftp.microsoft.com)
    '
    '
    m_strServer = strValue
    
End Property

Public Property Get Port() As Long
    Port = m_lngPort
End Property

Public Property Let Port(ByVal intPort As Long)
    m_lngPort = intPort
End Property

Public Property Get SessionActive() As Boolean
    ' Returns: True if there is an active session open to an FTP site
    '
    '
    SessionActive = m_fSessionActive
    
End Property

Public Property Get SessionHandle() As Long
    ' Returns: The handle to the Internet session
    '
    '
    SessionHandle = m_lngSessionHandle
    
End Property

Public Property Get TransferType() As EnumFTPTransferType
    ' Returns: The current TransferType setting
    '
    '
    TransferType = m_lngTransferType

End Property

Public Property Let TransferType(eValue As EnumFTPTransferType)
    ' eValue: Defines the FTP transfer type. Use one of the EnumFTPTransferType
    '         values.
    
    '
    m_lngTransferType = eValue
    
End Property

Public Property Get UserName() As String
    ' Returns: The current user name setting
    '
    '
    UserName = m_strUserName
    
End Property

Public Property Let UserName(strValue As String)
    ' strValue: The setting for the user name. By default, this class
    '           sets this value to 'anonymous' to support anonymous ftp logins
    '
    '
    m_strUserName = strValue
    
End Property

Public Sub CloseSession()
    ' Comments  : Closes the current handles to the Internet
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    On Error GoTo ErrorHandler

    ' Clear the error slot
    ClearDLLError
    
    If m_lngConnectionHandle <> 0 Then
        ' Close the connection handle
        InternetCloseHandle (m_lngConnectionHandle)
    End If
    
    If m_lngSessionHandle <> 0 Then
        ' Close the sessions handle
        InternetCloseHandle (m_lngSessionHandle)
    End If
    
    ' Save the error
    GetDLLError
    
    ' Set class property variables
    m_lngConnectionHandle = 0
    m_lngSessionHandle = 0
    
    If m_fSessionActive Then
        m_fSessionActive = False
    End If
    
ExitNow:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Sub

Public Function Connect() As Boolean
    ' Comments  : Connects to the named FTP server
    ' Parameters: None
    ' Returns   : True if successful, False otherwise
    '
    '
    Dim fConnected As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Assume failure
    fConnected = False
    
    ' Ensure we aren't already in an active session
    If Not m_fSessionActive Then
    
        ' Ensure we have a handle open
        If m_lngSessionHandle <> 0 Then
        
            ' Ensure we have a server name and user name
            If m_strServer <> "" And m_strUserName <> "" Then
            
                ' Clear the error slot
                ClearDLLError
                
                m_lngConnectionHandle = InternetConnect(m_lngSessionHandle, m_strServer, _
                    m_lngPort, m_strUserName, m_strPassword, _
                    INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
                    
                ' Save the error
                GetDLLError
                
                ' Check the return value
                If m_lngConnectionHandle <> 0 Then
                    ' We're active, so set various things
                    fConnected = True
                    m_fSessionActive = True
                                
                    ' Fill the collection
                    'Call FTPEnumerateDir
                                            
                Else
                    fConnected = False
                End If
            End If
        End If
    End If
            
ExitNow:
    Connect = fConnected
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Function

Public Function CreateDirectory(strDirectory As String) As Boolean
    ' Comments  : Creates the specified directory on the current FTP site
    ' Parameters: strDirectory - fully qualified directory to create
    ' Returns   : True if successful, False otherwise
    '
    '
    Dim fRetval As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Assume failure
    fRetval = True
    
    ' Clear the error slot
    ClearDLLError
    
    If strDirectory <> "" Then
        ' Make the call the create the directory
        fRetval = FtpCreateDirectory(m_lngConnectionHandle, strDirectory)
    End If
    
    ' Save the Error
    GetDLLError
    
ExitNow:
    CreateDirectory = fRetval
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub Disconnect()
    ' Comments  : Disconnects the link to the server
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    On Error GoTo ErrorHandler
    
    ' Only do this if there actually is a connetion
    If m_lngConnectionHandle <> 0 Then
    
        ' Clear the error slot
        ClearDLLError
        
        ' Make the call to close the connection handle
        InternetCloseHandle m_lngConnectionHandle
        
        ' Save the error
        GetDLLError
        
        ' Reset the class handle
        m_lngConnectionHandle = 0
        
        ' Set the active flag
        m_fSessionActive = False
        
    End If
    
ExitNow:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Sub

Public Function GetFileNames(strRemoteFile As String) As Collection
    ' Comments  : Retreives the name of the files from the remote server
    ' Parameters: strRemoteFile - Path and name of the file on the FTP server
    ' Returns   : True if successful, False otherwise
    '

    Dim strFileOnRemote As String
    Dim strDirOnRemote As String

    On Error GoTo ErrorHandler
    
    ' Ensure we are in a session
    If m_fSessionActive Then
        
        ' Check for a passed file names
        If strRemoteFile <> "" Then
            
            ' Get the path/file into temporary strings
            strDirOnRemote = strRemoteFile
            'strDirOnRemote = GetPathFromFTPPath(strRemoteFile)
                        
            ' Change the directory on the remote server
            'colFiles = RemoteChDirFiles(strDirOnRemote)
            RemoteChDir strDirOnRemote
            
            ' Clear the error slot
            ClearDLLError
            
            ' Save the error
            GetDLLError
            
        End If
        
    End If
    
ExitNow:
    'GetFiles = fRetval
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
        
End Function

Public Function GetFile( _
    strRemoteFile As String, _
    strLocalFile As String) _
    As Boolean
    ' Comments  : Retreives the specified file from the remote server
    ' Parameters: strRemoteFile - Path and name of the file on the FTP server
    '             strLocalFile - Path and name of the file to write to on
    '             the client.
    ' Returns   : True if successful, False otherwise
    '
    '
    Dim fRetval As Boolean
    Dim strFileOnRemote As String
    Dim strDirOnRemote As String

    On Error GoTo ErrorHandler
    
    ' Ensure we are in a session
    If m_fSessionActive Then
        
        ' Check for a passed file names
        If strRemoteFile <> "" And strLocalFile <> "" Then
            
            ' Get the path/file into temporary strings
            strDirOnRemote = GetPathFromFTPPath(strRemoteFile)
            strFileOnRemote = GetNameFromFTPPath(strRemoteFile)
                        
            ' Change the directory on the remote server
            RemoteChDir strDirOnRemote
            
            ' Clear the error slot
            ClearDLLError
            
            ' Retreive the file.
            fRetval = FtpGetFile(m_lngConnectionHandle, strFileOnRemote, _
                strLocalFile, False, FILE_ATTRIBUTE_NORMAL, _
                INTERNET_FLAG_RELOAD Or m_lngTransferType, 0)
            
            ' Save the error
            GetDLLError
            
        End If
        
    End If
    
ExitNow:
    GetFile = fRetval
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
        
End Function

Public Function RenameFile( _
    strCurrentName As String, _
    strNewName As String) _
    As Boolean
    ' Comments  : Retreives the specified file from the remote server
    ' Parameters: strCurrentName - Path and name of the file on the FTP server
    '             strNewName - Path and name of the file to write to on
    '             the client.
    ' Returns   : True if successful, False otherwise
    '
    '
    Dim fRetval As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Ensure we are in a session
    If m_fSessionActive Then
        
        ' Check for a passed file names
        If strCurrentName <> "" And strNewName <> "" Then
            
            ' Clear the error slot
            ClearDLLError
            
            ' Retreive the file.
            fRetval = FtpRenameFile(m_lngConnectionHandle, strCurrentName, _
                strNewName)
            
            ' Save the error
            GetDLLError
            
        End If
        
    End If
    
ExitNow:
    RenameFile = fRetval
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
        
End Function

Private Function GetPathFromFTPPath(strFTPPath As String) As String
    ' Comments  : Returns the path part of a fully qualified FTP file name
    ' Parameters: strFTPPath - path to parse
    ' Returns   : Path without the file name
    '
    '
    Dim intCounter As Integer
    
    On Error GoTo ErrorHandler
        
    For intCounter = Len(strFTPPath) To 1 Step -1
        If Mid$(strFTPPath, intCounter, 1) = "/" Then
            Exit For
        End If
    Next intCounter

    GetPathFromFTPPath = Left$(strFTPPath, intCounter)
    
ExitNow:
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Function

Public Function OpenSession() As Boolean
    ' Comments  : Opens a session to the Interenet resource
    ' Parameters: None
    ' Returns   : True if the session opened succesfully, False otherwise
    '
    '
    Dim fOpened As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Assume failure
    fOpened = False
    
    ' Clear the error slot.
    ClearDLLError
    
    ' If a proxy string is specified, we have to make a slightly
    ' different call.
    If Not m_blnUseProxy Then
        m_lngSessionHandle = InternetOpen(App.ProductName, INTERNET_OPEN_TYPE_DIRECT, _
        vbNullString, vbNullString, 0)
    ElseIf m_strProxyString = "" Then
        m_lngSessionHandle = InternetOpen(App.ProductName, INTERNET_OPEN_TYPE_PRECONFIG, _
        vbNullString, vbNullString, 0)
    Else
        m_lngSessionHandle = InternetOpen(App.ProductName, INTERNET_OPEN_TYPE_PROXY, _
        m_strProxyString, vbNullString, 0)
    End If
    
    ' Save the error
    GetDLLError
    
    ' Get the return value
    fOpened = (m_lngSessionHandle <> 0)
    
ExitNow:
    OpenSession = fOpened
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Function

Public Function PutFile(ByVal strLocalFile As String, ByVal strRemoteFile As String) As Boolean
    ' Comments  : Sends the named local file to the remote server
    ' Parameters: strLocalFile - Path and name of the local file
    '             strRemoteFile - Path and name of the local file
    ' Returns   : True if successful, False otherwise
    '
    '
    Dim fRetval As Boolean
    Dim strFileOnRemote As String
    Dim strDirOnRemote As String
    Dim strFileOnLocal As String
    Dim strTmp As String
    
    On Error GoTo ErrorHandler
    
    ' Ensure we are in a session
    If m_fSessionActive Then
        
        ' Check passed file names
        If strLocalFile <> "" And strRemoteFile <> "" Then
        
            ' Parse file names
'            strTmp = strRemoteFile
'            strDirOnRemote = Right(strTmp, Len(strTmp) - Len(m_strServer))
'            strFileOnRemote = strRemoteFile
            strDirOnRemote = GetPathFromFTPPath(strRemoteFile)
            strFileOnRemote = GetNameFromFTPPath(strRemoteFile)
            strFileOnLocal = strLocalFile
            
            ' Check for root directory
            If (strDirOnRemote = "") Then
                strDirOnRemote = "/"
            End If
            
            ' Change directories on the server
            RemoteChDir strDirOnRemote
    
            ' Clear the error slot
            ClearDLLError
            
            ' Send the file
            fRetval = FtpPutFile(m_lngConnectionHandle, strFileOnLocal, _
                strFileOnRemote, m_lngTransferType, 0)
                
            ' Save error status
            GetDLLError
            
        End If
        
    End If
    
ExitNow:
    PutFile = fRetval
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Function

Public Function RemoteChDir(strRemoteDir As String) As Boolean
    ' Comments  : Changes the directory on the remote server
    ' Parameters: strRemoteDir - directory to change to
    ' Returns   : True if successful, False otherwise
    '
    '
    Dim fRetval As Boolean
    Dim strPathFromRoot As String
    
    On Error GoTo ErrorHandler
    
    ' Assume failure
    fRetval = False
    
    ' Remove the server URL
    If InStr(1, strRemoteDir, m_strServer) Then
        strPathFromRoot = Mid(strRemoteDir, Len(m_strServer) + 1, _
            Len(strRemoteDir) - Len(m_strServer))
    Else
        strPathFromRoot = strRemoteDir
    End If
    
    ' If no directory specified, add a slash
    If strPathFromRoot = "" Then
        strPathFromRoot = "/"
    End If
    
    ' Clear the DLL error slot
    ClearDLLError
    
    ' Tell the server to set a new current directory
    fRetval = FtpSetCurrentDirectory(m_lngConnectionHandle, strPathFromRoot)
    
    ' Catch any errors
    GetDLLError
        
    If fRetval Then

        ' Directory was changed, refresh our directory collection
        FTPEnumerateDir
    End If
    
ExitNow:
    RemoteChDir = fRetval
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Function

Private Sub ClearCollections()
    ' Comments  : Clears the directory collections
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    Dim intCounter As Integer
    
    On Error GoTo ErrorHandler
    
    For intCounter = 1 To mcolDirEntries.Count
        mcolDirEntries.Remove 1
    Next intCounter
        
    For intCounter = 1 To mcolDirAttributes.Count
        mcolDirAttributes.Remove 1
    Next intCounter
    
    For intCounter = 1 To mcolDirSizes.Count
        mcolDirSizes.Remove 1
    Next intCounter
    
ExitNow:
    Exit Sub
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Sub

Private Sub ClearDLLError()
    ' Comments  : Clears the class's last DLL error
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    m_strLastDLLError = ""
    
End Sub

Private Sub FTPEnumerateDir()
    ' Comments  : Fills the class collections with the contents of the
    '             current directory.
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    Dim lngFindSession As Long
    Dim pData As WIN32_FIND_DATA
    Dim fRetval As Boolean
    Dim strItemName As String
    
    On Error GoTo ErrorHandler
    
    ' Clear out the current collection
    ClearCollections

    ' Set up the structure and get the first file
    pData.cFileName = String(MAX_PATH, 0)
    
    ClearDLLError
    
    lngFindSession = FtpFindFirstFile(m_lngConnectionHandle, "*.*", pData, 0, 0)
    
    GetDLLError
    
    ' If the call returned 0, the directory is empty
    If lngFindSession <> 0 Then
        
        ' Add the file attributes
        mcolDirAttributes.Add pData.dwFileAttributes
        
        ' Add the file size
        If Not IsNull(pData.nFileSizeHigh) Then
            mcolDirSizes.Add pData.nFileSizeHigh
        Else
            mcolDirSizes.Add "0"
        End If
        
        ' Add the file name
        strItemName = Left(pData.cFileName, _
            InStr(1, pData.cFileName, String(1, 0), _
            vbBinaryCompare) - 1)
        mcolDirEntries.Add strItemName
        
        ' Loop to find additional files
        Do
            pData.cFileName = String(MAX_PATH, 0)
            
            ClearDLLError
            
            fRetval = InternetFindNextFile(lngFindSession, pData)
            
            GetDLLError
            
            If fRetval Then
                ' Additional item found.
                mcolDirAttributes.Add pData.dwFileAttributes
                strItemName = Left(pData.cFileName, _
                    InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
                mcolDirEntries.Add strItemName
                
                ' Add the file size
                If Not IsNull(pData.nFileSizeHigh) Then
                    mcolDirSizes.Add pData.nFileSizeLow
                Else
                    mcolDirSizes.Add "0"
                End If
            
            Else
                ' Which error happened?
                If Err.LastDLLError = ERROR_NO_MORE_FILES Then
                    Exit Do
                Else
                    ' some other error happened
                End If
            End If
            
        Loop
        
        ' Close the handle to the Find session
        ClearDLLError
        
        InternetCloseHandle (lngFindSession)
        
        GetDLLError
        
    End If
    
ExitNow:
    Exit Sub
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
        
End Sub

Private Sub GetDLLError()
    ' Comments  : Set the class Last DLL error properties
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    On Error GoTo ErrorHandler
    
    m_strLastDLLError = Err.LastDLLError
    GetDLLErrorMsg
ExitNow:
    Exit Sub
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Sub

Private Sub GetDLLErrorMsg()
    ' Comments  : Get the error message
    ' Parameters: None
    ' Returns   : Nothing
    '
    '
    On Error GoTo PROC_ERR
    Dim blnRes As Boolean
    Dim lngErro As Long
    Dim sBuffer As String * 1024
    lngErro = Val(m_strLastDLLError)
    If lngErro = ERROR_INTERNET_EXTENDED_ERROR Then
        sBuffer = vbNullString
        blnRes = CBool(InternetGetLastResponseInfo(lngErro, sBuffer, Len(sBuffer)))
        m_strLastDLLErrorMsg = sBuffer
    Else
        m_strLastDLLErrorMsg = TranslateWinsockError(lngErro)
    End If
    
PROC_EXIT:
    Exit Sub
        
PROC_ERR:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume PROC_EXIT
    
End Sub

Private Function GetEnglishAttributes(lngAttr As Long) As String
    ' Comments  : Returns a human readable version of an FTP directory entry
    ' Parameters: lngAttr - long integer attribute as returned by the
    '             call to FtpFindFirstFile and FtpFinNextFile.
    ' Returns   : String
    '
    '
    Dim strTmp As String
    
    On Error GoTo ErrorHandler
    
    If (lngAttr And FILE_ATTRIBUTE_READONLY) > 0 Then
        strTmp = "Readonly "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_HIDDEN) > 0 Then
        strTmp = strTmp & "Hidden "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_SYSTEM) > 0 Then
        strTmp = strTmp & "System "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_DIRECTORY) > 0 Then
        strTmp = strTmp & "<directory> "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_TEMPORARY) > 0 Then
        strTmp = strTmp & "Temporary "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_ARCHIVE) > 0 Then
        strTmp = strTmp & "Archive "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_NORMAL) > 0 Then
        strTmp = strTmp & "normal "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_COMPRESSED) > 0 Then
        strTmp = strTmp & "Compressed "
    End If
    
    If (lngAttr And FILE_ATTRIBUTE_OFFLINE) > 0 Then
        strTmp = strTmp & "Offline "
    End If
    
    GetEnglishAttributes = Trim(strTmp)
    
ExitNow:
    Exit Function
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
    
End Function

Private Function GetNameFromFTPPath(strFTPPath As String) As String
    ' Comments  : Returns the name part of a fully qualified FTP file name
    ' Parameters: strFTPPath - path to parse
    ' Returns   : File name and extension
    '
    Dim intCounter As Integer
    Dim strTmp As String
    Dim chrTmp As String * 1

    On Error GoTo ErrorHandler
    
    ' Parse the string
    For intCounter = Len(strFTPPath) To 1 Step -1
        ' It its a slash, grab the sub string
        chrTmp = Mid$(strFTPPath, intCounter, 1)
        If chrTmp <> "/" Then
            strTmp = chrTmp & strTmp
        Else
            Exit For
        End If
    Next intCounter

    ' Return the value
    GetNameFromFTPPath = strTmp
    
ExitNow:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow

End Function

Private Function TrimNulls(ByVal strIn As String) As String
    ' Comments  : Returns the passed string terminated
    '             at the first null character
    ' Parameters: strIn - Value to parse
    ' Returns   : Parsed string
    '
    '
    Dim intPos As Integer
    
    On Error GoTo ErrorHandler
        
    intPos = InStr(strIn, vbNullChar)
    
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
            ' to the occurrence of the null character
            TrimNulls = Left$(strIn, intPos - 1)
        End If
    End If
        
ExitNow:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume ExitNow
        
End Function

