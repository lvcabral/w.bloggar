Attribute VB_Name = "basBloggar"
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
'Constants
Public Const BODYTAG = "<body bgcolor=""white"" link=""#336699"" vlink=""#003366"" alink=""red"">"
Public Const CSSTAG = "<link rel=""stylesheet"" href=""http://www.[yoursite].com/[yourfile].css"" type=""text/css"" />"
Public Const TITLESTYLE = "style=""font-family: Verdana, Arial, sans-serif;color: black;font-size:10px;font-weight: bold"""
Public Const POSTSTYLE = "style=""font-family: Verdana, Arial, sans-serif;color: black;font-size:10px;"""
Public Const POSTALIGN = "align=""left"""
Public Const POSTWIDTH = "width=""100%"""
Public Const MEDIASTR = "<div class=""media"">[%1%: %T% - %A% - %B% (%D%)]</div>"

Public Const TAB_EDITOR = 0
Public Const TAB_MORE = 1
Public Const TAB_PREVIEW = 2
Public Const MAX_UNDO = 100

Public Const XML_SETTINGS = "\settings.xml"
Public Const REGISTRY_KEY = "Bloggar"

'Public Objects
Public objXMLReg As XMLRegistry

'Data Structures
Public Type Settings
    Tray As Boolean
    ClearPost As Boolean
    PostFile As String
    PostTemplate As String
    FontFace As String
    FontSize As Integer
    TabSpaces As Integer
    ColorizeCode As Boolean
    StartMinimized As Boolean
    ShowHtmlBar As Boolean
    AutoConvert As Boolean
    SkinFolder As String
    Silent As Boolean
    DefaultBrowser As Boolean
    AppLCID As Long
    SpellLCID As Long
    CustomTag(1 To 12, 1 To 3) As String
    '4.0
    OpenLastFile As Boolean
    XHTML As Boolean
    BlogListSize As Integer
    PostMenu As String
End Type

Public Type Account
    Current As Integer
    Alias As String
    CMS As Integer
    Service As Boolean
    User As String
    Password As String
    SavePassword As Boolean
    PingWeblogs As Integer
    Host As String
    Page As String
    Port As Long
    Secure As Boolean
    Timeout As Long
    UseProxy As Integer
    ProxyServer As String
    ProxyPort As Long
    ProxyUser As String
    ProxyPassword As String
    GetCategMethod  As SupportedAPI
    PostMethod  As SupportedAPI
    GetPostsMethod  As SupportedAPI
    TemplateMethod  As SupportedAPI
    MultiCategory As Boolean
    TitleTag1 As String
    TitleTag2 As String
    CategTag1 As String
    CategTag2 As String
    '3.03
    BodyTag1 As String
    BodyTag2 As String
    MoreTextTag1 As String
    MoreTextTag2 As String
    '4.01
    UTF8 As Boolean
    '4.02
    UTF8OnPost As Boolean
    MoreTab As Boolean
    AdvancedOptions As Boolean
    AllowComments As Boolean
    AllowPings As Boolean
    TextFilters As Boolean
    PostDate As Boolean
    TrackBack As Boolean
    Extended As Boolean
    Excerpt As Boolean
    Keywords As Boolean
    UploadMethod As SupportedAPI
End Type

Public Type BlogSettings
    PreviewBody As String
    PreviewCSS As String
    PreviewTitle As String
    PreviewStyle As String
    PreviewAlign As String
    PreviewWidth As String
    PreviewAutoBR As Boolean
    APIUpload As Boolean
    FTPHost As String
    FTPPath As String
    FTPPort As Integer
    FTPProxy As Boolean
    FTPUser As String
    FTPPassword As String
    FTPLink As String
    MediaInsert As Integer
    MediaLink As Boolean
    MediaString As String
End Type

Public Type Blogs
    URL As String
    BlogID As String
    Name As String
    IsAdmin As Boolean
End Type

'Enumerators
Public Enum MessagesEnum
    msgGettingBlogs
    msgGettingPosts
    msgGettingPost
    msgGettingTemplate
    msgPosting
    msgPPublishing
    msgPostSuccess
    msgPostError
    msgUpdateSuccess
    msgWantDelete
    msgDeleting
    msgDPublishing
    msgDelSuccess
    msgDelError
    msgEnterPostID
    msgInvalidPostID
    msgMainTemplate
    msgArchiveTemplate
    msgNewPost
    msgFileFilter
    msgQueryUnload
    msgEnterUser
    msgEnterPassword
    msgEnterHost
    msgEnterPage
    msgEnterPort
    msgInvalidFontSize
    msgNothingToPost
    msgEnterProxy
    msgEnterURL
    msgEnterImage
    msgEnterRowCol
    msgInvalidFontFace
    msgNewAccount
    msgNotYetImpl
    msgNoMatches
    msgNoMatchesSel
    msgNoTextSearch
    msgReplMade
    msgPostingTo
    msgErrorPosting
    msgChooseAction
    msgOnlyNewPost
    msgBlogEmpty
    msgInvalidBody
    msgEnterTimeout
    msgSysInfoError
    msgValidatingUser
    msgPingingWeblogs
    msgErrTimeout
    msgErrInvalidURL
    msgErrServerName
    msgHowManyPosts
    msgMinIEVersion
    msgTplSavedNotPub
    msgWaitDicLoad
    msgEnterAlias
    msgEnterTabSpaces
    msgDefPostNotFound
    msgSelectDraft
    msgItemExists
    msgSpellComplete
    msgFileNotFound
    msgStartUpload
    msgOpenSession
    msgNotOpenSession
    msgConnecting
    msgNotConnect
    msgSendingFile
    msgNotSent
    msgDisconnecting
    msgEnterMenuCaption
    msgEnterCustomTag
    msgFTPSettings
    msgEditChanged
    msgDelAccount
    msgDelPosts
    msgGettingBlogCategs
    msgPublishingBlog
    msgDraftPostNotSaved
    msgFileFilterAll
    msgFileFilterText
    msgWizMsg00
    msgWizMsg01
    msgWizMsg02
    msgWizMsg03
    msgWizMsg04
    msgWizMsg05
    msgWizMsg06
    msgWizMsg07
    msgWizMsg08
    msgWizMsg09
    msgWizMsg10
    msgWizMsg11
    msgWizMsg12
    msgWizMsg13
    msgInvalidSettings
    msgSettingsImported
    msgRestartLanguage
    msgInvalidBlogListSize
    msgErrOpenSettings
    msgConvSettingsXML
    msgKeepSettings
    msgLoadingDict
    msgPostFileWithID
    msgLoadAsDraft
    msgXMLFileFilter
End Enum

Public Enum LabelsEnum
    lblOK
    lblCancel
    lblSettings
    lblYourAccount
    lblUser
    lblPassword
    lblReloadBlogs
    lblBloggerInfo
    lblHost
    lblPage
    lblPort
    lblOptions
    lblMinimizeTray
    lblClearAfter
    lblFontSize
    lblRecentPosts
    lblSelect
    lblInsertLink
    lblURL
    lblTitle
    lblClass
    lblTarget
    lblInsertImage
    lblImage
    lblTooltip
    lblAlign
    lblBorder
    lblSobre
    lblVersion
    lblDescription
    lblMadeInBrazil
    lblSysInfo
    lblNewWindow
    lblSameFrame
    lblSameWindow
    lblWidth
    lblHeight
    lblDontChange
    lblConnection
    lblSave
    lblProxy
    lblAddress
    lblColorize
    lblFontFace
    lblPostedBy
    lblAccount
    lblEditor
    lblPreview
    lblPost
    lblPostedAt
    lblSavePassword
    lblTranslated
    lblTranslatorName
    lblTranslatorURL
    lblDeleted_01
    lblFind
    lblReplace
    lblReplaceAll
    lblFindNext
    lblClose
    lblScope
    lblAllText
    lblSelected
    lblMatchCase
    lblWholeWord
    lblFormatFont
    lblColor
    lblSize
    lblPostMany
    lblPublish
    lblCellPadding
    lblCellSpacing
    lblRows
    lblColumns
    lblInsertTable
    lblGeneral
    lblBlog
    lblStartMin
    lblShowHTML
    lblCodeEditor
    lblPing
    lblDeleted_02
    lblPreviewFmt
    lblBodyTag
    lblPostStyle
    lblPostAlign
    lblRestoreDef
    lblTipBlog
    lblLeft
    lblRight
    lblMiddle
    lblAbsMiddle
    lblTop
    lblBottom
    lblCenter
    lblTextHere
    lblAutoConvert
    lblTimeout
    lblSeconds
    lblCustom
    lblNoProxy
    lblIEProxy
    lblMyProxy
    lblBlogTool
    lblAccountAlias
    lblPosts
    lblCategories
    lblTemplates
    lblTitleTags
    lblCategTags
    lblCMS
    lblNotSupported
    lblNotInDic
    lblChangeTo
    lblSuggestedWords
    lblChange
    lblIgnore
    lblIgnoreAll
    lblAdd
    lblTagHotkey
    lblMenuCaption
    lblTagOpen
    lblTagClose
    lblText
    lblDeleteAccount
    lblNew
    lblClickToEdit
    lblDelete
    lblPostFiles
    lblUpload
    lblFileExtension
    lblAssociate
    lblDefaultPost
    lblLoadDefault
    lblDraftPost
    lblLoadDraft
    lblSilentPost
    lblToolbarSkin
    lblDictionary
    lblTabSpaces
    lblFTPServer
    lblRemotePath
    lblLinkURL
    lblTipUpload
    lblUseAccountProxy
    lblTitleStyle
    lblPostWidth
    lblFile
    lblAfterUpload
    lblImageOnPost
    lblLinkToFile
    lblOpenSite
    lblDicNotFound
    lblNoSuggestions
    lblDefaultBrowser
    lblCSSTag
    lblConvertBR
    lblAddMediaInfo
    lblMediaOptions
    lblMediaManual
    lblMediaAutoTop
    lblMediaAutoBottom
    lblMediaLink
    lblMediaString
    lblMediaTitle
    lblMediaArtist
    lblMediaAlbum
    lblMediaDuration
    lblMediaCodes
    lblTipMedia
    lblMedia
    lblNoMedia
    lblListening
    lblCategory
    lblMoreText
    lblBackToMain
    lblMore
    lblExtendedEntry
    lblExcerpt
    lblKeywords
    lblAdvanced
    lblBlogListSize
    lblPixels
    lblUseXHTML
    lblReopenLastFile
    lblLanguage
    lblMoreTextTags
    lblEditTag
    lblBack
    lblNext
    lblFinish
    lblAddAccWizard
    lblDoYouHaveBlog
    lblWelcomeTo
    lblCustBlogSettings
    lblAccConnSettings
    lblAccProxyServer
    lblAccUsrPwd
    lblSubscribeProvider
    lblYesHaveBlog
    lblNoHaveBlog
    lblUseProxy
    lblAllowComments
    lblAllowPings
    lblTextFilters
    lblDateTime
    lblSendTrackbackTo
    lblUseCurrentDT
    lblAdvPostOpt
    lblNone
    lblOpen
    lblClosed
    lblYes
    lblNo
    lblDefault
    lblReload
    lblName
    lblFeatures
    lblSubscribe
    lblFree
    lblTrial
End Enum

'Public Variables
Public gAccount As Account
Public gSettings As Settings
Public gBlog As BlogSettings
Public gBlogs() As Blogs
Public gPostID As String
Public gIsXP As Boolean
Public gLCID As Long
Public objColor As ColorDialog
Public gAppDataPath As String

'Windows API Functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusWindow Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

'Windows API Types and Constants
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type DllVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type

Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const SW_SHOW = 5
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
Private Const WM_USER = &H400
Private Const WM_SETTEXT = &HC
Private Const EM_HIDESELECTION = WM_USER + 63
Private Const GW_HWNDPREV = 3
Private Const INFINITE = &HFFFFFFFF
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const WAIT_TIMEOUT = &H102&

'used for platform id
Private Const VER_PLATFORM_WIN32s = 0 'win 3.x
Private Const VER_PLATFORM_WIN32_WINDOWS = 1 'win 9.x
Private Const VER_PLATFORM_WIN32_NT = 2 'win nt,2000,XP
'used for product type
Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_SERVER = 3
'used for suite mask
Private Const VER_SUITE_DATACENTER = 128
Private Const VER_SUITE_ENTERPRISE = 2
Private Const VER_SUITE_PERSONAL = 512

Sub Main()
On Error GoTo ErrorHandler
Dim intLastAccount As Integer
Dim strOS As String
    'Uninstall flag
    If Left(LCase(Command()), 2) = "-u" Then
        Associate False
        End
    End If
    'Verify if there is another instance running
    If App.PrevInstance Then
        ActivatePrevInstance Command()
    End If
    'Verify Minimum IE Version
    #If compIE Then
    If GetIeVer() < "5.0" Then
        MsgBox GetMsg(msgMinIEVersion), vbCritical
        End
    End If
    #End If
    'Initialize Skinned controls on XP or 2003 Server
    strOS = GetOSVer()
    If Left(strOS, 10) = "Windows XP" Or Left(strOS, 12) = "Windows 2003" Then
        Call InitCommonControls
        gIsXP = True
    End If
    'Check if has Portable Settings File
    If FolderExists(App.Path & "\Data") Then
        gAppDataPath = App.Path & "\Data"
    Else
        'Configure the User Application Path
        gAppDataPath = GetShellAppdataLocation(0) & "\w.bloggar"
        If Not FolderExists(gAppDataPath) Then
            CreatePath gAppDataPath
        End If
    End If
    'Check for older versions
    If FileExists(gAppDataPath & XML_SETTINGS) Then
        LoadAccount True
        intLastAccount = gAccount.Current
    Else
        If CheckUpgrade() Then
            frmProgress.RunJob "ConvertSettings"
            LoadAccount True
            intLastAccount = gAccount.Current
            Unload frmPost
        End If
    End If
    'Load Application Settings
    LoadAppSettings
    If gAccount.Current = -1 Then 'No Account
        frmAccountWiz.Show vbModal
    ElseIf Not gAccount.SavePassword Then 'User Autentication
        frmLogin.Show vbModal
    End If
    'Open Main Window
    If gAccount.User <> "" Then
        frmPost.Show
        If Command$ <> "" Then
            'Load command line File Path
            If FileExists(Command$) Then
                frmPost.LoadPostFile Command$
                Exit Sub
            End If
        End If
        If gSettings.OpenLastFile And FileExists(gSettings.PostFile) Then
            DoEvents
            frmPost.LoadPostFile gSettings.PostFile
        Else
            frmPost.NewPost
        End If
    Else
        Unload frmPost
        End
    End If
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, "Main"
    End
End Sub

Public Function CheckUpgrade() As Boolean
On Error GoTo ErrorHandler
Dim strReg As String, strBlog As String
Dim strPage As String, a As Integer
Dim intCMS As Integer, bolSvc As Boolean
Dim strSec As String, strIni As String
Dim strVer As String
    'Verify if Upgrade is necessary
    gAccount.Current = GetSetting(REGISTRY_KEY, "Settings", "Account", "-1")
    strVer = Left(GetSetting(REGISTRY_KEY, "Settings", "Version", "0"), 1)
    If Val(strVer) = 4 Then 'No upgrade needed
        CheckUpgrade = True 'Load Account
        Exit Function
    End If
    If gAccount.Current >= 0 Then 'v2.x or newer
        strReg = "Accounts\" & Format(gAccount.Current, "00")
        gAccount.Alias = GetSetting(REGISTRY_KEY, strReg, "Alias", "")
        If gAccount.Alias <> "" Then 'v3.x
            'Upgrade to 4.x
            Call ShellMoveFile(0, App.Path & "\*.xml", gAppDataPath, False)
            On Error Resume Next
            Kill App.Path + "\*.chg"
            Kill App.Path + "\*.htm"
            CheckUpgrade = True 'Load Account
            Exit Function
        End If
        'Upgrade from 2.x to 4.x
        For a = 0 To 99
            strReg = "Accounts\" & Format(a, "00")
            strPage = GetSetting(REGISTRY_KEY, strReg, "Page", "*")
            If strPage = "*" Then Exit For
            SaveSetting REGISTRY_KEY, strReg, "Alias", Replace(GetLbl(lblAccount), "&", "") & " " & (a + 1)
            If InStr(1, strPage, "/api/RPC2", vbTextCompare) Then
                intCMS = CMS_BLOGGER
            ElseIf InStr(1, strPage, "mt-xmlrpc", vbTextCompare) Then
                intCMS = CMS_MT
            ElseIf InStr(1, strPage, "xmlrpc.php", vbTextCompare) Then
                intCMS = CMS_B2
            ElseIf InStr(1, strPage, "bloggerapi.php", vbTextCompare) Then
                intCMS = CMS_BBT
            ElseIf InStr(1, strPage, "xmlrpc/server.php", vbTextCompare) Then
                intCMS = CMS_NUCLEUS
            ElseIf InStr(1, strPage, "/rpc.php", vbTextCompare) Then
                intCMS = CMS_BLOGALIA
            ElseIf InStr(1, strPage, "listen.asp", vbTextCompare) Then
                intCMS = CMS_BWXML
            ElseIf InStr(1, strPage, "server.php", vbTextCompare) Then
                intCMS = CMS_XOOPS
            Else 'Custom Tool
                intCMS = 0
            End If
            SaveSetting REGISTRY_KEY, strReg, "CMS", Format(intCMS)
            If intCMS > 0 Then
                strSec = "CMS-" & Format(intCMS, "00")
                strIni = App.Path & "\CMS\CMS.ini"
                SaveSetting REGISTRY_KEY, strReg, "Service", ReadINI(strSec, "Service", strIni)
            End If
            On Error Resume Next
            DeleteSetting REGISTRY_KEY, strReg, "Name"
            DeleteSetting REGISTRY_KEY, strReg, "ID"
            On Error GoTo ErrorHandler
        Next
        On Error Resume Next
        DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Left"
        DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Top"
        DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Height"
        DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Width"
        DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Maximized"
        On Error GoTo ErrorHandler
        Associate True 'Associate to the .post file extension
        'Verify if post.txt exists and rename it to draft.post
        If FileExists(App.Path & "\post.txt") Then
            Name App.Path & "\post.txt" As gAppDataPath & "\draft.post"
            DeleteSetting REGISTRY_KEY, "Settings", "PostFile"
        End If
        Call ShellMoveFile(0, App.Path & "\*.xml", gAppDataPath, False)
        On Error Resume Next
        Kill App.Path + "\*.chg"
        Kill App.Path + "\*.htm"
        CheckUpgrade = True 'Load Account
        Exit Function
    End If
    gAccount.User = GetSetting(REGISTRY_KEY, "Settings", "User")
    If gAccount.User = "" Then 'New Installation
        Associate True 'Associate to the .post file extension
        CheckUpgrade = False
        Exit Function
    End If
    'Upgrade from 1.x to 4.x
    gAccount.SavePassword = Val(GetSetting(REGISTRY_KEY, "Settings", "SavePassword", "1"))
    If gAccount.SavePassword Then
        gAccount.Password = Decrypt(GetSetting(REGISTRY_KEY, "Settings", "Password"), "blg")
    End If
    gAccount.Host = GetSetting(REGISTRY_KEY, "Settings", "Host", "plant.blogger.com")
    gAccount.Page = GetSetting(REGISTRY_KEY, "Settings", "Page", "/api/RPC2")
    gAccount.Port = Val(GetSetting(REGISTRY_KEY, "Settings", "Port", "80"))
    gAccount.Secure = Val(GetSetting(REGISTRY_KEY, "Settings", "Secure", "0"))
    gAccount.UseProxy = Val(GetSetting(REGISTRY_KEY, "Settings", "UseProxy", "0"))
    gAccount.ProxyServer = GetSetting(REGISTRY_KEY, "Settings", "ProxyServer", "")
    gAccount.ProxyPort = Val(GetSetting(REGISTRY_KEY, "Settings", "ProxyPort", ""))
    On Error Resume Next
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Left"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Top"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Height"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Width"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Maximized"
    On Error GoTo ErrorHandler
    strBlog = GetSetting(REGISTRY_KEY, "Settings", "Blog", "0")
    'Upgrade to version 2.0 and 3.0 MultiAccount
    gAccount.Current = 0
    strReg = "Accounts\" & Format(gAccount.Current, "00")
    Call SaveSetting(REGISTRY_KEY, "Settings", "Account", Format(gAccount.Current))
    Call SaveSetting(REGISTRY_KEY, strReg, "User", gAccount.User)
    If gAccount.SavePassword Then
        Call SaveSetting(REGISTRY_KEY, strReg, "Password", Encrypt(gAccount.Password, "blg"))
    End If
    Call SaveSetting(REGISTRY_KEY, strReg, "SavePassword", Format(Abs(gAccount.SavePassword)))
    Call SaveSetting(REGISTRY_KEY, strReg, "Host", gAccount.Host)
    Call SaveSetting(REGISTRY_KEY, strReg, "Page", gAccount.Page)
    Call SaveSetting(REGISTRY_KEY, strReg, "Port", Format(gAccount.Port))
    Call SaveSetting(REGISTRY_KEY, strReg, "Secure", Format(Abs(gAccount.Secure)))
    Call SaveSetting(REGISTRY_KEY, strReg, "UseProxy", Format(gAccount.UseProxy))
    Call SaveSetting(REGISTRY_KEY, strReg, "ProxyServer", gAccount.ProxyServer)
    Call SaveSetting(REGISTRY_KEY, strReg, "ProxyPort", Format(gAccount.ProxyPort))
    Call SaveSetting(REGISTRY_KEY, strReg, "Blog", strBlog)
    If FileExists(App.Path & "\blogs.xml") Then
        ShellMoveFile 0, App.Path & "\blogs.xml", gAppDataPath & "\blogs" & Format(gAccount.Current, "00") & ".xml"
    End If
    Associate True 'Associate to the .post file extension
    'Verify if post.txt exists and rename it to draft.post
    If FileExists(App.Path & "\post.txt") Then
        Name App.Path & "\post.txt" As gAppDataPath & "\draft.post"
        DeleteSetting REGISTRY_KEY, "Settings", "PostFile"
    End If
    Call ShellMoveFile(0, App.Path & "\*.xml", gAppDataPath, False)
    On Error Resume Next
    Kill App.Path + "\*.chg"
    Kill App.Path + "\*.htm"
    CheckUpgrade = True
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, "CheckUpgrade"
End Function

Public Sub LoadAccount(Optional ByVal bolActive As Boolean)
On Error Resume Next
Dim strReg As String
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    If bolActive Then
        gAccount.Current = objXMLReg.GetSetting(App.Title, "Settings", "Account", "0")
    End If
    strReg = "Accounts/a" & Format(gAccount.Current, "00")
    'Account Info
    gAccount.Alias = objXMLReg.GetSetting(App.Title, strReg, "Alias")
    gAccount.User = objXMLReg.GetSetting(App.Title, strReg, "User")
    gAccount.SavePassword = Val(objXMLReg.GetSetting(App.Title, strReg, "SavePassword", "1"))
    If gAccount.SavePassword Then
        gAccount.Password = Decrypt(objXMLReg.GetSetting(App.Title, strReg, "Password"), "blg")
    Else
        gAccount.Password = ""
    End If
    gAccount.PingWeblogs = Val(objXMLReg.GetSetting(App.Title, strReg, "PingWeblogs", "-1"))
    gAccount.CMS = Val(objXMLReg.GetSetting(App.Title, strReg, "CMS"))
    gAccount.Service = Val(objXMLReg.GetSetting(App.Title, strReg, "Service", "0"))
    If gAccount.Service Then
        gAccount.Host = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Host", App.Path & "\CMS\CMS.ini")
        gAccount.Page = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Page", App.Path & "\CMS\CMS.ini")
        gAccount.Port = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Port", App.Path & "\CMS\CMS.ini")
        gAccount.Secure = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Https", App.Path & "\CMS\CMS.ini"))
    Else
        gAccount.Host = objXMLReg.GetSetting(App.Title, strReg, "Host")
        gAccount.Page = objXMLReg.GetSetting(App.Title, strReg, "Page")
        gAccount.Port = Val(objXMLReg.GetSetting(App.Title, strReg, "Port"))
        gAccount.Secure = Val(objXMLReg.GetSetting(App.Title, strReg, "Secure", "0"))
    End If
    gAccount.UTF8 = Val(objXMLReg.GetSetting(App.Title, strReg, "UTF-8", "1"))
    If gAccount.CMS = CMS_CUSTOM Then
        gAccount.PostMethod = Val(objXMLReg.GetSetting(App.Title, strReg, "PostMethod"))
        gAccount.GetPostsMethod = Val(objXMLReg.GetSetting(App.Title, strReg, "GetPostsMethod"))
        gAccount.GetCategMethod = Val(objXMLReg.GetSetting(App.Title, strReg, "CategMethod"))
        gAccount.TemplateMethod = Val(objXMLReg.GetSetting(App.Title, strReg, "TemplateMethod"))
        gAccount.MultiCategory = Val(objXMLReg.GetSetting(App.Title, strReg, "MultiCategory"))
        gAccount.TitleTag1 = objXMLReg.GetSetting(App.Title, strReg, "TitleTag1")
        gAccount.TitleTag2 = objXMLReg.GetSetting(App.Title, strReg, "TitleTag2")
        gAccount.CategTag1 = objXMLReg.GetSetting(App.Title, strReg, "CategTag1")
        gAccount.CategTag2 = objXMLReg.GetSetting(App.Title, strReg, "CategTag2")
        gAccount.MoreTextTag1 = objXMLReg.GetSetting(App.Title, strReg, "MoreTextTag1")
        gAccount.MoreTextTag2 = objXMLReg.GetSetting(App.Title, strReg, "MoreTextTag2")
        If gAccount.MoreTextTag1 <> "" Or gAccount.MoreTextTag2 <> "" Then
            gAccount.MoreTab = True
            gAccount.Extended = True
        Else
            gAccount.MoreTab = False
            gAccount.Extended = False
        End If
        gAccount.AdvancedOptions = False
        gAccount.Excerpt = False
        gAccount.Keywords = False
    Else
        LoadCMS
    End If
    gAccount.Timeout = Val(objXMLReg.GetSetting(App.Title, strReg, "Timeout", "30"))
    gAccount.UseProxy = Val(objXMLReg.GetSetting(App.Title, strReg, "UseProxy", "0"))
    gAccount.ProxyServer = objXMLReg.GetSetting(App.Title, strReg, "ProxyServer", "")
    gAccount.ProxyPort = Val(objXMLReg.GetSetting(App.Title, strReg, "ProxyPort", ""))
    gAccount.ProxyUser = objXMLReg.GetSetting(App.Title, strReg, "ProxyUser", "")
    gAccount.ProxyPassword = objXMLReg.GetSetting(App.Title, strReg, "ProxyPassword", "")
    Set objXMLReg = Nothing
End Sub

Public Sub LoadCMS()
    gAccount.GetCategMethod = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "GetCategMethod", App.Path & "\CMS\CMS.ini"))
    gAccount.PostMethod = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "PostMethod", App.Path & "\CMS\CMS.ini"))
    gAccount.GetPostsMethod = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "GetPostsMethod", App.Path & "\CMS\CMS.ini"))
    gAccount.TemplateMethod = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "TemplateMethod", App.Path & "\CMS\CMS.ini"))
    gAccount.MultiCategory = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "MultiCateg", App.Path & "\CMS\CMS.ini"))
    gAccount.TitleTag1 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "TitleTag1", App.Path & "\CMS\CMS.ini")
    gAccount.TitleTag2 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "TitleTag2", App.Path & "\CMS\CMS.ini")
    gAccount.CategTag1 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "CategTag1", App.Path & "\CMS\CMS.ini")
    gAccount.CategTag2 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "CategTag2", App.Path & "\CMS\CMS.ini")
    '3.03
    gAccount.BodyTag1 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "BodyTag1", App.Path & "\CMS\CMS.ini")
    gAccount.BodyTag2 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "BodyTag2", App.Path & "\CMS\CMS.ini")
    gAccount.MoreTextTag1 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "MoreTextTag1", App.Path & "\CMS\CMS.ini")
    gAccount.MoreTextTag2 = ReadINI("CMS-" & Format(gAccount.CMS, "00"), "MoreTextTag2", App.Path & "\CMS\CMS.ini")
    '4.02
    gAccount.UTF8OnPost = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "UTF8OnPost", App.Path & "\CMS\CMS.ini"))
    gAccount.MoreTab = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "MoreTab", App.Path & "\CMS\CMS.ini"))
    gAccount.AdvancedOptions = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "AdvancedOptions", App.Path & "\CMS\CMS.ini"))
    gAccount.AllowComments = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "AllowComments", App.Path & "\CMS\CMS.ini"))
    gAccount.AllowPings = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "AllowPings", App.Path & "\CMS\CMS.ini"))
    gAccount.TextFilters = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "TextFilters", App.Path & "\CMS\CMS.ini"))
    gAccount.PostDate = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "PostDate", App.Path & "\CMS\CMS.ini"))
    gAccount.TrackBack = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "TrackBack", App.Path & "\CMS\CMS.ini"))
    gAccount.Extended = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Extended", App.Path & "\CMS\CMS.ini"))
    gAccount.Excerpt = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Excerpt", App.Path & "\CMS\CMS.ini"))
    gAccount.Keywords = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "Keywords", App.Path & "\CMS\CMS.ini"))
    gAccount.UploadMethod = Val(ReadINI("CMS-" & Format(gAccount.CMS, "00"), "UploadMethod", App.Path & "\CMS\CMS.ini"))
End Sub

Public Sub LoadCMSCombo(cboCMS As Control, picCustom As StdPicture, Optional bolNew As Boolean = False)
On Error Resume Next
Dim i As Integer, c As Integer, intCount As Integer
Dim strIni As String, strIco As String
    'Load CMS from INI
    strIni = App.Path & "\CMS\CMS.ini"
    intCount = ReadINI("CMS", "Count", strIni)
    cboCMS.Clear
    For i = 2 To intCount
        strIco = App.Path & "\CMS\" & ReadINI("CMS-" & Format(i, "00"), "Icon", strIni)
        If FileExists(strIco) Then
            cboCMS.AddIcon LoadPicture(strIco)
        Else
            cboCMS.AddIcon picCustom
        End If
        cboCMS.AddItem ReadINI("CMS-" & Format(i, "00"), "Name", strIni), , i - 1
        cboCMS.ItemData(i - 2) = i
    Next
    cboCMS.AddIcon picCustom
    cboCMS.AddItem "(" & GetLbl(lblCustom) & ")", , i - 1
    If bolNew Then
        cboCMS.ListIndex = 0
    Else
        If gAccount.CMS > CMS_CUSTOM Then
            If gAccount.CMS > CMS_BLOGGER Then
                cboCMS.ListIndex = gAccount.CMS - 2
            Else
                cboCMS.ListIndex = 0
            End If
        Else
            cboCMS.ListIndex = cboCMS.ListCount - 1
        End If
    End If
End Sub

Public Sub LoadPingCombo(cboPing As Control, Optional bolNew As Boolean = False)
On Error Resume Next
Dim i As Integer, intCount As Integer
Dim strIni As String, strIco As String
    'Load CMS from INI
    strIni = App.Path & "\CMS\Ping.ini"
    intCount = ReadINI("Ping", "Count", strIni)
    cboPing.Clear
    For i = 1 To intCount
        cboPing.AddItem ReadINI("Ping-" & Format(i, "00"), "Name", strIni)
    Next
    If bolNew Then
        cboPing.ListIndex = 0
    Else
        If gAccount.PingWeblogs > 0 Then
            cboPing.ListIndex = gAccount.PingWeblogs - 1
        Else
            cboPing.ListIndex = 0
        End If
    End If
End Sub

Public Sub SaveAccount()
On Error Resume Next
Dim strReg As String
    Set objXMLReg = New XMLRegistry
    strReg = "Accounts/a" & Format(gAccount.Current, "00")
    'Account Info
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    Call objXMLReg.SaveSetting(App.Title, "Settings", "Account", Format(gAccount.Current))
    Call objXMLReg.SaveSetting(App.Title, strReg, "Alias", gAccount.Alias)
    Call objXMLReg.SaveSetting(App.Title, strReg, "User", gAccount.User)
    If gAccount.SavePassword Then
        Call objXMLReg.SaveSetting(App.Title, strReg, "Password", Encrypt(gAccount.Password, "blg"))
    Else
        Call objXMLReg.SaveSetting(App.Title, strReg, "Password", "")
    End If
    Call objXMLReg.SaveSetting(App.Title, strReg, "SavePassword", Format(Abs(gAccount.SavePassword)))
    Call objXMLReg.SaveSetting(App.Title, strReg, "PingWeblogs", Format(gAccount.PingWeblogs))
    Call objXMLReg.SaveSetting(App.Title, strReg, "CMS", gAccount.CMS)
    Call objXMLReg.SaveSetting(App.Title, strReg, "Service", Format(Abs(gAccount.Service)))
    If Not gAccount.Service Then
        Call objXMLReg.SaveSetting(App.Title, strReg, "Host", gAccount.Host)
        Call objXMLReg.SaveSetting(App.Title, strReg, "Page", gAccount.Page)
        Call objXMLReg.SaveSetting(App.Title, strReg, "Port", Format(gAccount.Port))
        Call objXMLReg.SaveSetting(App.Title, strReg, "Secure", Format(Abs(gAccount.Secure)))
    End If
    Call objXMLReg.SaveSetting(App.Title, strReg, "UTF-8", Format(Abs(gAccount.UTF8)))
    If gAccount.CMS = CMS_CUSTOM Then
        Call objXMLReg.SaveSetting(App.Title, strReg, "PostMethod", gAccount.PostMethod)
        Call objXMLReg.SaveSetting(App.Title, strReg, "GetPostsMethod", gAccount.GetPostsMethod)
        Call objXMLReg.SaveSetting(App.Title, strReg, "CategMethod", gAccount.GetCategMethod)
        Call objXMLReg.SaveSetting(App.Title, strReg, "TemplateMethod", gAccount.TemplateMethod)
        Call objXMLReg.SaveSetting(App.Title, strReg, "MultiCategory", Format(Abs(gAccount.MultiCategory)))
        Call objXMLReg.SaveSetting(App.Title, strReg, "TitleTag1", gAccount.TitleTag1)
        Call objXMLReg.SaveSetting(App.Title, strReg, "TitleTag2", gAccount.TitleTag2)
        Call objXMLReg.SaveSetting(App.Title, strReg, "CategTag1", gAccount.CategTag1)
        Call objXMLReg.SaveSetting(App.Title, strReg, "CategTag2", gAccount.CategTag2)
        Call objXMLReg.SaveSetting(App.Title, strReg, "MoreTextTag1", gAccount.MoreTextTag1)
        Call objXMLReg.SaveSetting(App.Title, strReg, "MoreTextTag2", gAccount.MoreTextTag2)
    End If
    Call objXMLReg.SaveSetting(App.Title, strReg, "Timeout", Format(gAccount.Timeout))
    Call objXMLReg.SaveSetting(App.Title, strReg, "UseProxy", Format(gAccount.UseProxy))
    Call objXMLReg.SaveSetting(App.Title, strReg, "ProxyServer", gAccount.ProxyServer)
    Call objXMLReg.SaveSetting(App.Title, strReg, "ProxyPort", Format(gAccount.ProxyPort))
    Call objXMLReg.SaveSetting(App.Title, strReg, "ProxyUser", gAccount.ProxyUser)
    Call objXMLReg.SaveSetting(App.Title, strReg, "ProxyPassword", gAccount.ProxyPassword)
    Call objXMLReg.SaveSetting(App.Title, strReg, "Deleted", "0")
    'Commented 4.00 - Call DeleteSetting(REGISTRY_KEY, strReg, "Deleted")
    Set objXMLReg = Nothing
End Sub

Public Sub DeleteAccount(Optional ByVal strAccount As String)
On Error Resume Next
Dim strReg As String
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    If strAccount = "" Then strAccount = gAccount.Current
    strReg = "Accounts/a" & Format(strAccount, "00")
    'Account Info
    Call objXMLReg.SaveSetting(App.Title, strReg, "Deleted", "1")
    Call objXMLReg.SaveSetting(App.Title, strReg, "User", "�Deleted�")
    Call objXMLReg.SaveSetting(App.Title, strReg, "ID", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Alias", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Password", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Name", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "SavePassword", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "PingWeblogs", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "CMS", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Service", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Host", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Page", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Port", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Secure", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "Timeout", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "UseProxy", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "ProxyServer", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "ProxyPort", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "PostMethod", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "CategMethod", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "TemplateMethod", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "MultiCategory", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "TitleTag1", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "TitleTag2", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "CategTag1", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "CategTag2", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "MoreTextTag1", "")
    Call objXMLReg.SaveSetting(App.Title, strReg, "MoreTextTag2", "")
    If FileExists(gAppDataPath & "\blogs" & Format(strAccount, "00") & ".xml") Then
        Kill gAppDataPath & "\blogs" & Format(strAccount, "00") & ".xml"
    End If
    Set objXMLReg = Nothing
End Sub

Public Sub Uninstall()
On Error GoTo ErrorHandler
Dim strReg As String, strUser As String
Dim strRoot As String, a As Integer
Dim aBlogs() As String
    'Delete Accounts
    strRoot = "Software\VB and VBA Program Settings\" & REGISTRY_KEY & "\"
    For a = 0 To 99
        strReg = "Accounts\" & Format(a, "00")
        strUser = GetSetting(REGISTRY_KEY, strReg, "User", "*")
        If strUser = "*" Then Exit For
        DeleteKey HKEY_CURRENT_USER, strRoot & strReg
    Next
    DeleteKey HKEY_CURRENT_USER, strRoot & "Accounts"
    'Delete Blogs
    aBlogs = EnumKeys(HKEY_CURRENT_USER, strRoot & "Blogs")
    If Not IsArrayEmpty(aBlogs) Then
        For a = 0 To UBound(aBlogs)
            DeleteKey HKEY_CURRENT_USER, strRoot & "Blogs\" & aBlogs(a)
        Next
    End If
    DeleteKey HKEY_CURRENT_USER, strRoot & "Blogs"
    'Delete Other Keys
    DeleteKey HKEY_CURRENT_USER, strRoot & "Colors"
    DeleteKey HKEY_CURRENT_USER, strRoot & "Forms"
    DeleteKey HKEY_CURRENT_USER, strRoot & "Images"
    DeleteKey HKEY_CURRENT_USER, strRoot & "Links"
    DeleteKey HKEY_CURRENT_USER, strRoot & "Search"
    DeleteKey HKEY_CURRENT_USER, strRoot & "Settings"
    DeleteKey HKEY_CURRENT_USER, strRoot & "Table"
    DeleteKey HKEY_CURRENT_USER, strRoot & "MRU"
    'Delete Root
    DeleteKey HKEY_CURRENT_USER, strRoot
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Public Sub LoadAppSettings()
Dim lngDefaultLCID As Long
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    gSettings.Tray = Val(objXMLReg.GetSetting(App.Title, "Settings", "Tray", "1"))
    gSettings.ClearPost = Val(objXMLReg.GetSetting(App.Title, "Settings", "ClearPost", "1"))
    gSettings.PostFile = objXMLReg.GetSetting(App.Title, "Settings", "PostFile", gAppDataPath & "\draft.post")
    gSettings.PostTemplate = objXMLReg.GetSetting(App.Title, "Settings", "PostTemplate", "")
    gSettings.FontFace = objXMLReg.GetSetting(App.Title, "Settings", "FontFace", "Arial")
    gSettings.FontSize = Val(objXMLReg.GetSetting(App.Title, "Settings", "FontSize", "8"))
    gSettings.TabSpaces = Val(objXMLReg.GetSetting(App.Title, "Settings", "TabSpaces", "3"))
    gSettings.ColorizeCode = Val(objXMLReg.GetSetting(App.Title, "Settings", "ColorizeCode", "1"))
    gSettings.StartMinimized = Val(objXMLReg.GetSetting(App.Title, "Settings", "StartMinimized", "0"))
    gSettings.ShowHtmlBar = Val(objXMLReg.GetSetting(App.Title, "Settings", "ShowHtmlBar", "1"))
    gSettings.AutoConvert = Val(objXMLReg.GetSetting(App.Title, "Settings", "AutoConvert", "0"))
    gSettings.SkinFolder = App.Path & "\Skins\" & objXMLReg.GetSetting(App.Title, "Settings", "Skin", "Windows XP")
    If Not FileExists(gSettings.SkinFolder & "\skin.ini") Then
        gSettings.SkinFolder = App.Path & "\Skins\Windows XP"
    End If
    gSettings.Silent = Val(objXMLReg.GetSetting(App.Title, "Settings", "Silent", "0"))
    gSettings.DefaultBrowser = Val(objXMLReg.GetSetting(App.Title, "Settings", "DefaultBrowser", "0"))
    If gSettings.SpellLCID = 0 Then
        lngDefaultLCID = GetUserDefaultLCID()
        If Not FileExists(App.Path & "\Lang\" & lngDefaultLCID & ".lng") Then
            lngDefaultLCID = 1033
        End If
    End If
    gSettings.AppLCID = Val(objXMLReg.GetSetting(App.Title, "Settings", "AppLCID", lngDefaultLCID))
    If gSettings.SpellLCID = 0 Then
        lngDefaultLCID = GetUserDefaultLCID()
        If Not FileExists(App.Path & "\Spell\" & lngDefaultLCID & ".dic") Then
            lngDefaultLCID = 1033
        End If
    End If
    gSettings.SpellLCID = Val(objXMLReg.GetSetting(App.Title, "Settings", "SpellLCID", lngDefaultLCID))
    gSettings.OpenLastFile = Val(objXMLReg.GetSetting(App.Title, "Settings", "OpenLastFile", "1"))
    gSettings.XHTML = Val(objXMLReg.GetSetting(App.Title, "Settings", "XHTML", "0"))
    gSettings.BlogListSize = Val(objXMLReg.GetSetting(App.Title, "Settings", "BlogListSize", "110"))
    gSettings.PostMenu = objXMLReg.GetSetting(App.Title, "Settings", "PostMenu", "5")
    Set objXMLReg = Nothing
End Sub

Public Sub SaveAppSettings()
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    Call objXMLReg.SaveSetting(App.Title, "Settings", "Version", App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000"))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "Tray", Format(Abs(gSettings.Tray)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "ClearPost", Format(Abs(gSettings.ClearPost)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "PostFile", gSettings.PostFile)
    Call objXMLReg.SaveSetting(App.Title, "Settings", "PostTemplate", gSettings.PostTemplate)
    Call objXMLReg.SaveSetting(App.Title, "Settings", "FontFace", gSettings.FontFace)
    Call objXMLReg.SaveSetting(App.Title, "Settings", "FontSize", Format(gSettings.FontSize))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "TabSpaces", Format(gSettings.TabSpaces))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "ColorizeCode", Format(Abs(gSettings.ColorizeCode)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "StartMinimized", Format(Abs(gSettings.StartMinimized)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "ShowHtmlBar", Format(Abs(gSettings.ShowHtmlBar)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "AutoConvert", Format(Abs(gSettings.AutoConvert)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "Skin", GetNamePart(gSettings.SkinFolder, True))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "Silent", Format(Abs(gSettings.Silent)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "DefaultBrowser", Format(Abs(gSettings.DefaultBrowser)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "AppLCID", Format(gSettings.AppLCID))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "SpellLCID", Format(gSettings.SpellLCID))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "OpenLastFile", Format(Abs(gSettings.OpenLastFile)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "XHTML", Format(Abs(gSettings.XHTML)))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "BlogListSize", Format(gSettings.BlogListSize))
    Call objXMLReg.SaveSetting(App.Title, "Settings", "PostMenu", gSettings.PostMenu)
    Set objXMLReg = Nothing
End Sub

Public Sub LoadColors()
Dim i As Integer
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    'Create ColorDialog object
    Set objColor = New ColorDialog
    objColor.Color = Val(objXMLReg.GetSetting(App.Title, "Colors", "Default", "0"))
    'Load custom colors
    For i = 0 To 15
        objColor.CustomColors(i) = Val(objXMLReg.GetSetting(App.Title, "Colors", "Custom" & Format(i, "00"), "0"))
    Next
    Set objXMLReg = Nothing
End Sub

Public Sub SaveColors()
Dim i As Integer
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    Call objXMLReg.SaveSetting(App.Title, "Colors", "Default", objColor.Color)
    For i = 0 To 15
        Call objXMLReg.SaveSetting(App.Title, "Colors", "Custom" & Format(i, "00"), objColor.CustomColors(i))
    Next
End Sub

Public Sub LoadCustomTags()
Dim i As Integer, strAux As String
On Error Resume Next
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    For i = 1 To 12
        strAux = objXMLReg.GetSetting(App.Title, "CustomTags", "CustomTag" & i)
        If strAux <> "" Then
            gSettings.CustomTag(i, 1) = Split(strAux, vbTab)(0)
            gSettings.CustomTag(i, 2) = Split(strAux, vbTab)(1)
            gSettings.CustomTag(i, 3) = Split(strAux, vbTab)(2)
            frmPost.acbMain.Bands("bndPopCustom").Tools("miCustomF" & i).Caption = gSettings.CustomTag(i, 1)
        End If
    Next
    Set objXMLReg = Nothing
End Sub

Public Sub SaveCustomTags()
Dim i As Integer, strAux As String
On Error Resume Next
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    For i = 1 To 12
        strAux = ""
        If gSettings.CustomTag(i, 1) <> "" Then
            strAux = gSettings.CustomTag(i, 1) & vbTab & _
                     gSettings.CustomTag(i, 2) & vbTab & _
                     gSettings.CustomTag(i, 3)
        End If
        Call objXMLReg.SaveSetting(App.Title, "CustomTags", "CustomTag" & i, strAux)
    Next
    Set objXMLReg = Nothing
End Sub

Public Sub LoadBlogSettings()
Dim strReg As String
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    If gAccount.CMS = CMS_BLOGGER Or _
       gAccount.CMS = CMS_BLOGGERPRO Then
        strReg = "Blogs/b" & gBlogs(frmPost.CurrentBlog).BlogID
    Else
        strReg = "Blogs/b" & Format(gAccount.Current, "00") & "-" & gBlogs(frmPost.CurrentBlog).BlogID
    End If
    gBlog.PreviewBody = objXMLReg.GetSetting(App.Title, strReg, "PreviewBody", BODYTAG)
    gBlog.PreviewCSS = objXMLReg.GetSetting(App.Title, strReg, "PreviewCSS", CSSTAG)
    gBlog.PreviewTitle = objXMLReg.GetSetting(App.Title, strReg, "PreviewTitle", TITLESTYLE)
    gBlog.PreviewStyle = objXMLReg.GetSetting(App.Title, strReg, "PreviewStyle", POSTSTYLE)
    gBlog.PreviewAlign = objXMLReg.GetSetting(App.Title, strReg, "PreviewAlign", POSTALIGN)
    gBlog.PreviewWidth = objXMLReg.GetSetting(App.Title, strReg, "PreviewWidth", POSTWIDTH)
    gBlog.PreviewAutoBR = Val(objXMLReg.GetSetting(App.Title, strReg, "PreviewAutoBR", "1"))
    gBlog.APIUpload = Val(objXMLReg.GetSetting(App.Title, strReg, "APIUpload", IIf(gAccount.UploadMethod = API_METAWEBLOG, "1", "0")))
    gBlog.FTPHost = objXMLReg.GetSetting(App.Title, strReg, "FTPHost")
    gBlog.FTPPath = objXMLReg.GetSetting(App.Title, strReg, "FTPPath")
    gBlog.FTPPort = Val(objXMLReg.GetSetting(App.Title, strReg, "FTPPort", "21"))
    gBlog.FTPProxy = Val(objXMLReg.GetSetting(App.Title, strReg, "FTPProxy", "0"))
    gBlog.FTPUser = objXMLReg.GetSetting(App.Title, strReg, "FTPUser")
    gBlog.FTPPassword = Decrypt(objXMLReg.GetSetting(App.Title, strReg, "FTPPassword"), "blg")
    gBlog.FTPLink = objXMLReg.GetSetting(App.Title, strReg, "FTPLink")
    gBlog.MediaInsert = Val(objXMLReg.GetSetting(App.Title, strReg, "MediaInsert", "0"))
    gBlog.MediaLink = Val(objXMLReg.GetSetting(App.Title, strReg, "MediaLink", "0"))
    gBlog.MediaString = objXMLReg.GetSetting(App.Title, strReg, "MediaString")
    Set objXMLReg = Nothing
End Sub

Public Sub SaveBlogSettings()
Dim strReg As String
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    If gAccount.CMS = CMS_BLOGGER Or _
       gAccount.CMS = CMS_BLOGGERPRO Then
        strReg = "Blogs/b" & gBlogs(frmPost.CurrentBlog).BlogID
    Else
        strReg = "Blogs/b" & Format(gAccount.Current, "00") & "-" & gBlogs(frmPost.CurrentBlog).BlogID
    End If
    Call objXMLReg.SaveSetting(App.Title, strReg, "PreviewBody", gBlog.PreviewBody)
    Call objXMLReg.SaveSetting(App.Title, strReg, "PreviewCSS", gBlog.PreviewCSS)
    Call objXMLReg.SaveSetting(App.Title, strReg, "PreviewTitle", gBlog.PreviewTitle)
    Call objXMLReg.SaveSetting(App.Title, strReg, "PreviewStyle", gBlog.PreviewStyle)
    Call objXMLReg.SaveSetting(App.Title, strReg, "PreviewAlign", gBlog.PreviewAlign)
    Call objXMLReg.SaveSetting(App.Title, strReg, "PreviewWidth", gBlog.PreviewWidth)
    Call objXMLReg.SaveSetting(App.Title, strReg, "PreviewAutoBR", Format(Abs(gBlog.PreviewAutoBR)))
    Call objXMLReg.SaveSetting(App.Title, strReg, "FTPHost", gBlog.FTPHost)
    Call objXMLReg.SaveSetting(App.Title, strReg, "FTPPath", gBlog.FTPPath)
    Call objXMLReg.SaveSetting(App.Title, strReg, "FTPPort", Format(gBlog.FTPPort))
    Call objXMLReg.SaveSetting(App.Title, strReg, "FTPProxy", Format(Abs(gBlog.FTPProxy)))
    Call objXMLReg.SaveSetting(App.Title, strReg, "FTPUser", gBlog.FTPUser)
    If Trim(gBlog.FTPPassword) <> "" Then
        Call objXMLReg.SaveSetting(App.Title, strReg, "FTPPassword", Encrypt(gBlog.FTPPassword, "blg"))
    Else
        Call objXMLReg.SaveSetting(App.Title, strReg, "FTPPassword", "")
    End If
    Call objXMLReg.SaveSetting(App.Title, strReg, "FTPLink", gBlog.FTPLink)
    Call objXMLReg.SaveSetting(App.Title, strReg, "MediaInsert", Format(gBlog.MediaInsert))
    Call objXMLReg.SaveSetting(App.Title, strReg, "MediaLink", Format(Abs(gBlog.MediaLink)))
    Call objXMLReg.SaveSetting(App.Title, strReg, "MediaString", gBlog.MediaString)
    Set objXMLReg = Nothing
End Sub

Public Function GetMsg(ByVal ID As MessagesEnum) As String
    GetMsg = frmPost.acbMain.Tools(Format(30000 + ID)).Description
End Function

Public Function GetLbl(ByVal ID As LabelsEnum) As String
    GetLbl = frmPost.acbMain.Tools(Format(35000 + ID)).Caption
End Function

Public Function GetBinaryFile(ByVal FileName As String) As String
On Error GoTo ErrorHandler
Dim iFile As Integer
Dim sBuffer As String
    FileName = Replace(FileName, """", "")
    sBuffer = String(FileLen(FileName), Chr(0))
    iFile = FreeFile
    Open FileName For Binary Access Read As #iFile
    Get #iFile, , sBuffer
    Close #iFile
    GetBinaryFile = sBuffer
    Exit Function
ErrorHandler:
    GetBinaryFile = ""
End Function

Public Function SaveBinaryFile(ByVal FileName As String, ByVal Buffer As String) As Boolean
On Error GoTo ErrorHandler
Dim iFile As Integer
Dim sBuffer As String
    iFile = FreeFile
    Open FileName For Output As #iFile 'Truncate File
    Close #iFile
    DoEvents
    Open FileName For Binary Access Write As #iFile
    Put #iFile, , Buffer
    Close #iFile
    SaveBinaryFile = True
    Exit Function
ErrorHandler:
    SaveBinaryFile = False
End Function

Public Function FileExists(ByVal strFile As String) As Boolean
    Dim nArq%
    nArq% = FreeFile
    On Error Resume Next
    strFile = Replace(strFile, """", "")
    Open strFile$ For Input As #nArq%
    If Err = 0 Then
       FileExists = True
    Else
       FileExists = False
    End If
    Close #nArq%
End Function

Public Function Encrypt(ByVal vText$, ByVal vKey$) As String
On Error Resume Next
   Dim i%, n%, p%, s&, t&, Aux$
   s& = 1325
   For i% = 1 To Len(vKey$)
      s& = s& + Asc(Mid$(vKey$, i%, 1)) * ((Len(vKey$) + 1) - i%)
   Next
   t& = 5213
   n% = 1
   Aux$ = ""
   For i% = 1 To Len(vText$)
      t& = t& + Asc(Mid$(vText$, i%, 1)) * ((Len(vText$) + 1) - i%)
      Aux$ = Aux$ + Chr$(Asc(Mid$(vText$, i%, 1)) - (Val(Mid$(Format$(s&), n%, 1)) + 3))
      If n% > Len(vKey$) Then n% = 1 Else n% = n% + 1
   Next
   p% = Min(Val(Right$(Format$(t&), 1)), Len(vText$) - 1)
   Aux$ = Mid$(Aux$, p% + 1) + Left$(Aux$, p%)
   Aux$ = Left$(Aux$, p%) + Chr$(p% + 70) + Mid$(Aux$, p% + 1)
   Aux$ = Chr$(p% + 65) + Aux$
   Encrypt = Aux$
End Function

Public Function Decrypt(ByVal vText$, ByVal vKey$) As String
On Error Resume Next
Dim i%, n%, p%, s&, Result$
   s& = 1325
   For i% = 1 To Len(vKey$)
      s& = s& + Asc(Mid$(vKey$, i%, 1)) * ((Len(vKey$) + 1) - i%)
   Next
   p% = Asc(Left$(vText$, 1)) - 65
   vText$ = Mid$(vText$, 2)
   vText$ = Left$(vText$, p%) + Mid$(vText$, p% + 2)
   vText$ = Right$(vText$, p%) + Left$(vText$, Len(vText$) - p%)
   n% = 1
   Result$ = ""
   For i% = 1 To Len(vText$)
      Result$ = Result$ + Chr$(Asc(Mid$(vText$, i%, 1)) + (Val(Mid$(Format$(s&), n%, 1)) + 3))
      If n% > Len(vKey$) Then n% = 1 Else n% = n% + 1
   Next
   Decrypt = Result$
End Function

Public Function Min(ParamArray Vals())
Dim n As Integer, MinVal
    MinVal = Vals(0)
    For n = 0 To UBound(Vals)
        If Vals(n) < MinVal Then MinVal = Vals(n)
    Next n
    Min = MinVal
End Function

Public Function Max(ParamArray Vals())
Dim n As Integer, MaxVal
    For n = 0 To UBound(Vals)
        If Vals(n) > MaxVal Then MaxVal = Vals(n)
    Next n
    Max = MaxVal
End Function

Public Sub ActiveComboBox(ctrComboBox As Object, KeyAscii As Integer)
   Dim Index As Long
   Dim FindString As String

   If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub

   If ctrComboBox.SelLength = 0 Then
      FindString = ctrComboBox.Text & Chr$(KeyAscii)
   Else
      FindString = Left$(ctrComboBox.Text, ctrComboBox.SelStart) & Chr$(KeyAscii)
   End If

   Index = SendMessage(ctrComboBox.hwnd, CB_FINDSTRING, -1, ByVal FindString)

   If Index <> CB_ERR Then
      ctrComboBox.ListIndex = Index
      ctrComboBox.SelStart = Len(FindString)
      ctrComboBox.SelLength = Len(ctrComboBox.Text) - ctrComboBox.SelStart
      KeyAscii = 0
   End If
End Sub

Public Function SearchComboBox(ctrComboBox As Object, strText As String) As Boolean
On Error GoTo ErrorHandler
   Dim Index As Long
   Dim FindString As String

   FindString = strText

   Index = SendMessage(ctrComboBox.hwnd, CB_FINDSTRING, -1, ByVal FindString)

   If Index <> CB_ERR Then
      ctrComboBox.ListIndex = Index
      SearchComboBox = True
   End If
    Exit Function
ErrorHandler:
    SearchComboBox = False
End Function

Public Function SearchItemData(ctrComboBox As Object, ByVal ItemData As Variant) As Boolean
On Error GoTo ErrorHandler
Dim Index As Long
    For Index = 0 To ctrComboBox.ListCount - 1
        If ctrComboBox.ItemData(Index) = ItemData Then
            ctrComboBox.ListIndex = Index
            SearchItemData = True
            Exit For
        End If
    Next
    Exit Function
ErrorHandler:
    SearchItemData = False
End Function

Public Function SaveFormSettings(frm As Form)
On Error Resume Next
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    If frm.WindowState = vbNormal Then
        objXMLReg.SaveSetting App.Title, "Forms", frm.Name & ".Left", Format(frm.Left)
        objXMLReg.SaveSetting App.Title, "Forms", frm.Name & ".Top", Format(frm.Top)
        If frm.BorderStyle = 2 Or frm.BorderStyle = 5 Then
            objXMLReg.SaveSetting App.Title, "Forms", frm.Name & ".Width", Format(frm.Width)
            objXMLReg.SaveSetting App.Title, "Forms", frm.Name & ".Height", Format(frm.Height)
        End If
        objXMLReg.SaveSetting App.Title, "Forms", frm.Name & ".Maximized", "0"
    ElseIf frm.WindowState = vbMaximized Then
        objXMLReg.SaveSetting App.Title, "Forms", frm.Name & ".Maximized", "1"
    Else
        objXMLReg.SaveSetting App.Title, "Forms", frm.Name & ".Maximized", "0"
    End If
    Set objXMLReg = Nothing
End Function

Public Function LoadFormSettings(frm As Form, Optional ByVal DefaultLeft, _
                                              Optional ByVal DefaultTop, _
                                              Optional ByVal DefaultWidth, _
                                              Optional ByVal DefaultHeight, _
                                              Optional ByVal Name)
On Error Resume Next
    Set objXMLReg = New XMLRegistry
    objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
    If IsMissing(DefaultLeft) Then DefaultLeft = frm.Left
    If IsMissing(DefaultTop) Then DefaultTop = frm.Top
    If IsMissing(DefaultWidth) Then DefaultWidth = 10050
    If IsMissing(DefaultHeight) Then DefaultHeight = 7005
    frm.Left = objXMLReg.GetSetting(App.Title, "Forms", frm.Name & ".Left", DefaultLeft)
    frm.Top = objXMLReg.GetSetting(App.Title, "Forms", frm.Name & ".Top", DefaultTop)
    If frm.BorderStyle = 2 Or frm.BorderStyle = 5 Then
        frm.Width = objXMLReg.GetSetting(App.Title, "Forms", frm.Name & ".Width", DefaultWidth)
        frm.Height = objXMLReg.GetSetting(App.Title, "Forms", frm.Name & ".Height", DefaultHeight)
    End If
    If objXMLReg.GetSetting(App.Title, "Forms", frm.Name & ".Maximized", "0") = "1" Then
        frm.WindowState = vbMaximized
    Else
        frm.WindowState = vbNormal
    End If
    'Save Command Text Handle
    If frm.Name = "frmPost" Then
        SaveSetting REGISTRY_KEY, "External", "Link", Format(frm.txtCommand.hwnd)
    End If
    Set objXMLReg = Nothing
End Function

Sub SaveCombo(ByVal objCombo As ComboBox, ByVal strRegSec As String, ByVal strRegKey As String, Optional intStart As Integer = 1)
'Saves Combo itens on Windows Registry
Dim i As Integer, J As Integer, bolClose As Boolean
    If objXMLReg Is Nothing Then
        Set objXMLReg = New XMLRegistry
        objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
        bolClose = True
    End If
    objXMLReg.SaveSetting App.Title, strRegSec, strRegKey & "00", Trim(objCombo.Text)
    If objCombo.ListCount > 0 Then
        For i = intStart To Min(objCombo.ListCount, 29 + intStart)
            If (i - 1) <> objCombo.ListIndex Then
                J = J + 1
                objXMLReg.SaveSetting App.Title, strRegSec, strRegKey & Format(J, "00"), objCombo.List(i - 1)
            End If
        Next
    End If
    If bolClose Then Set objXMLReg = Nothing
End Sub

Public Function Rgb2Html(ByVal RGBColor As Long) As String
Dim strBGR As String, strRGB As String
    strBGR = Right("0000" & Hex(RGBColor), 6)
    strRGB = Right(strBGR, 2) & Mid(strBGR, 3, 2) & Left(strBGR, 2)
    Rgb2Html = "#" & strRGB
End Function

Function ReadINI(Section As String, KeyName As String, FileName As String, Optional Default As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, Default, sRet, Len(sRet), FileName))
End Function

Function WriteINI(Section As String, KeyName As String, NewString As String, FileName As String) As Integer
    Call WritePrivateProfileString(Section, KeyName, NewString, FileName)
End Function

Function ConvertHTMLEntities(ByVal sString As String, ByVal toHTML As Boolean) As String
Dim sChars As String
Dim aHTML(110) As String
Dim i As Integer
    'Verificar par�metros
    sString = Trim(sString)
    
    If Len(sString) = 0 Then
       ConvertHTMLEntities = ""
       Exit Function
    End If
    
    'Preencher a string dos caracteres acentuados
    sChars = "�����������������������ݟ�����������������������������" & _
             "������ߩ���������������������޷��������������������������"
    
    'Preencher o array correspondente
    aHTML(0) = "&Aring;"
    aHTML(1) = "&Aacute;"
    aHTML(2) = "&Agrave;"
    aHTML(3) = "&Acirc;"
    aHTML(4) = "&Auml;"
    aHTML(5) = "&Atilde;"
    aHTML(6) = "&Eacute;"
    aHTML(7) = "&Egrave;"
    aHTML(8) = "&Euml;"
    aHTML(9) = "&Ecirc;"
    aHTML(10) = "&Iacute;"
    aHTML(11) = "&Igrave;"
    aHTML(12) = "&Iuml;"
    aHTML(13) = "&Icirc;"
    aHTML(14) = "&Oacute;"
    aHTML(15) = "&Ograve;"
    aHTML(16) = "&Ouml;"
    aHTML(17) = "&Ocirc;"
    aHTML(18) = "&Otilde;"
    aHTML(19) = "&Uacute;"
    aHTML(20) = "&Ugrave;"
    aHTML(21) = "&Uuml;"
    aHTML(22) = "&Ucirc;"
    aHTML(23) = "&Yacute;"
    aHTML(24) = "&Yuml;"
    aHTML(25) = "&Ccedil;"
    aHTML(26) = "&Ntilde;"
    aHTML(27) = "&aring;"
    aHTML(28) = "&aacute;"
    aHTML(29) = "&agrave;"
    aHTML(30) = "&acirc;"
    aHTML(31) = "&auml;"
    aHTML(32) = "&atilde;"
    aHTML(33) = "&eacute;"
    aHTML(34) = "&egrave;"
    aHTML(35) = "&euml;"
    aHTML(36) = "&ecirc;"
    aHTML(37) = "&iacute;"
    aHTML(38) = "&igrave;"
    aHTML(39) = "&iuml;"
    aHTML(40) = "&icirc;"
    aHTML(41) = "&oacute;"
    aHTML(42) = "&ograve;"
    aHTML(43) = "&ouml;"
    aHTML(44) = "&ocirc;"
    aHTML(45) = "&otilde;"
    aHTML(46) = "&uacute;"
    aHTML(47) = "&ugrave;"
    aHTML(48) = "&uuml;"
    aHTML(49) = "&ucirc;"
    aHTML(50) = "&yacute;"
    aHTML(51) = "&yuml;"
    aHTML(52) = "&ccedil;"
    aHTML(53) = "&ntilde;"
    aHTML(54) = "&ordf;"
    aHTML(55) = "&ordm;"
    aHTML(56) = "&sup2;"
    aHTML(57) = "&sup3;"
    aHTML(58) = "&deg;"
    aHTML(59) = "&sect;"
    aHTML(60) = "&szlig;"
    aHTML(61) = "&copy;"
    aHTML(62) = "&reg;"
    aHTML(63) = "&yen;"
    aHTML(64) = "&euro;"
    aHTML(65) = "&micro;"
    aHTML(66) = "&times;"
    aHTML(67) = "&divide;"
    aHTML(68) = "&para;"
    aHTML(69) = "&frac14;"
    aHTML(70) = "&frac12;"
    aHTML(71) = "&frac34;"
    aHTML(72) = "&lsquo;"
    aHTML(73) = "&rsquo;"
    aHTML(74) = "&laquo;"
    aHTML(75) = "&raquo;"
    aHTML(76) = "&oslash;"
    aHTML(77) = "&Oslash;"
    aHTML(78) = "&aelig;"
    aHTML(79) = "&AElig;"
    aHTML(80) = "&eth;"
    aHTML(81) = "&ETH;"
    aHTML(82) = "&thorn;"
    aHTML(83) = "&THORN;"
    aHTML(84) = "&nbsp;"
    aHTML(85) = "&#8212;"
    aHTML(86) = "&#8211;"
    aHTML(87) = "&#8220;"
    aHTML(88) = "&#8221;"
    aHTML(89) = "&#8230;"
    aHTML(90) = "&iquest;"
    aHTML(91) = "&iexcl;"
    aHTML(92) = "&cent;"
    aHTML(93) = "&pound;"
    aHTML(94) = "&curren;"
    aHTML(95) = "&plusmn;"
    aHTML(96) = "&not;"
    aHTML(97) = "&dagger;"
    aHTML(98) = "&Dagger;"
    aHTML(99) = "&fnof;"
    aHTML(100) = "&permil;"
    aHTML(101) = "&Scaron;"
    aHTML(102) = "&scaron;"
    aHTML(103) = "&trade;"
    aHTML(104) = "&OElig;"
    aHTML(105) = "&oelig;"
    aHTML(106) = "&bull;"
    aHTML(107) = "&macr;"
    aHTML(108) = "&cedil;"
    aHTML(109) = "&sup1;"
    aHTML(110) = "&acute;"

    'Substituir os caracteres acentuados
    For i = 0 To (Len(sChars) - 1)
        If toHTML Then
            If InStr(sString, Mid(sChars, i + 1, 1)) > 0 Then
                sString = Replace(sString, Mid(sChars, i + 1, 1), aHTML(i))
            End If
        Else
            If InStr(sString, aHTML(i)) > 0 Then
                sString = Replace(sString, aHTML(i), Mid(sChars, i + 1, 1))
            End If
        End If
    Next
    'Retornar o valor
    ConvertHTMLEntities = sString
End Function

Function IsArrayEmpty(varArray As Variant) As Boolean
  ' Determines whether an array contains any elements.
  ' Returns False if it does contain elements, True
  ' if it does not.
  Dim lngUBound As Long
 
  On Error Resume Next
  ' If the array is empty, an error occurs when you
  ' check the array's bounds.
  lngUBound = UBound(varArray)
  If Err.Number <> 0 Or lngUBound < 0 Then
    IsArrayEmpty = True
  Else
    IsArrayEmpty = False
  End If
End Function

Public Function FillPost(objPost As xmlStruct) As PostData
On Error GoTo ErrorHandler
Dim strPost As String, strDate As String
Dim strCateg As String, varCateg
Dim udtPost As New PostData
    'Store Post Content
    udtPost.AccountID = gAccount.Current
    udtPost.BlogID = gBlogs(frmPost.CurrentBlog).BlogID
    udtPost.PostID = objPost.Member("postid").Value
    If gAccount.GetPostsMethod = API_METAWEBLOG Or _
       gAccount.GetPostsMethod = API_MT Then
        udtPost.Title = objPost.Member("title").Value
        strCateg = ""
        If gAccount.GetPostsMethod = API_MT Then
            strPost = Replace(objPost.Member("description").Value, vbLf, vbCrLf)
            If gAccount.BodyTag1 <> "" And gAccount.MoreTextTag1 <> "" Then 'Drupal
                udtPost.Text = GetBody(strPost)
                udtPost.More = GetMore(strPost)
                udtPost.Excerpt = GetExcerpt(strPost)
            Else
                udtPost.Text = strPost
            End If
        Else '3.03 - Get metaWeblog categories
            strPost = Replace(objPost.Member("description").Value, vbLf, vbCrLf)
            udtPost.Text = GetBody(strPost)
            udtPost.More = GetMore(strPost)
            udtPost.Excerpt = GetExcerpt(strPost)
            On Error Resume Next
            varCateg = objPost.Member("categories").Value
            If Err.Number <> 0 Then
                Err = 0
                varCateg = objPost.Member("category").Value
            End If
            If Err.Number = 0 Then
                If TypeName(varCateg) = "Variant()" Then
                    If Not IsArrayEmpty(varCateg) Then
                        strCateg = Join(varCateg, vbTab)
                    End If
                End If
            End If
            On Error GoTo ErrorHandler
        End If
        'Inspect MT Extended
        On Error Resume Next
        If gAccount.Extended And Trim(objPost.Member("mt_text_more").Value) <> "" Then
            udtPost.More = Replace(objPost.Member("mt_text_more").Value, vbLf, vbCrLf)
        End If
        If gAccount.Excerpt And Trim(objPost.Member("mt_excerpt").Value) <> "" Then
            udtPost.Excerpt = Replace(objPost.Member("mt_excerpt").Value, vbLf, vbCrLf)
        End If
        If gAccount.Keywords And Trim(objPost.Member("mt_keywords").Value) <> "" Then
            udtPost.Keywords = objPost.Member("mt_keywords").Value
        End If
        If gAccount.TextFilters And Trim(objPost.Member("mt_convert_breaks").Value) <> "" Then
            udtPost.TextFilter = objPost.Member("mt_convert_breaks").Value
        End If
        If gAccount.AllowComments And objPost.Member("mt_allow_comments").Value >= 0 Then
            udtPost.AllowComments = objPost.Member("mt_allow_comments").Value
        End If
        If gAccount.AllowPings And objPost.Member("mt_allow_pings").Value >= 0 Then
            udtPost.AllowPings = objPost.Member("mt_allow_pings").Value
        End If
        If gAccount.TrackBack And Trim(objPost.Member("mt_tb_ping_urls").Value) <> "" Then
            udtPost.TrackBack = objPost.Member("mt_tb_ping_urls").Value
        End If
        On Error GoTo ErrorHandler
        'Convert the Special Characters
        If gAccount.UTF8 Then
            udtPost.Title = UTF8_Decode(udtPost.Title)
            udtPost.Text = UTF8_Decode(udtPost.Text)
            udtPost.More = UTF8_Decode(udtPost.More)
            udtPost.Excerpt = UTF8_Decode(udtPost.Excerpt)
            udtPost.Keywords = UTF8_Decode(udtPost.Keywords)
        End If
        If gSettings.AutoConvert Then
            udtPost.Title = ConvertHTMLEntities(udtPost.Title, False)
            udtPost.Text = ConvertHTMLEntities(udtPost.Text, False)
            udtPost.More = ConvertHTMLEntities(udtPost.More, False)
            udtPost.Excerpt = ConvertHTMLEntities(udtPost.Excerpt, False)
            udtPost.Keywords = ConvertHTMLEntities(udtPost.Keywords, False)
        End If
        If gAccount.UTF8 Then
            udtPost.Categories = UTF8_Decode(strCateg)
        Else
            udtPost.Categories = strCateg
        End If
    ElseIf gAccount.BodyTag1 <> "" Then 'pMachine
        strPost = Replace(objPost.Member("content").Value, vbLf, vbCrLf)
        If gAccount.UTF8 Then
            strPost = UTF8_Decode(strPost)
        End If
        If gSettings.AutoConvert Then
            strPost = ConvertHTMLEntities(strPost, False)
        End If
        udtPost.Title = Replace(Left(strPost, InStr(strPost, gAccount.BodyTag1) - 1), vbCrLf, "")
        udtPost.Text = GetBody(Replace(Replace(Mid(strPost, InStr(strPost, gAccount.BodyTag1)), gAccount.BodyTag1, ""), gAccount.BodyTag2, ""))
        udtPost.More = GetMore(Replace(Replace(Mid(strPost, InStr(strPost, gAccount.BodyTag1)), gAccount.BodyTag1, ""), gAccount.BodyTag2, ""))
    ElseIf gAccount.TitleTag2 <> "" Then
        strPost = Replace(objPost.Member("content").Value, vbLf, vbCrLf)
        If gAccount.UTF8 Then
            strPost = UTF8_Decode(strPost)
        End If
        If gSettings.AutoConvert Then
            strPost = ConvertHTMLEntities(strPost, False)
        End If
        udtPost.Title = GetTitle(strPost)
        udtPost.Categories = GetCateg(strPost)
        udtPost.Text = GetBody(strPost)
        udtPost.More = GetMore(strPost)
        udtPost.Excerpt = GetExcerpt(strPost)
    Else
        strPost = Replace(objPost.Member("content").Value, vbLf, vbCrLf)
        If gAccount.UTF8 Then
            strPost = UTF8_Decode(strPost)
        End If
        If gSettings.AutoConvert Then
            strPost = ConvertHTMLEntities(strPost, False)
        End If
        udtPost.Text = GetBody(strPost)
        udtPost.More = GetMore(strPost)
        udtPost.Excerpt = GetExcerpt(strPost)
    End If
    On Error Resume Next 'New API
    If udtPost.DateTime = CDate(0) Then
        If TypeName(objPost.Member("dateCreated").Value) = "Date" Then
            udtPost.DateTime = objPost.Member("dateCreated").Value
        Else
            strDate = CStr(objPost.Member("dateCreated").Value)
            If Len(strDate) = 17 Then
                udtPost.DateTime = DateSerial(Mid(strDate, 1, 4), Mid(strDate, 5, 2), Mid(strDate, 7, 2)) + _
                                   TimeSerial(Mid(strDate, 10, 2), Mid(strDate, 13, 2), Mid(strDate, 16, 2))
            Else
                udtPost.DateTime = CDate(0)
            End If
        End If
    End If
    udtPost.Author = objPost.Member("authorName").Value
    udtPost.Author = objPost.Member("author").Value
    Set FillPost = udtPost
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetMTCategories(ByVal strPostID As String) As String
Dim strCateg As String, varCateg As Variant, c As Integer
On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    varCateg = GetPostCategories(strPostID)
    If TypeName(varCateg) = "Variant()" Then
        If Not IsArrayEmpty(varCateg) Then
            strCateg = ""
            On Error Resume Next
            For c = 0 To UBound(varCateg)
                If Not varCateg(c).Member("isPrimary").Value Then
                    strCateg = strCateg & Format(varCateg(c).Member("categoryId").Value, CATEG_ID_MASK) & vbTab
                Else
                    strCateg = Format(varCateg(c).Member("categoryId").Value, CATEG_ID_MASK) & vbTab & strCateg
                End If
            Next
            GetMTCategories = strCateg
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Function GetTitle(ByVal strPost As String, _
                  Optional bolUseAccount As Boolean = True) As String
On Error Resume Next
Dim i As Integer, f As Integer
Dim t1 As String, t2 As String
    If bolUseAccount Then
        t1 = gAccount.TitleTag1
        t2 = gAccount.TitleTag2
    Else
        t1 = "<title>"
        t2 = "</title>"
    End If
    i = InStr(strPost, t1) + Len(t1)
    f = InStr(i, strPost, t2)
    GetTitle = Mid(strPost, i, f - i)
End Function

Function GetCateg(ByVal strPost As String, _
                  Optional bolUseAccount As Boolean = True) As String
On Error Resume Next
Dim i As Integer, f As Integer
Dim t1 As String, t2 As String
    If bolUseAccount Then
        t1 = gAccount.CategTag1
        t2 = gAccount.CategTag2
    Else
        t1 = "<category>"
        t2 = "</category>"
    End If
    i = InStr(strPost, t1) + Len(t1)
    f = InStr(i, strPost, t2)
    GetCateg = Mid(strPost, i, f - i)
End Function

Function GetBody(ByVal strPost As String, _
                  Optional bolUseAccount As Boolean = True) As String
On Error Resume Next
Dim i As Integer, f As Integer
Dim t1 As String, t2 As String
Dim c1 As String, c2 As String
Dim m1 As String, m2 As String
Dim b1 As String, b2 As String
    If bolUseAccount Then
        t1 = gAccount.TitleTag1
        t2 = gAccount.TitleTag2
        c1 = gAccount.CategTag1
        c2 = gAccount.CategTag2
        m1 = gAccount.MoreTextTag1
        m2 = gAccount.MoreTextTag2
        b1 = gAccount.BodyTag1
        b2 = gAccount.BodyTag2
    Else
        t1 = "<title>"
        t2 = "</title>"
        c1 = "<category>"
        c2 = "</category>"
    End If
    'Extract Title
    i = InStr(strPost, t1)
    If (t1 <> "" And i > 0) Or t1 = "" Then
        f = InStr(i, strPost, t2)
        If f > 0 Then
            f = f + Len(t2)
            strPost = Left(strPost, i - 1) & Mid(strPost, f)
        End If
    End If
    'Extract Category
    i = InStr(strPost, c1)
    If (c1 <> "" And i > 0) Or c1 = "" Then
        f = InStr(i, strPost, c2)
        If f > 0 Then
            f = f + Len(c2)
            strPost = Left(strPost, i - 1) & Mid(strPost, f)
        End If
    End If
    'Extract More Text
    If (m1 <> "" Or m2 <> "") And bolUseAccount Then
        i = InStr(strPost, m1)
        If i > 0 Then
            If InStr(i, strPost, m2) > i Then
                f = InStr(i, strPost, m2) + Len(m2)
                strPost = Left(strPost, i - 1) & Mid(strPost, f)
            Else
                strPost = Left(strPost, i - 1)
            End If
        End If
    End If
    'Get the body using Account BodyTag's
    If (b1 <> "" And InStr(strPost, b1)) Or (b2 <> "" And InStr(strPost, b2)) Then
        If bolUseAccount Then
            i = InStr(strPost, b1) + Len(b1)
            f = InStr(i, strPost, b2)
            If i = Len(b1) And f = 0 Then
                strPost = ""
            ElseIf f > i Then
                strPost = Mid(strPost, i, f - i)
            Else
                strPost = Mid(strPost, i)
            End If
        End If
    End If
    'Return Post Body
    GetBody = strPost
End Function

Function GetExcerpt(ByVal strPost As String) As String
On Error Resume Next
Dim i As Integer, f As Integer
Dim t1 As String, t2 As String
Dim c1 As String, c2 As String
Dim m1 As String, m2 As String
Dim b1 As String, b2 As String
    t1 = gAccount.TitleTag1
    t2 = gAccount.TitleTag2
    c1 = gAccount.CategTag1
    c2 = gAccount.CategTag2
    m1 = gAccount.MoreTextTag1
    m2 = gAccount.MoreTextTag2
    b1 = gAccount.BodyTag1
    b2 = gAccount.BodyTag2
    If (b1 <> "" And InStr(strPost, b1)) Or (b2 <> "" And InStr(strPost, b2)) Then
        'Extract Title
        i = InStr(strPost, t1)
        If (t1 <> "" And i > 0) Or t1 = "" Then
            f = InStr(i, strPost, t2)
            If f > 0 Then
                f = f + Len(t2)
                strPost = Left(strPost, i - 1) & Mid(strPost, f)
            End If
        End If
        'Extract Category
        i = InStr(strPost, c1)
        If (c1 <> "" And i > 0) Or c1 = "" Then
            f = InStr(i, strPost, c2)
            If f > 0 Then
                f = f + Len(c2)
                strPost = Left(strPost, i - 1) & Mid(strPost, f)
            End If
        End If
        'Extract More Text
        If (m1 <> "" Or m2 <> "") Then
            i = InStr(strPost, m1)
            If i > 0 Then
                If InStr(i, strPost, m2) > i Then
                    f = InStr(i, strPost, m2) + Len(m2)
                    strPost = Left(strPost, i - 1) & Mid(strPost, f)
                Else
                    strPost = Left(strPost, i - 1)
                End If
            End If
        End If
        'Extract Body
        If (b1 <> "" Or b2 <> "") Then
            i = InStr(strPost, b1)
            If i > 0 Then
                If InStr(i, strPost, b2) > i Then
                    f = InStr(i, strPost, b2) + Len(b2)
                    strPost = Left(strPost, i - 1) & Mid(strPost, f)
                Else
                    strPost = Left(strPost, i - 1)
                End If
            End If
        End If
        'Return Excerpt
        GetExcerpt = strPost
    End If
End Function

Function GetMore(ByVal strPost As String) As String
On Error Resume Next
Dim i As Integer, f As Integer
Dim t1 As String, t2 As String
    t1 = gAccount.MoreTextTag1
    t2 = gAccount.MoreTextTag2
    If (t1 <> "" And InStr(strPost, t1)) Or (t2 <> "" And InStr(strPost, t2)) Then
        i = InStr(strPost, t1) + Len(t1)
        f = InStr(i, strPost, t2)
        If i = Len(t1) And f = 0 Then
            GetMore = ""
        ElseIf f > i Then
            GetMore = Mid(strPost, i, f - i)
        Else
            GetMore = Mid(strPost, i)
        End If
    End If
End Function

Sub ActivatePrevInstance(ByVal strCommand As String)
Dim OldTitle As String
Dim PrevHndl As Long
Dim ChildHndl As Long
Dim Result As Long
    
    'Save the title of the application.
    OldTitle = App.Title
    
    'Rename the title of this application so FindWindow
    'will not find this application instance.
    App.Title = "unwanted instance"
    
    'Attempt to get window handle using VB6 class name
    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    
    'Check if found
    If PrevHndl = 0 Then
       'No previous instance found.
       Exit Sub
    End If
    
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    
    'Restore the program.
    Result = OpenIcon(PrevHndl)
    
    'Activate the application.
    Result = SetForegroundWindow(PrevHndl)
    
    If FileExists(strCommand) Then
        ChildHndl = Val(GetSetting(REGISTRY_KEY, "External", "Link"))
        If ChildHndl <> 0 Then
            SendMessage ChildHndl, WM_SETTEXT, 0, ByVal CStr(strCommand)
        End If
    End If
    'End the application.
    End
End Sub

Public Function GetLocaleName(ByVal dwLocaleID As Long) As String

   Dim sReturn As String
   Dim r As Long
   Const LOCALE_SLANGUAGE             As Long = &H2    'localized name of lang

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, LOCALE_SLANGUAGE, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, LOCALE_SLANGUAGE, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetLocaleName = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function

Public Function GetNamePart(ByVal strFullPath As String, Optional ByVal bolExtension As Boolean = True) As String
    ' Comments  : Returns the name [and extension] of a fully qualified file name
    ' Parameters: strFullPath - path and name to parse
    ' Returns   : file name [+ extension]
    '
    On Error GoTo ErrorHandler
    Dim intCounter As Integer
    Dim strTmp As String
    
    ' Parse the string
    For intCounter = Len(strFullPath) To 1 Step -1
        ' It its a slash, grab the sub string
        If Mid$(strFullPath, intCounter, 1) <> "\" Then
            strTmp = Mid$(strFullPath, intCounter, 1) & strTmp
        Else
            Exit For
        End If
    Next intCounter

    ' Return the value
    If bolExtension = False And InStr(strTmp, ".") > 0 Then
        strTmp = Left(strTmp, InStrRev(strTmp, ".") - 1)
    End If
    GetNamePart = strTmp
    Exit Function
    
ErrorHandler:
    GetNamePart = ""
End Function

Public Function GetExtension(ByVal strFullPath As String) As String
    ' Comments  : Returns the extension] of a fully qualified file name
    ' Parameters: strFullPath - path and name to parse
    '
    On Error GoTo ErrorHandler
    Dim intCounter As Integer
    Dim strTmp As String
    
    ' Return the value
    If InStr(strFullPath, ".") > 0 Then
        strTmp = Mid(strFullPath, InStrRev(strFullPath, ".") + 1)
    End If
    GetExtension = strTmp
    Exit Function
    
ErrorHandler:
    GetExtension = ""
End Function

Public Function GetMimeType(ByVal strFullPath As String) As String
    Select Case LCase(GetExtension(strFullPath))
    Case "jpg", "jpeg", "jpe"
        GetMimeType = "image/jpeg"
    Case "gif"
        GetMimeType = "image/gif"
    Case "png"
        GetMimeType = "image/png"
    Case "bmp"
        GetMimeType = "image/bmp"
    Case "tif", "tiff"
        GetMimeType = "image/tiff"
    Case "psd"
        GetMimeType = "image/photoshop"
    Case "txt", "asc"
        GetMimeType = "text/plain"
    Case "rtf"
        GetMimeType = "text/rtf"
    Case "html", "htm"
        GetMimeType = "text/html"
    Case "css"
        GetMimeType = "text/css"
    Case "xml", "xsl"
        GetMimeType = "text/xml"
    Case "mp3", "mp2", "mpga", "mpa"
        GetMimeType = "audio/mpeg"
    Case "m3u"
        GetMimeType = "audio/x-mpegurl"
    Case "mid", "midi", "kar"
        GetMimeType = "audio/midi"
    Case "wav"
        GetMimeType = "audio/x-wav"
    Case "mpg", "mpeg", "mpe"
        GetMimeType = "video/mpeg"
    Case "mov", "qt"
        GetMimeType = "video/quicktime"
    Case "avi"
        GetMimeType = "video/x-msvideo"
    Case Else
        GetMimeType = ""
    End Select
End Function

Public Function GetIeVer() As String
Dim VersionInfo As DllVersionInfo
    VersionInfo.cbSize = Len(VersionInfo)
    
    Call DllGetVersion(VersionInfo)
    GetIeVer = VersionInfo.dwMajorVersion & "." & VersionInfo.dwMinorVersion
End Function

Public Function GetOSVer() As String
    Dim osv As OSVERSIONINFOEX
    osv.dwOSVersionInfoSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.dwPlatformId
        Case Is = VER_PLATFORM_WIN32s
            GetOSVer = "Windows 3.x"
        Case Is = VER_PLATFORM_WIN32_WINDOWS
            Select Case osv.dwMinorVersion
            Case Is = 0
                If InStr(UCase(osv.szCSDVersion), "C") Then
                    GetOSVer = "Windows 95 OSR2"
                Else
                    GetOSVer = "Windows 95"
                End If
            Case Is = 10
                If InStr(UCase(osv.szCSDVersion), "A") Then
                    GetOSVer = "Windows 98 SE"
                Else
                    GetOSVer = "Windows 98"
                End If
            Case Is = 90
                GetOSVer = "Windows Me"
            End Select
        Case Is = VER_PLATFORM_WIN32_NT
            Select Case osv.dwMajorVersion
            Case Is = 3
                Select Case osv.dwMinorVersion
                Case Is = 0: GetOSVer = "Windows NT 3"
                Case Is = 1: GetOSVer = "Windows NT 3.1"
                Case Is = 5: GetOSVer = "Windows NT 3.5"
                Case Is = 51: GetOSVer = "Windows NT 3.51"
                End Select
            Case Is = 4
                GetOSVer = "Windows NT 4"
            Case Is = 5
                Select Case osv.dwMinorVersion
                Case Is = 0 'win 2000
                    Select Case osv.wProductType
                    Case Is = VER_NT_WORKSTATION
                        GetOSVer = "Windows 2000 Professional"
                    Case Is = VER_NT_SERVER
                        Select Case osv.wSuiteMask
                        Case Is = VER_SUITE_DATACENTER
                            GetOSVer = "Windows 2000 DataCenter Server"
                        Case Is = VER_SUITE_ENTERPRISE
                            GetOSVer = "Windows 2000 Advanced Server"
                        Case Else
                            GetOSVer = "Windows 2000 Server"
                        End Select
                    End Select
                Case Is = 1 'win XP or win 2003 server
                    Select Case osv.wProductType
                    Case Is = VER_NT_WORKSTATION 'win XP
                        If osv.wSuiteMask = VER_SUITE_PERSONAL Then
                            GetOSVer = "Windows XP Home Edition"
                        Else
                            GetOSVer = "Windows XP Professional"
                        End If
                    Case Else
                        If osv.wSuiteMask = VER_SUITE_ENTERPRISE Then
                            GetOSVer = "Windows 2003 Enterprise Server"
                        Else
                            GetOSVer = "Windows 2003 Server"
                        End If
                    End Select
                End Select
            End Select
        End Select
    End If
End Function

Private Function InvokeParser(strInvokation As String) As Boolean
    Dim objDummy As Object
    On Error Resume Next
    Set objDummy = CreateObject(strInvokation)
    If Err.Number = 0 Then
        InvokeParser = True
    Else
        InvokeParser = False
    End If
    Set objDummy = Nothing
End Function

Public Function GetXMLVer() As String
    ' Check latest versions first

    If InvokeParser("Msxml2.DOMDocument.4.0") = True Then
        GetXMLVer = "4.0"
        Exit Function
    End If
    
    If InvokeParser("Msxml2.DOMDocument.3.0") = True Then
        GetXMLVer = "3.0"
        Exit Function
    End If
    
    If InvokeParser("Msxml2.DOMDocument.2.6") = True Then
        GetXMLVer = "2.6"
        Exit Function
    End If

    If InvokeParser("Msxml.DOMDocument") = True Then
        GetXMLVer = "2.0"
        Exit Function
    End If
    
    ' No XML Parser detected
    GetXMLVer = "No XML Parser detected"
End Function

Public Function ReplaceMediaInfo(ByVal strText As String) As String
On Error Resume Next
Dim strTitle As String, strAuthor As String
    strTitle = MediaPlayerInfo("Title")
    If Trim(strTitle) <> "" Then
        strText = Replace(strText, "%T%", strTitle)
        strAuthor = MediaPlayerInfo("Author")
        If gBlog.MediaLink And Trim(strAuthor) <> "" Then
            strAuthor = "<a href=""http://www.windowsmedia.com/mg/search.asp?srch=" & Replace(strAuthor, " ", "+") & """>" & strAuthor & "</a>"
        End If
        strText = Replace(strText, "%A%", strAuthor)
        strText = Replace(strText, "%B%", MediaPlayerInfo("Album"))
        strText = Replace(strText, "%D%", MediaPlayerInfo("DurationString"))
    End If
    ReplaceMediaInfo = strText
End Function

Public Function IsCompiled() As Boolean
On Error GoTo ErrorHandler
    'in compiledmode the next line is not
    'available, so no error occurs !
    Debug.Print 1 / 0
    IsCompiled = True
    Exit Function
ErrorHandler:
    IsCompiled = False
End Function

Public Sub AutoSelect(txtTextBox As Control)
    If Len(txtTextBox.Text) > 0 Then
        txtTextBox.SelStart = 0
        txtTextBox.SelLength = Len(txtTextBox.Text)
    End If
End Sub

Public Function OnlyNumbers(ByVal KeyAscii) As Integer
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        OnlyNumbers = 0
    Else
        OnlyNumbers = KeyAscii
    End If
End Function

'Public Sub WriteDebugInfo(ByVal strInfo As String)
'On Error Resume Next
'    WritePrivateProfileString App.Title, Format(Now), strInfo, gAppDataPath & "\debug.log"
'End Sub

Public Sub ErrorMessage(ByVal errNumber As Long, ByVal errDescription As String, ByVal errWhere As String)
    MsgBox errDescription, vbCritical, App.Title & ": " & errNumber
    WritePrivateProfileString errWhere, Format(Now), "Erro: " & errNumber & " - " & errDescription, gAppDataPath & "\error.log"
End Sub

Function SafeFileName(strProposed As String) As String
  Dim bad() As String, i As Integer
  bad = Split(">,<,&,/,\,:,.,|,?,*,""", ",")
  For i = 0 To UBound(bad)
    strProposed = Replace(strProposed, bad(i), "")
  Next
  SafeFileName = strProposed
End Function
