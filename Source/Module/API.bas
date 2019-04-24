Attribute VB_Name = "basAPI"
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
Public Const APPKEY = "E7F0F69AB4125D811D74C797B11A4DC768E9B2BC"
Public Const CATEG_ID_MASK = "00000000"

Public Enum SupportedCMS
    CMS_CUSTOM = 0
    CMS_BLOGGER = 1
    CMS_BLOGGERPRO = 2
    CMS_MT = 3
    CMS_B2 = 4
    CMS_BBT = 5
    CMS_NUCLEUS = 6
    CMS_BLOGALIA = 7
    CMS_BWXML = 8
    CMS_DRUPAL = 9
    CMS_XOOPS = 10
End Enum

Public Enum SupportedAPI
    API_NOTSUPPORTED = 0
    API_BLOGGER = 1
    API_BLOGGER2 = 2
    API_METAWEBLOG = 3
    API_MT = 4
    API_B2 = 5
End Enum

Public Function LoadBlogs(ByVal XMLCache As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strMethod As String
Dim varStruct(), b As Integer
Dim objBlog As xmlStruct
Dim strBlogName As String

    frmPost.Message = GetMsg(msgGettingBlogs)
    Screen.MousePointer = vbHourglass
    DoEvents
    strMethod = "blogger.getUsersBlogs"
    Set objClient = GetXMLClient()
    If XMLCache And FileExists(gAppDataPath & "\blogs" & Format(gAccount.Current, "00") & ".xml") Then
        Set DOMDocument = New DOMDocument
        DOMDocument.Load gAppDataPath & "\blogs" & Format(gAccount.Current, "00") & ".xml"
    Else
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                            APPKEY, gAccount.User, gAccount.Password)
    End If
    objClient.ResponseToVariant DOMDocument, varStruct
    frmPost.acbMain.Bands("bndTools").Tools("miBlogs").CBList.Clear
    ReDim gBlogs(0)
    For b = 0 To UBound(varStruct)
        Set objBlog = varStruct(b)
        strBlogName = objBlog.Member("blogName").Value
        If gAccount.UTF8 Then
            frmPost.acbMain.Bands("bndTools").Tools("miBlogs").CBList.AddItem UTF8_Decode(strBlogName)
        Else
            frmPost.acbMain.Bands("bndTools").Tools("miBlogs").CBList.AddItem strBlogName
        End If
        ReDim Preserve gBlogs(b)
        gBlogs(b).URL = objBlog.Member("url").Value
        gBlogs(b).BlogID = objBlog.Member("blogid").Value
        gBlogs(b).Name = objBlog.Member("blogName").Value
        On Error Resume Next 'New API
        gBlogs(b).IsAdmin = objBlog.Member("isAdmin").Value
        If Err <> 0 Then gBlogs(b).IsAdmin = True
        On Error GoTo ErrorHandler
    Next
    If frmPost.acbMain.Bands("bndTools").Tools("miBlogs").CBList.Count > 0 Then
        'Save the blogs.xml if the blog list was downloaded
        If Not XMLCache Then
            DOMDocument.Save gAppDataPath & "\blogs" & Format(gAccount.Current, "00") & ".xml"
        End If
        'Get last Blog used
        Set objXMLReg = New XMLRegistry
        objXMLReg.OpenXMLFile gAppDataPath & XML_SETTINGS
        b = Val(objXMLReg.GetSetting(App.Title, "Accounts/a" & Format(gAccount.Current, "00"), "Blog", "0"))
        'Verify if it still exists then set the toolbar combo
        If b >= frmPost.acbMain.Bands("bndTools").Tools("miBlogs").CBList.Count Then
            If FileExists(gAppDataPath & "\categs.xml") Then
                Kill gAppDataPath & "\categs.xml"
            End If
            frmPost.CurrentBlog = 0
            objXMLReg.SaveSetting App.Title, "Accounts/a" & Format(gAccount.Current, "00"), "Blog", "0"
        Else
            frmPost.CurrentBlog = b
        End If
        Set objXMLReg = Nothing
    End If
    LoadBlogs = LoadTextFilters(XMLCache)
ExitNow:
    Set objBlog = Nothing
    Set DOMDocument = Nothing
    Set objClient = Nothing
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    LoadBlogs = False
    ErrorMessage Err.Number, Err.Description, "LoadBlogs"
    Resume ExitNow
End Function

Public Function Post(ByVal strTitle As String, _
                     ByVal strPost As String, _
                     ByVal strMore As String, _
                     ByVal strExcerpt As String, _
                     ByVal strKeywords As String, _
                     ByRef varCategs As Variant, _
                     ByVal strBlogID As String, _
                     ByVal strPostID As String, _
                     ByVal bolPublish As Boolean, _
                     ByVal bolSilent As Boolean, _
                     Optional ByVal intAllowComments As Integer = -1, _
                     Optional ByVal intAllowPings As Integer = -1, _
                     Optional ByVal datDateTime As Date, _
                     Optional ByVal strTrackBack As String, _
                     Optional ByVal strTextFilter As String) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim objLogin As xmlStruct
Dim objPost As xmlStruct
Dim objOptions As xmlStruct
Dim objActions As xmlStruct
Dim DOMDocument As DOMDocument
Dim strMethod As String
Dim varResponse
    Screen.MousePointer = vbHourglass
    DoEvents
    Set objClient = GetXMLClient()
    If gAccount.PostMethod = API_METAWEBLOG Or _
       gAccount.PostMethod = API_MT Then
        'Process Text
        If gSettings.AutoConvert Then 'Convers�o HTML
            strTitle = ConvertHTMLEntities(strTitle, True)
            strPost = ConvertHTMLEntities(strPost, True)
            strMore = ConvertHTMLEntities(strMore, True)
            strExcerpt = ConvertHTMLEntities(strExcerpt, True)
            strKeywords = ConvertHTMLEntities(strKeywords, True)
        End If
        If gAccount.UTF8 Or gAccount.UTF8OnPost Then 'Convers�o UTF-8
            strTitle = UTF8_Encode(strTitle)
            strPost = UTF8_Encode(strPost)
            strMore = UTF8_Encode(strMore)
            strExcerpt = UTF8_Encode(strExcerpt)
            strKeywords = UTF8_Encode(strKeywords)
        End If
        'Create the Struct
        Set objPost = New xmlStruct
        objPost.Add "title", strTitle
        objPost.Add "description", strPost
        If datDateTime > CDate(0) Then
            objPost.Add "pubDate", datDateTime
        End If
        If datDateTime > CDate(0) Then
            objPost.Add "dateCreated", datDateTime
        End If
        If gAccount.Extended And (Trim(strMore) <> "" Or strPostID <> "") Then
            objPost.Add "mt_text_more", strMore
        End If
        If gAccount.Excerpt And (Trim(strExcerpt) <> "" Or strPostID <> "") Then
            objPost.Add "mt_excerpt", strExcerpt
        End If
        If gAccount.TrackBack And (Trim(strTrackBack) <> "" Or strPostID <> "") Then
            objPost.Add "mt_tb_ping_urls", strTrackBack
        End If
        If gAccount.Keywords And (Trim(strKeywords) <> "" Or strPostID <> "") Then
            objPost.Add "mt_keywords", strKeywords
        End If
        If gAccount.TextFilters And (Trim(strTextFilter) <> "" Or strPostID <> "") Then
            objPost.Add "mt_convert_breaks", strTextFilter
        End If
        If gAccount.AllowPings And intAllowPings >= 0 Then
            objPost.Add "mt_allow_pings", intAllowPings
        End If
        If gAccount.PostMethod = API_MT And gAccount.GetCategMethod = API_MT Then
            If gAccount.AllowComments And intAllowComments >= 0 Then
                objPost.Add "mt_allow_comments", intAllowComments
            End If
        Else
            If gAccount.AllowComments And intAllowComments >= 0 Then
                objPost.Add "mt_allow_comments", (intAllowComments = 1)
            End If
            If Not IsArrayEmpty(varCategs) Then
                objPost.Add "categories", varCategs
            ElseIf strPostID <> "" Then
                objPost.Add "categories", Array("")
            End If
        End If
        'Post Content
        If strPostID = "" Then
            'Create the post
            strMethod = "metaWeblog.newPost"
            'If there are Categories on new MT post: don't publish now!
            If gAccount.PostMethod = API_MT And gAccount.GetCategMethod = API_MT And Not IsArrayEmpty(varCategs) Then
                Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                                    strBlogID, gAccount.User, gAccount.Password, _
                                                    objPost, False)
            Else
                Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                                    strBlogID, gAccount.User, gAccount.Password, _
                                                    objPost, bolPublish)
            End If
        Else
            'Set Categories of the posts (only MT)
            If gAccount.PostMethod = API_MT And gAccount.GetCategMethod = API_MT Then
                SetPostCategories strPostID, varCategs
            End If
            'Edit the post
            strMethod = "metaWeblog.editPost"
            Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                                strPostID, gAccount.User, gAccount.Password, _
                                                objPost, bolPublish)
        End If
    ElseIf gAccount.PostMethod = API_BLOGGER2 Then
        'Process Text
        If gSettings.AutoConvert Then
            strTitle = ConvertHTMLEntities(strTitle, True)
            strPost = ConvertHTMLEntities(strPost, True)
        End If
        If gAccount.UTF8 Or gAccount.UTF8OnPost Then
            strTitle = UTF8_Encode(strTitle)
            strPost = UTF8_Encode(strPost)
        End If
        'Create the Login Struct
        Set objLogin = New xmlStruct
        objLogin.Add "username", gAccount.User
        objLogin.Add "password", gAccount.Password
        objLogin.Add "appkey", APPKEY
        'Create the Post Struct
        Set objPost = New xmlStruct
        objPost.Add "blogID", strBlogID
        objPost.Add "title", strTitle 'as implemented (!)
        objPost.Add "body", strPost
        'Create the postOptions Struct
        Set objOptions = New xmlStruct
        objOptions.Add "title", strTitle 'as defined (!)
        If Not IsArrayEmpty(varCategs) Then
            objOptions.Add "categories", varCategs
        End If
        objPost.Add "postOptions", objOptions
        'Create the Actions Struct
        Set objActions = New xmlStruct
        objActions.Add "doPublish", bolPublish
        'Select the Method
        If strPostID <> "" Then
            objPost.Add "postID", strPostID
            strMethod = "blogger2.editPost"
        Else
            strMethod = "blogger2.newPost"
        End If
        'Post Text
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                             objLogin, objPost, objActions)
    Else
        'Process Text
        If Trim(gAccount.TitleTag2) <> "" Then
            strPost = gAccount.TitleTag1 & strTitle & gAccount.TitleTag2 & strPost
        End If
        If Not IsArrayEmpty(varCategs) Then
            If gAccount.MultiCategory Then
                strPost = gAccount.CategTag1 & Join(varCategs, ",") & gAccount.CategTag2 & strPost
            Else
                strPost = gAccount.CategTag1 & varCategs(0) & gAccount.CategTag2 & strPost
            End If
        End If
        If Trim(strMore) <> "" And (Trim(gAccount.MoreTextTag1) <> "" Or Trim(gAccount.MoreTextTag2) <> "") Then
            strPost = strPost & gAccount.MoreTextTag1 & strMore & gAccount.MoreTextTag2
        End If
        If gSettings.AutoConvert Then strPost = ConvertHTMLEntities(strPost, True)
        If gAccount.UTF8 Or gAccount.UTF8OnPost Then strPost = UTF8_Encode(strPost)
        'Post Text
        If strPostID = "" Then
            strMethod = "blogger.newPost"
            Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                                APPKEY, strBlogID, gAccount.User, gAccount.Password, _
                                                strPost, bolPublish)
        Else
            strMethod = "blogger.editPost"
            Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                                APPKEY, strPostID, gAccount.User, gAccount.Password, _
                                                strPost, bolPublish)
        End If
    End If
    'Process Server Response
    objClient.ResponseToVariant DOMDocument, varResponse
    If strPostID = "" Then  'New Post
        gPostID = varResponse
        strPostID = gPostID
        'Set Categories of the posts (only MT)
        If gAccount.PostMethod = API_MT And gAccount.GetCategMethod = API_MT And Not IsArrayEmpty(varCategs) Then
            SetPostCategories strPostID, varCategs
            'If was published then republish to show the categories correctly
            If bolPublish Then
                'Republish the post
                strMethod = "mt.publishPost"
                Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                                    strPostID, gAccount.User, gAccount.Password)
            End If
        End If
        If gAccount.PingWeblogs > 0 And bolPublish Then
            WeblogsPing ReadINI("Ping-" & Format(gAccount.PingWeblogs, "00"), "Host", App.Path & "\CMS\Ping.ini"), _
                        ReadINI("Ping-" & Format(gAccount.PingWeblogs, "00"), "Page", App.Path & "\CMS\Ping.ini")
        End If
        If bolSilent Then
            frmPost.Message = GetMsg(msgPostSuccess)
        Else
            MsgBox GetMsg(msgPostSuccess) & varResponse, vbInformation
        End If
    Else 'Edit Post
        If CBool(varResponse) Then
            If bolSilent Then
                frmPost.Message = GetMsg(msgUpdateSuccess)
            Else
                MsgBox GetMsg(msgUpdateSuccess), vbInformation
            End If
        Else
            MsgBox GetMsg(msgPostError) & varResponse, vbInformation
            GoTo ExitNow
        End If
    End If
    'Post OK
    Post = True
ExitNow:
    Set DOMDocument = Nothing
    Set objPost = Nothing
    Set objLogin = Nothing
    Set objActions = Nothing
    Set objClient = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, "Post"
    Resume ExitNow
End Function

Public Function DeletePost(ByVal PostID As String, ByVal bolPublish As Boolean, Optional ByVal bolSilent As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strMethod As String
Dim varResponse
    Screen.MousePointer = vbHourglass
    If bolPublish Then
        frmPost.Message = GetMsg(msgDPublishing)
    Else
        frmPost.Message = GetMsg(msgDeleting)
    End If
    DoEvents
    Set objClient = GetXMLClient()
    strMethod = "blogger.deletePost"
    Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                        APPKEY, PostID, gAccount.User, gAccount.Password, _
                                        bolPublish)
    objClient.ResponseToVariant DOMDocument, varResponse
    If varResponse Then
        If bolSilent Then
            frmPost.Message = GetMsg(msgDelSuccess)
        Else
            MsgBox GetMsg(msgDelSuccess), vbInformation
        End If
        DeletePost = True
    Else
        MsgBox GetMsg(msgDelError) & varResponse, vbInformation
        DeletePost = False
        GoTo ExitNow
    End If
    frmPost.NewPost
    'frmPost.SaveDraftPost
ExitNow:
    Set DOMDocument = Nothing
    Set objClient = Nothing
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    DeletePost = False
    ErrorMessage Err.Number, Err.Description, "DeletePost"
    Resume ExitNow
End Function

Public Function GetRecentPosts(ByVal intPosts As Integer)
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim objLogin As xmlStruct
Dim objFilter As xmlStruct
Dim DOMDocument As DOMDocument
Dim varStruct As Variant
Dim strMethod As String, strBlogID As String
    strBlogID = gBlogs(frmPost.CurrentBlog).BlogID
    Set objClient = GetXMLClient()
    If Not IsCompiled() Then
        If FileExists(gAppDataPath & "\debug.xml") Then
            Set DOMDocument = New DOMDocument
            DOMDocument.Load gAppDataPath & "\debug.xml"
            GoTo DebugJump:
        End If
    End If
    Select Case gAccount.GetPostsMethod
    Case API_METAWEBLOG, API_MT
        strMethod = "metaWeblog.getRecentPosts"
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                            strBlogID, gAccount.User, _
                                            gAccount.Password, intPosts)
'    Case API_BLOGGER2 + 10 'To Wait the correct implementation
'        strMethod = "blogger2.getPosts"
'        'Create the Login Struct
'        Set objLogin = New xmlStruct
'        objLogin.Add "username", gAccount.User
'        objLogin.Add "password", gAccount.Password
'        objLogin.Add "appkey", APPKEY
'        'Create the Filters Struct
'        Set objFilter = New xmlStruct
'        objFilter.Add "numOfPosts", intPosts
'        'objFilter.Add "onDate", Date - 2
'        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
'                                            objLogin, strBlogID, objFilter)
    Case Else
        strMethod = "blogger.getRecentPosts"
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                            APPKEY, strBlogID, gAccount.User, _
                                            gAccount.Password, intPosts)
    End Select
DebugJump:
    objClient.ResponseToVariant DOMDocument, varStruct
    If UBound(varStruct) > 0 Then DOMDocument.Save gAppDataPath & "\posts.xml"
    GetRecentPosts = varStruct
    Set DOMDocument = Nothing
    Set objClient = Nothing
    Set objLogin = Nothing
    Set objFilter = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function GetPost(Optional ByVal strPostID As String, Optional ByVal bolEdit = True) As PostData
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim udtPost As New PostData
Dim strMethod As String, strBlogID As String
Dim varCateg, c As Integer
Dim objLogin As New xmlStruct
Dim objPost As New xmlStruct
    strBlogID = gBlogs(frmPost.CurrentBlog).BlogID
    If strPostID = "" Then strPostID = InputBox(GetMsg(msgEnterPostID))
    If strPostID = "" Then
        Exit Function
    ElseIf InStr(Trim(strPostID), " ") <> 0 Then
        MsgBox GetMsg(msgInvalidPostID), vbExclamation
        Exit Function
    End If
    If bolEdit Then
        frmPost.Message = GetMsg(msgGettingPost)
        Screen.MousePointer = vbHourglass
        DoEvents
    End If
    Set objClient = GetXMLClient()
    If gAccount.GetPostsMethod = API_METAWEBLOG Or _
       gAccount.GetPostsMethod = API_MT Then
        strMethod = "metaWeblog.getPost"
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                            strPostID, gAccount.User, gAccount.Password)
        objClient.ResponseToVariant DOMDocument, objPost
    
    ElseIf gAccount.GetPostsMethod = API_BLOGGER2 Then
        'Create the Login Struct
        Set objLogin = New xmlStruct
        objLogin.Add "username", gAccount.User
        objLogin.Add "password", gAccount.Password
        objLogin.Add "appkey", APPKEY
        'Create the Post Struct
        Set objPost = New xmlStruct
        objPost.Add "blogID", strBlogID
        objPost.Add "postID", strPostID
        'Call the Method
        strMethod = "blogger2.getPost"
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                            objLogin, objPost)
        objClient.ResponseToVariant DOMDocument, objPost
    Else
        strMethod = "blogger.getPost"
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                            APPKEY, strPostID, gAccount.User, gAccount.Password)
        objClient.ResponseToVariant DOMDocument, objPost
    End If
    Set udtPost = FillPost(objPost)
    If gAccount.GetPostsMethod = API_MT Then
        udtPost.Categories = GetMTCategories(udtPost.PostID)
    End If

    Set GetPost = udtPost
    If bolEdit Then
        frmPost.ClearUndo
        frmPost.EditPost udtPost
    End If
ExitNow:
    Set objPost = Nothing
    Set objLogin = Nothing
    Set DOMDocument = Nothing
    Set objClient = Nothing
    If bolEdit Then
        frmPost.Message = ""
        Screen.MousePointer = vbDefault
    End If
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, "GetPost"
    Resume ExitNow
End Function

Public Function GetPostCategories(ByVal strPostID As String)
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim varStruct As Variant
Dim strMethod As String
    Set objClient = GetXMLClient()
    strMethod = "mt.getPostCategories"
    Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                        strPostID, gAccount.User, gAccount.Password)
    objClient.ResponseToVariant DOMDocument, varStruct
    GetPostCategories = varStruct
    Set DOMDocument = Nothing
    Set objClient = Nothing
    Exit Function
ErrorHandler:
    If InStr(Err.Description, "Can't call method") <> 0 Then
        Err.Raise Err.Number, Err.Source, "The Post " & strPostID & " has Secondary Categories but no Primary Category!"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

Public Function SetPostCategories(ByVal strPostID As String, _
                                  ByVal varCategs As Variant) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim objStruct As xmlStruct
Dim DOMDocument As DOMDocument
Dim varResponse As Variant
Dim strMethod As String
    If IsArrayEmpty(varCategs) Then
        Set objStruct = New xmlStruct
        objStruct.Add "categoryId", -1
        varCategs = Array(objStruct)
    End If
    'Send Array
    Set objClient = GetXMLClient()
    strMethod = "mt.setPostCategories"
    Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                        strPostID, gAccount.User, gAccount.Password, varCategs)
    objClient.ResponseToVariant DOMDocument, varResponse
    SetPostCategories = varResponse
    Set DOMDocument = Nothing
    Set objClient = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SaveTemplate(ByVal strTemplate As String, _
                             ByVal strBlogID As String, _
                             ByVal bolPublish As Boolean, _
                             Optional ByVal bolSilent As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strMethod As String
Dim varResponse
    Screen.MousePointer = vbHourglass
    DoEvents
    Set objClient = GetXMLClient()
    'Process Text
    If gSettings.AutoConvert Then
        strTemplate = ConvertHTMLEntities(strTemplate, True)
    End If
    If gAccount.UTF8 Then
        strTemplate = UTF8_Encode(strTemplate)
    End If
    'Save the Template
    strMethod = "blogger.setTemplate"
    Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                        APPKEY, strBlogID, gAccount.User, gAccount.Password, _
                                        strTemplate, gPostID)
    objClient.ResponseToVariant DOMDocument, varResponse
    If varResponse Then
        If bolPublish And (gAccount.CMS = CMS_BLOGGER Or gAccount.CMS = CMS_BLOGGERPRO) Then
            If Not Publish(strBlogID) Then
                MsgBox GetMsg(msgTplSavedNotPub), vbInformation
                GoTo ExitNow
            End If
        End If
        If bolSilent Then
            frmPost.Message = GetMsg(msgUpdateSuccess)
        Else
            MsgBox GetMsg(msgUpdateSuccess), vbInformation
        End If
    Else
        MsgBox GetMsg(msgPostError) & varResponse, vbInformation
        GoTo ExitNow
    End If
    'SaveTemplate OK
    SaveTemplate = True
ExitNow:
    Set DOMDocument = Nothing
    Set objClient = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, "SaveTemplate"
    Resume ExitNow
End Function

Public Sub GetTemplate(ByVal strType As String)
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strMethod As String, strBlogID As String
Dim strTemplate As String
Dim varResponse
    frmPost.Message = GetMsg(msgGettingTemplate)
    Screen.MousePointer = vbHourglass
    DoEvents
    strMethod = "blogger.getTemplate"
    strBlogID = gBlogs(frmPost.CurrentBlog).BlogID
    
    frmPost.ClearUndo
    
    Set objClient = GetXMLClient()
    Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                        APPKEY, strBlogID, gAccount.User, _
                                        gAccount.Password, strType)
    objClient.ResponseToVariant DOMDocument, varResponse
    strTemplate = Replace(varResponse, vbLf, vbCrLf)
    If gAccount.UTF8 Then
        strTemplate = UTF8_Decode(strTemplate)
    End If
    If gSettings.AutoConvert Then
        strTemplate = ConvertHTMLEntities(strTemplate, False)
    End If
    frmPost.txtPost.Text = strTemplate
    gPostID = strType
    Select Case strType
    Case "main"
        frmPost.imgStatus.Picture = frmPost.acbMain.Tools("miTemplate").GetPicture(0)
        frmPost.lblStatus.Caption = GetMsg(msgMainTemplate)
    Case "archiveIndex"
        frmPost.imgStatus.Picture = frmPost.acbMain.Tools("miTemplate").GetPicture(0)
        frmPost.lblStatus.Caption = GetMsg(msgArchiveTemplate)
    End Select
    If gSettings.ColorizeCode Then frmPost.txtPost.Colorize
    frmPost.TemplateMode True
    frmPost.tabPost.CurrTab = TAB_EDITOR
    frmPost.PostData.AccountID = gAccount.Current
    frmPost.PostData.BlogID = strBlogID
    frmPost.PostData.PostID = gPostID
    frmPost.Changed = False
ExitNow:
    Set DOMDocument = Nothing
    Set objClient = Nothing
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, "GetTemplate"
    Resume ExitNow
End Sub

Public Function LoadCategories(ByVal XMLCache As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strMethod As String, strBlogID As String
Dim varStruct, b As Integer, t As Integer
Dim objCateg As xmlStruct, strCategName As String
Dim arrCateg() As String

    frmPost.cboPostCat.Clear
    Select Case gAccount.GetCategMethod
    Case API_METAWEBLOG
        strMethod = "metaWeblog.getCategories"
    Case API_MT
        strMethod = "mt.getCategoryList"
    Case API_B2
        strMethod = "b2.getCategories"
    Case Else
        'Categories not supported
        With frmPost
        .cboPostCat.Visible = False
        .cmdCategories.Visible = False
        .txtPostTit.Visible = SupportsTitle()
        
        .pnlEditor.Grid(gsColWidth, 2) = 15
        .pnlEditor.Grid(gsColWidth, 3) = 0
        .pnlEditor.Grid(gsColWidth, 4) = 0
        .pnlEditor.Grid(gsRowHeight, 0) = IIf(SupportsTitle(), 375, 15)
        End With
        If FileExists(gAppDataPath & "\categs.xml") Then
            Kill gAppDataPath & "\categs.xml"
        End If
        Exit Function
    End Select
    frmPost.Message = GetMsg(msgGettingBlogCategs)
    Screen.MousePointer = vbHourglass
    'DoEvents
    strBlogID = gBlogs(frmPost.CurrentBlog).BlogID
    Set objClient = GetXMLClient()
    If XMLCache And FileExists(gAppDataPath & "\categs.xml") Then
        Set DOMDocument = New DOMDocument
        DOMDocument.Load gAppDataPath & "\categs.xml"
    Else
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                            strBlogID, gAccount.User, gAccount.Password)
        DOMDocument.Save gAppDataPath & "\categs.xml"
    End If
    objClient.ResponseToVariant DOMDocument, varStruct
    If TypeName(varStruct) = "xmlStruct" Then
        For b = 1 To varStruct.Count
            frmPost.cboPostCat.AddItem varStruct.Member(b).Name
        Next
    Else
        If Not IsArrayEmpty(varStruct) Then
            For b = 0 To UBound(varStruct)
                Set objCateg = varStruct(b)
                Select Case gAccount.GetCategMethod
                Case API_MT, API_B2
                    With frmPost.cboPostCat
                        If gAccount.UTF8 Then
                            .AddItem UTF8_Decode(objCateg.Member("categoryName").Value)
                        Else
                            .AddItem objCateg.Member("categoryName").Value
                        End If
                        .ItemData(.NewIndex) = objCateg.Member("categoryID").Value
                    End With
                Case Else
                    On Error Resume Next
                    strCategName = ""
                    strCategName = objCateg.Member("description").Value
                    strCategName = objCateg.Member("title").Value
                    If strCategName <> "" Then
                        If gAccount.UTF8 Then
                            frmPost.cboPostCat.AddItem UTF8_Decode(strCategName)
                        Else
                            frmPost.cboPostCat.AddItem strCategName
                        End If
                    End If
                End Select
            Next
        End If
    End If
    With frmPost
        .txtPostTit.Visible = True
        .cboPostCat.Visible = True
        .cmdCategories.Visible = gAccount.MultiCategory
        .pnlEditor.Grid(gsRowHeight, 0) = 375
        .pnlEditor.Grid(gsColWidth, 2) = 1065
        .pnlEditor.Grid(gsColWidth, 3) = 1635
        .pnlEditor.Grid(gsColWidth, 4) = IIf(gAccount.MultiCategory, 390, 0)
        .pnlEditor.Refresh
    End With
    frmPost.cboPostCat.AddItem "(" & LCase(GetLbl(lblNone)) & ")", 0
    frmPost.cboPostCat.AddItem "(" & LCase(GetLbl(lblReload)) & ")", frmPost.cboPostCat.ListCount
    frmPost.cboPostCat.ListIndex = 0
    LoadCategories = True
ExitNow:
    Set objCateg = Nothing
    Set DOMDocument = Nothing
    Set objClient = Nothing
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    LoadCategories = False
    ErrorMessage Err.Number, Err.Description, "LoadCategories"
    Resume ExitNow
End Function

Public Function LoadTextFilters(ByVal XMLCache As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strMethod As String
Dim varStruct, b As Integer, t As Integer
Dim objStruct As xmlStruct

    Set frmPost.TextFilters = New Collection
    Select Case gAccount.GetPostsMethod
    Case API_MT
        strMethod = "mt.supportedTextFilters"
    Case Else
        'TextFilters not supported
        LoadTextFilters = True
        If FileExists(gAppDataPath & "\filters.xml") Then
            Kill gAppDataPath & "\filters.xml"
        End If
        Exit Function
    End Select
    Set objClient = GetXMLClient()
    If XMLCache And FileExists(gAppDataPath & "\filters.xml") Then
        Set DOMDocument = New DOMDocument
        DOMDocument.Load gAppDataPath & "\filters.xml"
    Else
        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod)
        DOMDocument.Save gAppDataPath & "\filters.xml"
    End If
    objClient.ResponseToVariant DOMDocument, varStruct
    If Not IsArrayEmpty(varStruct) Then
        If TypeName(varStruct(0)) = "xmlStruct" Then
            With frmPost.TextFilters
                .Add Array(0, "0", GetLbl(lblNone)), "0"
                For b = 0 To UBound(varStruct)
                    Set objStruct = varStruct(b)
                    .Add Array(b + 1, objStruct.Member("key").Value, objStruct.Member("label").Value), objStruct.Member("key").Value
                Next
            End With
        End If
    End If
    LoadTextFilters = True
ExitNow:
    Set objStruct = Nothing
    Set DOMDocument = Nothing
    Set objClient = Nothing
    Exit Function
ErrorHandler:
    LoadTextFilters = False
    ErrorMessage Err.Number, Err.Description, "LoadTextFilters"
    Resume ExitNow
End Function

Public Function Publish(ByVal strBlogID As String) As Boolean
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strPost As String, strMethod As String
Dim objPost As xmlStruct
Dim objLogin As New xmlStruct
Dim udtPost As PostData
Dim varResponse, varStruct()
    frmPost.Message = GetMsg(msgPublishingBlog)
    Screen.MousePointer = vbHourglass
    Set objClient = GetXMLClient()
    If gAccount.GetPostsMethod = API_BLOGGER Then
        varStruct = GetRecentPosts(1)
        Set objPost = varStruct(0)
        strPost = Replace(objPost.Member("content").Value, vbLf, vbCrLf)
        If gAccount.UTF8 Then
            strPost = UTF8_Decode(strPost)
        End If
        If gSettings.AutoConvert Then
            strPost = ConvertHTMLEntities(strPost, False)
        End If
        Publish = Post("", strPost, "", "", "", Array(), strBlogID, objPost.Member("postid").Value, True, True)
    ElseIf gAccount.GetPostsMethod = API_BLOGGER2 Then
        varStruct = GetRecentPosts(1)
        Set objPost = varStruct(0)
        Set udtPost = GetPost(objPost.Member("postid").Value, False)
        Publish = Post(udtPost.Title, udtPost.Text, udtPost.More, udtPost.Excerpt, udtPost.Keywords, Array(), strBlogID, udtPost.PostID, True, True)
'    ElseIf gAccount.GetPostsMethod = API_BLOGGER2 + 10 Then   'To Wait the correct
'        'Create the Login Struct
'        Set objLogin = New xmlStruct
'        objLogin.Add "username", gAccount.User
'        objLogin.Add "password", gAccount.Password
'        objLogin.Add "appkey", APPKEY
'        'Call the Method
'        strMethod = "blogger2.publish"
'        Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
'                                            objLogin, strBlogID)
'        objClient.ResponseToVariant DOMDocument, varResponse
'        Publish = varResponse
    End If
ExitNow:
    Set objLogin = Nothing
    Set objPost = Nothing
    Set DOMDocument = Nothing
    Set objClient = Nothing
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    Publish = False
    ErrorMessage Err.Number, Err.Description, "Publish"
    Resume ExitNow
End Function

Public Function UploadMediaObject(ByVal strBlogID As String, ByVal strFilePath As String) As String
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strPost As String, strMethod As String
Dim objPost As xmlStruct
Dim objMedia As New xmlStruct
Dim udtPost As PostData
Dim varResponse, varStruct(), byteFile() As Byte
    frmPost.Message = GetMsg(msgSendingFile)
    Screen.MousePointer = vbHourglass
    Set objClient = GetXMLClient()
    'Create the Login Struct
    Set objMedia = New xmlStruct
    objMedia.Add "name", GetNamePart(strFilePath)
    objMedia.Add "type", GetMimeType(strFilePath)
    byteFile = StrConv(GetBinaryFile(strFilePath), vbFromUnicode)
    objMedia.Add "bits", byteFile
    'Call the Method
    strMethod = "metaWeblog.newMediaObject"
    Set DOMDocument = objClient.Execute(gAccount.Host, gAccount.Page, strMethod, _
                                        strBlogID, gAccount.User, gAccount.Password, objMedia)
    objClient.ResponseToVariant DOMDocument, varResponse
    UploadMediaObject = varResponse.Member("url").Value
ExitNow:
    Set objMedia = Nothing
    Set DOMDocument = Nothing
    Set objClient = Nothing
    frmPost.Message = ""
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    UploadMediaObject = ""
    ErrorMessage Err.Number, Err.Description, "Publish"
    Resume ExitNow
End Function

Public Function WeblogsPing(strHost As String, strPage As String) As Integer
On Error GoTo ErrorHandler
Dim objClient As xmlClient
Dim DOMDocument As DOMDocument
Dim strMethod As String
Dim objResponse As New xmlStruct
    
    frmPost.Message = GetMsg(msgPingingWeblogs) & "..."
    Screen.MousePointer = vbHourglass
    DoEvents
    
    strMethod = "weblogUpdates.ping"
    
    Set objClient = GetXMLClient()
    Set DOMDocument = objClient.Execute(strHost, strPage, strMethod, _
                                        gBlogs(frmPost.CurrentBlog).Name, _
                                        gBlogs(frmPost.CurrentBlog).URL)
    objClient.ResponseToVariant DOMDocument, objResponse
    If objResponse.Member("flerror").Value Then
        frmPost.Message = GetLbl(lblPing) & ": " & _
        objResponse.Member("message").Value
        WeblogsPing = True
    Else
        frmPost.Message = ""
        WeblogsPing = True
    End If
ExitNow:
    Set objResponse = Nothing
    Set DOMDocument = Nothing
    Set objClient = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrorHandler:
    WeblogsPing = False
    frmPost.Message = GetLbl(lblPing) & ": " & Err.Description
    Resume ExitNow
End Function

Public Function SupportsTitle() As Boolean
    If gAccount.PostMethod = API_METAWEBLOG Or _
       gAccount.PostMethod = API_B2 Or _
       gAccount.PostMethod = API_MT Or _
       gAccount.PostMethod = API_BLOGGER2 Then
        SupportsTitle = True
    ElseIf Trim(gAccount.TitleTag2) <> "" Then
        SupportsTitle = True
    End If
End Function

Public Function SupportsCategory() As Boolean
    If gAccount.PostMethod = API_METAWEBLOG Or _
       gAccount.PostMethod = API_B2 Or _
       gAccount.PostMethod = API_MT Then
        SupportsCategory = True
    ElseIf Trim(gAccount.CategTag2) <> "" Then
        SupportsCategory = True
    End If
End Function

Private Function GetXMLClient() As xmlClient
Dim objClient As xmlClient
    Set objClient = New xmlClient
    objClient.Port = gAccount.Port
    objClient.Secure = gAccount.Secure
    objClient.Timeout = gAccount.Timeout * 1000
    If gAccount.UseProxy > 0 Then
        objClient.UseProxy = True
        If gAccount.UseProxy > 1 Then
            objClient.ProxyServer = gAccount.ProxyServer
            objClient.ProxyPort = gAccount.ProxyPort
            objClient.ProxyUserID = gAccount.ProxyUser
            objClient.ProxyPassword = gAccount.ProxyPassword
        End If
    Else
        objClient.UseProxy = False
    End If
    Set GetXMLClient = objClient
    Exit Function
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, "GetXMLCliente"
End Function
