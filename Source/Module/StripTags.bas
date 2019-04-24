Attribute VB_Name = "basStripTags"
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
Option Compare Text
Option Explicit

Public Function HTML2Text(HTMLString As String, Optional SaveAsFile As String) As String
  Const MAX_ROW_LENGTH = 65
  Const MAX_LINE_LENGTH = 75
  
  Dim sHTML As String
  Dim sOut As String
  Dim sWkg As String
  Dim lLen As Long
  Dim lngLoop As Long, lngCtr As Long

  Dim sChar As String
  Dim sTag As String
  Dim bBodyStart As Boolean, bBodyTag As Boolean
  
  Dim bPrevSpace As Boolean
  Dim sCharCode As String
  Dim bOL As Boolean, bUL As Boolean
  Dim iPlaceInList As Integer
  
  Dim iFileNum As Integer
  
  Dim bOneCrLf As Boolean
  Dim bTwoCrlf As Boolean
  Dim lTempCtr As Long, iTempCtr As Integer
  Dim lTempCtr2 As Long

  Dim bFormatCell As Boolean
  
  Dim lRowLength As Long
  Dim iLineCount As Integer
  
  
  Dim bInComment As Boolean
  Dim sTemp As String, sTemp2 As String
  
  
  Dim bFlag As Boolean
  Dim bSubFlag As Boolean
  
  Dim bOutputCells As Boolean
  Dim lRowCharCount As Long
  Dim sNestedTag As String
  Dim sCharInCell As String
  Dim sTagInCell As String
  Dim sEndTag As String
  
  Dim bInCells As Boolean
  Dim bInScript As Boolean
  

  sHTML = HTMLString
  lLen = Len(sHTML)

  For lngCtr = 1 To lLen

    sTag = ""
    sChar = Mid(sHTML, lngCtr, 1)
    

        If sChar = "<" And Not bInComment Then
     

                lngCtr = lngCtr + 1
                lngLoop = 1
                sWkg = ""
                'start new loop to get the tag name
                Do
                'if we never find end, then we must exit
                If lngCtr = lLen Then Exit For

                sChar = Mid(sHTML, lngCtr, 1)
                    If sChar <> ">" Then
                        sWkg = sWkg & sChar
                        
                        lngCtr = lngCtr + 1
                        
                     End If

                     If sChar = ">" Or lngCtr >= lLen Then
                        If lngCtr < lLen Then
                            If Asc(Mid(sHTML, lngCtr + 1, 1)) < 32 Then bPrevSpace = True
                         
                        End If
                        'lTempCtr = lngCtr

                        Exit Do
                       End If
                        
                        
                    
                    If RemoveAllSpaces(sWkg = "!--") Then Exit Do
               
                     
                     'lngCtr = lngCtr + 1
                      
                    Loop
                    sTag = Trim(sWkg)
          
                     'determine if another tag is coming because
                        'if so, we don't want to output any spaces.
        
                        bFlag = False
                        bSubFlag = False
                        lTempCtr = lngCtr + 1
                        If lTempCtr >= lLen Then Exit For
                        sTemp = Mid(sHTML, lTempCtr, 1)
                        If Asc(sTemp) <= 32 Then
                            sTemp = ""
                    
                          
                            Do
                      
                                sTemp = sTemp & Mid(sHTML, lTempCtr, 1)
                                If sTemp = "<" Then
                                    If bFlag Then bPrevSpace = True
                                    Exit Do
                                
                                ElseIf Asc(sTemp) > 32 Then
                            
                                    bPrevSpace = Not (bSubFlag)
                                
                                    Exit Do
                                ElseIf Asc(sTemp) <= 32 Then
                                    bFlag = True
                                    bSubFlag = Asc(sTemp) = 32
                            
                                End If
                       
                            
                                lTempCtr = lTempCtr + 1
                                If lTempCtr >= lLen Then Exit For
                                sTemp = ""
                            Loop
                        End If
    
                    'Certain tags interest us: TITLE, <BR><P>
                  
                    If InStr(Left(sTag, 1), "/") = 0 Then
                         If Left(sTag, 5) = "TITLE" Then
                         lngCtr = lngCtr + 1
                           
                            Do
                                
                                sChar = Mid(sHTML, lngCtr, 1)

                                If (sChar = "<" And sChar <> Chr$(13) And sChar <> Chr$(10)) Or lngCtr = lLen Then
                                    If Not bInComment And Not bInScript Then sOut = sOut & vbCrLf & vbCrLf
                                    iLineCount = 0
                                    bTwoCrlf = True
                                    lngCtr = lngCtr - 1
                                    Exit Do
                                End If
                               
                            sOut = sOut & sChar
                         
                            lngCtr = lngCtr + 1
                            Loop
                        ElseIf Left(sTag, 4) = "BODY" And Not bInScript Then
                            bBodyTag = True
                        ElseIf (sTag = "P" Or Left(sTag, 2) = "P ") And Not bInScript Then
                           If bBodyStart And Not bTwoCrlf And Not bInScript And Not bInComment Then
                                sOut = sOut & vbCrLf & vbCrLf
                                iLineCount = 0
                                bTwoCrlf = True
                            End If
                        ElseIf (sTag = "TR" Or Left(sTag, 3) = "TR ") And Not bInScript Then
                                
                                lTempCtr = lngCtr + 1
                                lRowCharCount = 0
                                bFlag = False
                                Do
                                    sTemp = Mid(sHTML, lTempCtr, 1)
                                        If sTemp = "<" Then 'get name of tag
                                            sNestedTag = ""
                                            Do
                                                lTempCtr = lTempCtr + 1
                                                If lTempCtr >= lLen Then Exit For
                                                sTemp2 = Mid(sHTML, lTempCtr, 1)
                                                If sTemp2 = ">" Then Exit Do
                                                sNestedTag = sNestedTag & sTemp2
                                            Loop
                                        End If
                                        
                                        If (sNestedTag = "/TR" Or sNestedTag = "/TABLE") And Not bInScript Then
                                            bOutputCells = (lRowCharCount < MAX_ROW_LENGTH)
                                            Exit Do
                                        ElseIf (sNestedTag = "TABLE" Or Left$(sNestedTag, 6) = "TABLE ") And Not bInScript Then
                                            bOutputCells = False
                                            Exit Do
                                        ElseIf (sNestedTag = "TD" Or Left$(sNestedTag, 3) = "TD " _
                                            Or sNestedTag = "TH" Or Left$(sNestedTag, 3) = "TH ") And Not bInScript Then
                                            lTempCtr = lTempCtr + 1
                                            
                                            bFlag = False
                                                Do
                                                If lTempCtr >= lLen Then Exit For
                                                sCharInCell = Mid(sHTML, lTempCtr, 1)
                                                Select Case sCharInCell
                                                    Case "<" 'nested tag
                                                        lTempCtr = lTempCtr + 1
                                                        sTagInCell = ""
                                                         Do
                                                          If lTempCtr >= lLen Then Exit For
                                                           sTemp2 = Mid(sHTML, lTempCtr, 1)
                                                            If sTemp2 <> ">" Then
                                                             sTagInCell = sTagInCell & sTemp2
                                                           Else
                                                            Exit Do
                                                           End If
                                                           lTempCtr = lTempCtr + 1
                                                         Loop
                                                       
                                                         
                                                        If RemoveAllSpaces(sTagInCell) = "/TD" Then
                                                            sNestedTag = ""
                                                            Exit Do
                                                        ElseIf (sTagInCell = "P" Or Left$(sTagInCell, 2) = "P " _
                                                            Or sTagInCell = "BR" Or Left$(sTagInCell, 3) _
                                                            = "BR ") And Not bInScript Then
                                                                
                                                                lRowCharCount = MAX_ROW_LENGTH + 1
                                                                Exit Do
                                                       End If
                                                        Case Else
                                                            If Not bFlag And Not bInScript Then lRowCharCount = lRowCharCount + 1
                                                        End Select
                                                      lTempCtr = lTempCtr + 1
                                                    Loop
                                                End If 'td tag
                                                
                                    If lTempCtr = lLen Then Exit For
                                    lTempCtr = lTempCtr + 1
                                  Loop 'loop begins under the TR condition
                               lRowCharCount = 0
                              
                                sOut = sOut & vbCrLf
                                iLineCount = 0
                                bOneCrLf = True
                           
                               bInCells = False
                        ElseIf sTag = "TD" Or Left(sTag, 3) = "TD " _
                            Or sTag = "TH" Or Left(sTag, 3) = "TH " Then
                                If bOutputCells Then
                                   If bInCells Then sOut = sOut & Space$(3)
                                   bInCells = True
                                Else
                                    sOut = sOut & vbCrLf
                                    bOneCrLf = True
                               End If
                        
                        ElseIf sTag = "BR" Or sTag = "TABLE" Or Left$(sTag, 5) = "TABLE" Then
                              If bBodyStart And Not bOneCrLf Then
                                    sOut = sOut & vbCrLf
                                    iLineCount = 0
                                    bOneCrLf = True
                                End If
                        ElseIf sTag = "OPTION" Or Left(sTag, 7) = "OPTION " Then
                                sOut = sOut & vbCrLf & vbTab
                                iLineCount = 0
                        ElseIf sTag = "SCRIPT" Or Left(sTag, 7) = "SCRIPT " Then
                                bInScript = True
                                
                        ElseIf Left(sTag, 3) = "!--" And bBodyTag Then
                                bInComment = True
                        ElseIf sTag = "OL" Or Left(sTag, 3) = "OL " Then
                            bOL = True
                            sOut = sOut & vbCrLf & vbCrLf
                            iLineCount = 0
                        ElseIf sTag = "UL" Or Left(sTag, 3) = "UL " Then
                            bUL = True
                            sOut = sOut & vbCrLf & vbCrLf
                            iLineCount = 0
                        ElseIf sTag = "LI" Or Left(sTag, 3) = "LI " Then
                            'if not in the middle of a numbered list, just add bullet
                            sOut = sOut & vbCrLf
                            iLineCount = 0
                            If bOL Then
                                iPlaceInList = iPlaceInList + 1
                                sOut = sOut & iPlaceInList & ". "
                                iLineCount = iLineCount + 2
                            Else
                                sOut = sOut & Chr$(149) & " "
                                iLineCount = iLineCount + 2
                            End If
                            
                        End If

                    Else 'end tag
                    
                    If Left(RemoveAllSpaces(sTag), 7) = "/SCRIPT" Then bInScript = False
                      
                      If bBodyStart Then
                        'we need to find the end for bOL and bUL
                        'if you want to process other end tags
                        'do it here.
                        Select Case Left(RemoveAllSpaces(sTag), 3)
                            Case "/OL"
                                bOL = False
                                If Not bTwoCrlf Then
                                    sOut = sOut & vbCrLf & vbCrLf
                                    iLineCount = 0
                                    bTwoCrlf = True
                                End If
                                iPlaceInList = 0
                            Case "/UL"
                                bUL = False
                                If Not bTwoCrlf Then
                                    sOut = sOut & vbCrLf & vbCrLf
                                    iLineCount = 0
                                    bTwoCrlf = True
                                End If
                            End Select
                    
                    End If 'instr(stag, "/")
                End If 'bbodystart

           Else 'not a tag
            sChar = Mid(sHTML, lngCtr, 1)
           If bBodyTag Then
            Select Case sChar
                Case "<" 'another new tag
                If Not bInComment And Not bInScript Then
                    lngCtr = lngCtr - 1 'go back and let top of loop handle tag
                    sTag = ""
                    sWkg = ""
                End If

                Case " "
                    'only one space is processed in HTML
                    'rest are ignored
                    If bPrevSpace = False Then
                        If Not bInComment And Not bInScript Then sOut = sOut & sChar
                        bPrevSpace = True
                        iLineCount = iLineCount + 1
                        
                        If iLineCount >= MAX_LINE_LENGTH Then
                            sOut = sOut & vbCrLf
                            iLineCount = 0
                        End If
                    End If
                Case "-" 'see if this is the end of the comment
                If bBodyStart Then
                    If bInComment Then
                         sTemp = ""
                        
                       
                        lTempCtr = lngCtr
                        lTempCtr2 = 0
                        
                        Do
                            sTemp = sTemp & Mid(sHTML, lTempCtr, 1)
                            
                            If Mid(sHTML, lTempCtr, 1) = ">" Then
                                sTemp2 = RemoveAllSpaces(sTemp)
                                If Right$(sTemp2, 3) = "-->" Then
                                    bInComment = False
                                    Exit Do
                                End If
                            End If
                            If lTempCtr = lLen Then Exit For
                            lTempCtr = lTempCtr + 1
                            lTempCtr2 = lTempCtr2 + 1
                        Loop
                        If lTempCtr < lLen Then lngCtr = lngCtr + lTempCtr2
                    Else
                        
                        bPrevSpace = False
                        sOut = sOut & "-"
                        bOneCrLf = False
                        bTwoCrlf = False
                        iLineCount = iLineCount + 1
                     End If
                 End If
                Case "&" 'special character code, or maybe just an ampersand
                    
                    sTemp = ""
                    bFlag = False
                   
                        For lTempCtr = (lngCtr + 1) To (lngCtr + 7)
                            sTemp = Mid(sHTML, lTempCtr, 1)
                            If sTemp = ";" Then
                                bFlag = True
                                Exit For
                            ElseIf sTemp = "&" Then
                                bFlag = False
                                Exit For
                            End If
                        
                        Next
                    
                If bFlag Then
                    sCharCode = ""
                    lngCtr = lngCtr + 1
                    Do
                        sChar = Mid(sHTML, lngCtr, 1)
                        If sChar = ";" Then Exit Do
                        sCharCode = sCharCode + sChar
                        lngCtr = lngCtr + 1
                        
                    Loop
                    'special character. must end with ";"
                    If Not bInComment And Not bInScript Then
                        sTemp2 = HTMLSpecChar2ASCII(sCharCode)
                        sOut = sOut & sTemp2
                        bPrevSpace = False
                        bOneCrLf = False
                        bTwoCrlf = False
                        iLineCount = iLineCount + Len(sTemp2)
                    End If
                Else
                    If Not bInComment And Not bInScript Then
                        sOut = sOut & "&"
                        bPrevSpace = False
                        bOneCrLf = False
                        bTwoCrlf = False
                        iLineCount = iLineCount + 1
                    End If
                End If


                Case Else
                    bBodyStart = True
                    'asc below 31 = nonprintable
                  If Asc(sChar) < 31 Then
                    If bPrevSpace = False Then
                         If Not bInComment And Not bInScript Then sOut = sOut & " "
                         iLineCount = iLineCount + 1
                         bPrevSpace = True
                         
                         If iLineCount >= MAX_LINE_LENGTH Then
                            sOut = sOut & vbCrLf
                            iLineCount = 0
                        End If
                    End If
                Else
                    If Not bInComment And Not bInScript And Asc(sChar) > 31 Then
                        sOut = sOut & sChar
                        
                        bPrevSpace = False
                        bOneCrLf = False
                        bTwoCrlf = False
                        iLineCount = iLineCount + 1
                  
                    End If
                End If
            End Select
            End If 'bbodystart
        End If 'sChar = "<"

    DoEvents

   
    Next



  'return output
HTML2Text = sOut

  If SaveAsFile <> "" Then
    On Error GoTo ErrorHandler
    'save output to string
    iFileNum = FreeFile
    Open SaveAsFile For Output As #iFileNum
    Print #iFileNum, sOut
    Close #iFileNum
    
  End If


Exit Function
ErrorHandler:

On Error Resume Next
Close #iFileNum
Exit Function

End Function

Private Function BinaryEqualityTest(String1 As String, _
String2 As String) As Boolean

        BinaryEqualityTest = (StrComp(String1, String2, _
           vbBinaryCompare) = 0)

End Function
Private Function HTMLSpecChar2ASCII(ByVal HTMLCode As String) As String

Dim sAns As String, sInput As String

sInput = LCase(HTMLCode)
If Left$(sInput, 1) = "#" Then
   sInput = Mid(sInput, 2)
End If

If IsNumeric(sInput) Then
    sAns = Chr$(Val(sInput))
Else
    Select Case sInput
    Case "quot"
        sAns = ""
    Case "amp"
        sAns = "&"
    Case "lt"
        sAns = "<"
    Case "gt"
        sAns = ">"
    Case "nbsp"
        sAns = Chr$(160)
    Case "iexcl"
        sAns = Chr$(161)
    Case "cent"
        sAns = Chr$(162)
    Case "pound"
        sAns = Chr$(163)
    Case "curren"
        sAns = Chr$(164)
    Case "yen"
        sAns = Chr$(165)
    Case "brvbar"
        sAns = Chr$(166)
    Case "sect"
        sAns = Chr$(167)
    Case "uml"
        sAns = Chr$(168)
    Case "copy"
        sAns = Chr$(169)
    Case "ordf"
        sAns = Chr$(170)
    Case "laquo"
        sAns = Chr$(171)
    Case "not"
        sAns = Chr$(172)
    Case "shy"
        sAns = Chr$(173)
    Case "reg"
        sAns = Chr$(174)
    Case "macr"
        sAns = Chr$(175)
    Case "deg"
        sAns = Chr$(176)
    Case "plusmn"
        sAns = Chr$(177)
    Case "sup2"
        sAns = Chr$(178)
    Case "sup3"
        sAns = Chr$(179)
    Case "acute"
        sAns = Chr$(180)
    Case "micro"
        sAns = Chr$(181)
    Case "para"
        sAns = Chr$(182)
    Case "middot"
        sAns = Chr$(183)
    Case "cedil"
        sAns = Chr$(184)
    Case "supl"
        sAns = Chr$(185)
    Case "ordm"
        sAns = Chr$(186)
    Case "raquo"
        sAns = Chr$(187)
    Case "frac14"
        sAns = Chr$(188)
    Case "frac12"
        sAns = Chr$(189)
    Case "frac34"
        sAns = Chr$(190)
    Case "iquest"
        sAns = Chr$(191)
    Case "agrave"
       sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(224), Chr$(192))
    Case "aacute"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(225), Chr$(193))
    Case "acirc"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(226), Chr$(194))
    Case "atilde"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(227), Chr$(195))
    Case "auml"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(228), Chr$(196))
    Case "aring"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(229), Chr$(197))
    Case "aelig"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(230), Chr$(198))
    Case "ccedil"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(231), Chr$(199))
    Case "egrave"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(232), Chr$(200))
    Case "eacute"
         sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(233), Chr$(201))
    Case "ecirc"
         sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(234), Chr$(202))
    Case "euml"
         sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(235), Chr$(203))
    Case "igrave"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(236), Chr$(204))
    Case "iacute"
         sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(237), Chr$(205))
    Case "icirc"
         sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(238), Chr$(206))
    Case "iuml"
         sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(239), Chr$(207))
    Case "eth"
         sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(240), Chr$(208))
    Case "ntilde"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(241), Chr$(209))
    Case "ograve"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(242), Chr$(210))
    Case "oacute"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(243), Chr$(211))
    Case "ocirc"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(244), Chr$(212))
    Case "otilde"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(245), Chr$(213))
    Case "otilde"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(245), Chr$(213))
    Case "ouml"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(246), Chr$(214))
    Case "times"
        sAns = Chr$(215)
    Case "oslash"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(248), Chr$(216))
    Case "ugrave"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(249), Chr$(217))
    Case "uacute"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(250), Chr$(218))
    Case "ucirc"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(251), Chr$(219))
     Case "uuml"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(252), Chr$(220))
     Case "yacute"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(253), Chr$(221))
     Case "thorn"
        sAns = IIf(BinaryEqualityTest(sInput, HTMLCode) = True, Chr$(254), Chr$(222))
    Case "szlig"
        sAns = Chr$(223)
    Case "divide"
        sAns = Chr$(247)
    Case "yuml"
        sAns = Chr$(255)
    
        
    End Select
End If

HTMLSpecChar2ASCII = sAns
End Function

Private Function RemoveAllSpaces(ByVal InputString As String) _
As String

Dim sAns As String
Dim lLen As String
Dim lCtr As Long, lCtr2 As Long
Dim sChar As String


lLen = Len(InputString)
sAns = InputString
lCtr2 = 1

For lCtr = 1 To lLen
    sChar = Mid(InputString, lCtr, 1)
    If sChar <> " " Then
        Mid(sAns, lCtr2, 1) = sChar
        lCtr2 = lCtr2 + 1
    End If
Next

If lCtr2 > 1 Then
    sAns = Left(sAns, lCtr2 - 1)
Else
    sAns = ""
End If

RemoveAllSpaces = sAns
End Function
