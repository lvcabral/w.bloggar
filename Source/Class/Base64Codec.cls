VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlBase64Codec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*
'*  CLASS INFORMATION
'*
'*  Product       : XML-RPC implementation for COM/MTS
'*  File name     : Base64Codec.cls
'*  Class name    : Base64Codec
'*  Main module   : None
'*
'*  Date          : 08 Jul 2001
'*  Programmer    : Jan G.P. Sijm
'*
'*  Description   :
'*    This class implements the base64 encode/decode algorithm. This algorithm is
'*    described in RFC 2045.
'*
'*  Comments      :
'*    Some routines in this class require the concatenation of strings,
'*    but the & operand in Visual Basic is very slow when called repeatedly.
'*    That's why this class uses a pre-allocated string and the Mid$ statement
'*    to update parts of that string. The result is the same, but the
'*    performance is much better.
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
Private Const Base64Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

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
'*  This method will encode the specified array of bytes using the
'*  base64 encoding algorithm described in RFC 2045. A string containing
'*  the encoded bytes is returned.
'*
'*  Input    : xBytes, Array of bytes to be encoded
'*  Returns  : String containing base64 encoded bytes
'*-------------
Public Function EncodeArray(ByRef xBytes() As Byte) As String

  Dim lUBound As Long, lIndex As Long, lCount As Long
  Dim sEncoded As String
  Dim xTriplet(2) As Byte
  Dim iLength As Integer
  
  lUBound = UBound(xBytes)
  sEncoded = String$(Int((lUBound + 3) / 3) * 4, 0)
  
  '
  ' Loop through all bytes in the array
  '
  For lIndex = 0 To lUBound
    
    xTriplet(iLength) = xBytes(lIndex)
    iLength = iLength + 1
    
    If iLength = 3 Then
      
      '
      ' Encode a sequence of 3 bytes into a sequence of 4 bytes
      '
      Mid$(sEncoded, (lCount * 4) + 1, 4) = EncodeTriplet(xTriplet, iLength)
      Erase xTriplet
      iLength = 0
      lCount = lCount + 1
      
    End If
    
  Next

  If iLength <> 0 Then
  
    '
    ' Encode the last triplet (may be an incomplete triplet)
    '
    Mid$(sEncoded, (lCount * 4) + 1, 4) = EncodeTriplet(xTriplet, iLength)
    lCount = lCount + 1
  
  End If

  '
  ' Limit line length to 76 characters
  '
  sEncoded = AddLineBreaks(sEncoded)

  EncodeArray = sEncoded

End Function

'*-------------
'*  This method will decode the base64 encoded string into an
'*  array of bytes. The decoded array of bytes is returned.
'*
'*  Input    : sEncoded, String containing base64 encoded bytes
'*  Returns  : Array of decoded bytes
'*-------------
Public Function DecodeArray(ByVal sEncoded As String) As Byte()

  Dim lIndex As Long
  Dim xBytes() As Byte, xQuad(3) As Byte, xTriplet() As Byte
  Dim sCharacter As String
  Dim iLength As Integer
    
  '
  ' Loop through all characters in the encoded string
  '
  For lIndex = 1 To Len(sEncoded)
  
    sCharacter = Mid$(sEncoded, lIndex, 1)
    
    '
    ' Skip all characters not in the 'base64 alphabet'
    '
    If InStr(Base64Alphabet, sCharacter) <> 0 Then
    
      xQuad(iLength) = InStr(Base64Alphabet, sCharacter) - 1
      iLength = iLength + 1
      
      If iLength = 4 Then
      
        '
        ' Decode a sequence of 4 characters into a sequence of 3 bytes
        '
        xTriplet = DecodeTriplet(xQuad, iLength)
        AddTriplet xBytes, xTriplet
        
        Erase xQuad
        iLength = 0
      
      End If
    
    End If
  
  Next
  
  If iLength <> 0 Then
  
    '
    ' Decode the last characters into a triplet
    '
    xTriplet = DecodeTriplet(xQuad, iLength)
    AddTriplet xBytes, xTriplet
  
  End If

  DecodeArray = xBytes

End Function

'*-------------
'*  This function will take an array of 1 to 3 bytes and convert it into
'*  a string of 4 base64 encoded characters.
'*
'*  Input    : xTriplet, An array containing 3 bytes to be encoded
'*  Returns  : A 4 byte base64 encoded string
'*-------------
Private Function EncodeTriplet(ByRef xTriplet() As Byte, ByVal iLength As Integer) As String

  Dim xEncoded(3) As Byte
  Dim sEncoded As String
  
  '
  ' Convert 3 bytes into 4 bytes of 6-bit values
  '
  xEncoded(0) = Int(xTriplet(0) / 4)
  xEncoded(1) = ((xTriplet(0) And 3) * 16) + Int(xTriplet(1) / 16)
  xEncoded(2) = ((xTriplet(1) And 15) * 4) + Int(xTriplet(2) / 64)
  xEncoded(3) = xTriplet(2) And 63
 
  '
  ' Convert the bytes to a string and pad with '=' to a length of 4 characters
  '
  Select Case iLength
    Case 1
      sEncoded = Mid$(Base64Alphabet, xEncoded(0) + 1, 1) & Mid$(Base64Alphabet, xEncoded(1) + 1, 1) & "=="
    Case 2
      sEncoded = Mid$(Base64Alphabet, xEncoded(0) + 1, 1) & Mid$(Base64Alphabet, xEncoded(1) + 1, 1) & Mid$(Base64Alphabet, xEncoded(2) + 1, 1) & "="
    Case 3
      sEncoded = Mid$(Base64Alphabet, xEncoded(0) + 1, 1) & Mid$(Base64Alphabet, xEncoded(1) + 1, 1) & Mid$(Base64Alphabet, xEncoded(2) + 1, 1) & Mid$(Base64Alphabet, xEncoded(3) + 1, 1)
  End Select

  EncodeTriplet = sEncoded

End Function

'*-------------
'*  TODO : Describe the routine...
'*
'*  Input    : None
'*  Returns  : TODO : Describe return value
'*-------------
Private Function DecodeTriplet(ByRef xQuad() As Byte, ByVal iLength As Integer) As Byte()

  Dim xTriplet() As Byte
  
  ReDim xTriplet(2) As Byte
  
  '
  ' Convert 4 bytes of 6-bit values into 3 bytes
  '
  xTriplet(0) = (xQuad(0) * 4) + (Int(xQuad(1) And 48) / 16)
  xTriplet(1) = ((xQuad(1) And 15) * 16) + (Int(xQuad(2) And 252) / 4)
  xTriplet(2) = ((xQuad(2) And 3) * 64) + xQuad(3)

  '
  ' Convert the triplet to the correct length
  '
  Select Case iLength
    Case 2
      ReDim Preserve xTriplet(0) As Byte
    Case 3
      ReDim Preserve xTriplet(1) As Byte
    Case 4
      ReDim Preserve xTriplet(2) As Byte
  End Select

  DecodeTriplet = xTriplet

End Function

'*-------------
'*  This function will split the string into line of 76 or less characters.
'*  Lines are separated by a CR/LF sequence.
'*
'*  Input    : sEncoded, A base64 encoded string
'*  Returns  : A base64 encoded string with line length limited to 76 characters
'*-------------
Private Function AddLineBreaks(ByVal sEncoded As String) As String

  Dim sBreaksAdded As String
  Dim lIndex As Long, lLength As Long, lCount As Long
  
  lLength = Len(sEncoded)
  lIndex = 1
  
  '
  ' Pre-allocate space for the new string
  '
  sBreaksAdded = String$(Int((lLength + 75) / 76) * 78, 0)
  
  Do While lIndex <= lLength
  
    '
    ' Get the next line of max. 76 characters and add CRLF
    '
    Mid$(sBreaksAdded, (lCount * 78) + 1, 78) = Mid$(sEncoded, lIndex, 76) & vbCrLf
    lIndex = lIndex + 76
    lCount = lCount + 1
  
  Loop

  '
  ' Set the length of the pre-allocated string to the correct value
  '
  lIndex = InStr(sBreaksAdded, Chr$(0))

  If lIndex <> 0 Then
    sBreaksAdded = Left$(sBreaksAdded, lIndex - 1)
  End If

  AddLineBreaks = sBreaksAdded

End Function

'*-------------
'*  This procedure will add a triplet to the byte array.
'*
'*  Input    : xBytes  , Add the triplet to this array
'*             xTriplet, Add this triplet
'*-------------
Private Sub AddTriplet(ByRef xBytes() As Byte, ByRef xTriplet() As Byte)

  Dim lCount As Long, lUBound As Long, lIndex As Long
  
  On Error Resume Next
  lCount = UBound(xBytes) + 1
  On Error GoTo 0
  
  lUBound = UBound(xTriplet)
  
  '
  ' Increase the size of the byte array
  '
  ReDim Preserve xBytes(lCount + lUBound)
  
  '
  ' Store the triplet in the array
  '
  For lIndex = 0 To lUBound
    xBytes(lCount + lIndex) = xTriplet(lIndex)
  Next

End Sub
