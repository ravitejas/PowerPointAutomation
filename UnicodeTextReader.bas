Attribute VB_Name = "UnicodeTextReader"
' Code taken from https://github.com/francescofoti/Reading_Text_Files/blob/master/MTextFile.bas ,
' and slightly modified

Option Explicit

Private Const CP_UTF8                       As Long = 65001 'UTF-8 Code Page
Private Const ERROR_NO_UNICODE_TRANSLATION  As Long = 1113& 'No mapping for the Unicode character exists in the target multi-byte code page.
Private Const ERROR_INSUFFICIENT_BUFFER     As Long = 122&
Private Const ERROR_INVALID_PARAMETER       As Long = 87&
Private Const MB_ERR_INVALID_CHARS          As Long = &H8&
Private Const NORM_FORM_C                   As Long = 1&

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
  ByVal CodePage As Long, _
  ByVal dwFlags As Long, _
  ByVal lpMultiByteStr As LongPtr, _
  ByVal cchMultiByte As Long, _
  ByVal lpWideCharStr As LongPtr, _
  ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function IsNormalizedString Lib "kernel32" ( _
  ByVal NormForm As Integer, _
  ByVal lpString As LongPtr, _
  ByVal cwLength As Long) As Long
Private Declare PtrSafe Function NormalizeString Lib "kernel32" ( _
  ByVal NormForm As Integer, _
  ByVal lpSrcString As LongPtr, _
  ByVal cwSrvLength As Long, _
  ByVal lpDstString As LongPtr, _
  ByVal cwDstLength As Long) As Long


Private msLastErr   As String 'Just for UcNormalizeString()

'Ask the WinAPI for the buffer size needed to convert string
'or byte data at address plSrcMemoryAddress to UTF8 (pass CP_UTF8 to plCP).
'Returns -1 if a conversion error (an invalid UTF8 sequence) is encountered.
Private Function MBWCCalcDecodeBufferSize( _
  ByVal plCP As Long, _
  ByVal plSrcMemoryAddress As LongPtr, _
  ByVal plSrcMemoryByteSize As Long, _
  Optional ByVal plStart As Long = 0&, _
  Optional ByVal pfFailOnInvalidChars As Boolean = False) As Long
  
  Dim lCalcSize   As Long
  Dim lFlags      As Long
  
  On Error Resume Next
  If plSrcMemoryByteSize = 0 Then Exit Function
  'Find the length of the buffer we need to create (UTF8 is "MultiByte")
  If pfFailOnInvalidChars Then lFlags = MB_ERR_INVALID_CHARS
  lCalcSize = MultiByteToWideChar(plCP, lFlags, plSrcMemoryAddress + plStart, plSrcMemoryByteSize - plStart, 0, 0)
  If lCalcSize = 0 Then
    If GetLastError() = ERROR_NO_UNICODE_TRANSLATION Then 'True if there were invalid UTF8 chars
      lCalcSize = -1&
    End If
  End If
  
  MBWCCalcDecodeBufferSize = lCalcSize
End Function

'Wrap MultiByteToWideChar() so we can invoke it on a byte array or a string
Private Function MBWCDecodeBuffer( _
  ByVal plCP As Long, _
  ByVal plSrcMemoryAddress As LongPtr, _
  ByVal plSrcMemoryByteSize As Long, _
  ByVal plDstMemoryAddress As LongPtr, _
  ByVal plDstMemoryByteSize As Long, _
  Optional ByVal pfFailOnInvalidChars As Boolean = False) As Long

  'Convert directly into string memory
  Dim lFlags      As Long
  If pfFailOnInvalidChars Then lFlags = MB_ERR_INVALID_CHARS
  MBWCDecodeBuffer = MultiByteToWideChar(plCP, 0, plSrcMemoryAddress, plSrcMemoryByteSize, plDstMemoryAddress, plDstMemoryByteSize)
End Function

Public Function UTF8DecodeByteArrayToString( _
  ByRef pabBytes() As Byte, _
  Optional ByVal plStart As Long = 0&) As String
  
  Dim lConvLen    As Long
  Dim lLength     As Long
  Dim sRes        As String
  
  On Error Resume Next
  lLength = UBound(pabBytes) - LBound(pabBytes) + 1
  If lLength = 0 Then Exit Function
  'Find the length of the buffer we need to create (UTF8 is "MultiByte")
  lConvLen = MBWCCalcDecodeBufferSize(CP_UTF8, VarPtr(pabBytes(LBound(pabBytes))), lLength, plStart)
  If lConvLen > 0 Then
    sRes = String$(lConvLen, 0) 'create the buffer as string (vb strings are UNICODE ie, "WideChar")
    'Now convert directly into string memory
    Call MBWCDecodeBuffer(CP_UTF8, VarPtr(pabBytes(LBound(pabBytes))) + plStart, lLength - plStart, StrPtr(sRes), lConvLen)
  End If
  
  UTF8DecodeByteArrayToString = sRes
End Function

Public Function UTF8DecodeString(ByVal psSource As String) As String
  Dim lConvLen    As Long
  Dim lLength     As Long
  Dim sRes        As String
  Dim sDecoded    As String
  Dim sDecodedACP As String
  Dim lUtf8Len    As Long
  
  On Error Resume Next
  lLength = Len(psSource)
  If lLength = 0 Then Exit Function
  'DumpStringBytes psSource
  'Find the length of the buffer we need to create (UTF8 is "MultiByte")
  lConvLen = MBWCCalcDecodeBufferSize(CP_UTF8, StrPtr(psSource), lLength, 0)
  If lConvLen > 0 Then
    sRes = String$(lConvLen, 0) 'create the buffer as string (vb strings are UNICODE ie, "WideChar")
    'Now convert directly into string memory
    lUtf8Len = MBWCDecodeBuffer(CP_UTF8, StrPtr(psSource), lLength, StrPtr(sRes), lConvLen)
    'DumpStringBytes sRes
    If Not UcIsNormalizedString(sRes) Then
      sRes = UcNormalizeString(sRes)
    End If
    'DumpStringBytes sRes
    'Debug.Print "DECODED C: " & sRes & "< (len=" & Len(sRes) & ")"
  End If
  
  UTF8DecodeString = sRes
End Function

'Just wrap Win32 GetLastError to expose it publicly
Public Function UcGetLastError() As Long
  UcGetLastError = GetLastError()
End Function

Public Function UcGetLastErrorText() As String
  UcGetLastErrorText = msLastErr
End Function

Public Function UcIsNormalizedString(ByVal psText As String) As Boolean
  UcIsNormalizedString = CBool(IsNormalizedString(NORM_FORM_C, StrPtr(psText), Len(psText)) > 0)
End Function

Public Function UcNormalizeString(ByVal psText As String) As String
  Dim sResult         As String
  Dim iSizeGuess      As Long
  Dim iActualSize     As Long
  Dim iFoolGuardCt    As Integer
  Const MAX_ITERATIONS As Integer = 100
  
  msLastErr = ""
  iSizeGuess = NormalizeString(NORM_FORM_C, StrPtr(psText), Len(psText), 0&, 0&)
  If iSizeGuess = 0 Then
    msLastErr = "Error checking for size"
    Exit Function
  End If
  
  Do While iSizeGuess > 0
    sResult = String$(iSizeGuess, vbNullChar)
    iActualSize = NormalizeString(NORM_FORM_C, StrPtr(psText), Len(psText), StrPtr(sResult), iSizeGuess)
    If iActualSize > 0 Then Exit Do
    If iActualSize <= 0 Then
      If GetLastError() = ERROR_INSUFFICIENT_BUFFER Then
        iSizeGuess = -iActualSize
      ElseIf GetLastError() = ERROR_NO_UNICODE_TRANSLATION Then
        msLastErr = "Invalid unicode found at index " & -iActualSize
        Exit Do
      Else
        msLastErr = "Error #" & GetLastError()
      End If
    End If
    iFoolGuardCt = iFoolGuardCt + 1
    If iFoolGuardCt > MAX_ITERATIONS Then
      msLastErr = "Too many iterations (>" & MAX_ITERATIONS & ") trying to expand normalized string"
      Exit Do
    End If
  Loop
  
  UcNormalizeString = sResult
End Function

'GetFileText()
'
'Synopsis
'--------
' Reads a text file as a binary file in memory, and converts the bytes into a VB string.
' Detects BOM (Byte Order Mark) if there's one and handles BE/LE (Big/Little Endian).
'
'Parameters
'----------
' psFilename
'   Full or relative path of file to read
'
' Rules applied for psInFileFormat:
'   2. "utf8", if there's a BOM, its not included in the returned
'      text, and the text is UTF8decoded to a VB/A string (UTF16)
'
'Returns
'-------
' The decoded file text.
'
'Notes:
'-----
'1. Can handle files which size fits into a long, and in memory.
'2. Use error trapping in your calling code to catch unexpected errors.
'3. You can use Notepad++ to produce test files (See the "Encoding" menu)
'   Notepad++ encoding        | Translates to
'   --------------------------+--------------
'   "Encode in UTF-8"         | "utf8"
'   "Encode in UTF-8-BOM"     | "utf8"
'4. Tools like Typora (https://typora.io/) save files in utf8 with no BOM,
'   Use "GetFileText(YourFilename,"utf8") to load them.
'
Sub GetFileText(ByVal psFileName As String, ByVal psInFileFormat As String, ByRef sFileContents As String)
  Dim hFile       As Integer
  Dim lLength     As Long
  Dim sText       As String
  Dim lStart      As Long
  Dim fIsOpen     As Boolean
  Dim fNoBOM      As Boolean
  Dim I           As Long
  Dim fLittEndian As Boolean
  
  hFile = FreeFile()
  
  ' Let others read but not write
  Open psFileName For Binary Access Read As hFile
  fIsOpen = True
  lLength = LOF(hFile)
  If lLength > 0 Then
      ReDim abBytes(0 To lLength - 1) As Byte
      Get hFile, 1, abBytes()
      Close hFile: fIsOpen = False
      If (psInFileFormat = "utf8") Then
        'skip the BOM bytes if present (3 bytes EF,BB,BF)
        If (lLength > 3) And (abBytes(0) = &HEF) And (abBytes(1) = &HBB) And (abBytes(2) = &HBF) Then
          lStart = 3&
        End If
        sText = UTF8DecodeByteArrayToString(abBytes(), lStart)
      End If
  End If
  
  If fIsOpen Then
    Close hFile
  End If
  
  sFileContents = sText
End Sub
