Attribute VB_Name = "resource"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long



'------------------------------------------------------------------------------------
' 리소스 관련 함수들 공통모듈로 빼야 할듯...
'------------------------------------------------------------------------------------
Public Function LoadWaveResource( _
    ByVal ResourceID As Variant, _
    ByVal sResourceType As String, _
    Optional TempFile) As String

    Dim sFileName As String
    
    'Check if the TempFile Name has been specified
    If IsMissing(TempFile) Then
        'Create a temp file name such as "~res1234.tmp"
        GetTempFile "", "~rs", 0, sFileName
    Else
        'Use the specified temp file name
        sFileName = TempFile
    End If
    
    'Save the resource item to disk
    If SaveResItemToDisk(ResourceID, sResourceType, sFileName) = 0 Then
    
        'Return the picture
'        Set LoadPictureResource = LoadPicture(sFileName)
         
        '에러나서 막음. sndPlaySound 정의 후 풀것
        Call sndPlaySound(sFileName, SND_ASYNC)
        
        'Delete the temp file
        Kill sFileName
    End If
End Function

Public Function LoadPictureResource( _
    ByVal ResourceID As Long, _
    ByVal sResourceType As String, _
    Optional TempFile _
    ) As Picture
    '=====================================================
    'Returns a picture object from a resource file.
    'Used for loading images other than ICO and BMP into a
    'Picture property. (Such as GIF and JPG images)
    '=====================================================

    'EXAMPLE CALL:
    'Set Picture1.Picture = LoadPictureResource(101, "Custom", "C:\temp\temp.tmp")
    Dim sFileName As String
    
    'Check if the TempFile Name has been specified
    If IsMissing(TempFile) Then
        'Create a temp file name such as "~res1234.tmp"
        GetTempFile "", "~rs", 0, sFileName
    Else
        'Use the specified temp file name
        sFileName = TempFile
    End If
    
    'Save the resource item to disk
    If SaveResItemToDisk(ResourceID, sResourceType, sFileName) = 0 Then
    
        'Return the picture
        Set LoadPictureResource = LoadPicture(sFileName)
        
        'Delete the temp file
        Kill sFileName
    End If
    
End Function

Public Function SaveResItemToDisk( _
            ByVal iResourceNum As Variant, _
            ByVal sResourceType As String, _
            ByVal sDestFileName As String _
            ) As Long
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    
    'Example Call:
    ' iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")
    
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    
    On Error GoTo SaveResItemToDisk_err
    
    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    
    'Get Free File Handle
    iFileNumOut = FreeFile
    
    'Open the output file
    Open sDestFileName For Binary Access Write As #iFileNumOut
        
        'Write the resource to the file
        Put #iFileNumOut, , bytResourceData
    
    'Close the file
    Close #iFileNumOut
    
    'Return 0 for success
    SaveResItemToDisk = 0
    
    Exit Function
SaveResItemToDisk_err:
    'Return error number
    SaveResItemToDisk = err.Number
End Function

Public Function GetTempFile( _
    ByVal strDestPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Integer, _
    ByRef lpTempFileName As String _
    ) As Boolean
    '==========================================================================
    ' Get a temporary filename for a specified drive and filename prefix
    ' PARAMETERS:
    '   strDestPath - Location where temporary file will be created.  If this
    '                 is an empty string, then the location specified by the
    '                 tmp or temp environment variable is used.
    '   lpPrefixString - First three characters of this string will be part of
    '                    temporary file name returned.
    '   wUnique - Set to 0 to create unique filename.  Can also set to integer,
    '             in which case temp file name is returned with that integer
    '             as part of the name.
    '   lpTempFilename - Temporary file name is returned as this variable.
    ' RETURN:
    '   True if function succeeds; false otherwise
    '==========================================================================
    
    If strDestPath = "" Then
        ' No destination was specified, use the temp directory.
        strDestPath = String(255, vbNullChar)
        If GetTempPath(255, strDestPath) = 0 Then
            GetTempFile = False
            Exit Function
        End If
    End If
    lpTempFileName = String(255, vbNullChar)
    GetTempFile = GetTempFileName(strDestPath, lpPrefixString, wUnique, lpTempFileName) > 0
    lpTempFileName = StripTerminator(lpTempFileName)
End Function


Public Function StripTerminator(ByVal strString As String) As String
    '==========================================================
    ' Returns a string without any zero terminator.  Typically,
    ' this was a string returned by a Windows API call.
    '
    ' IN: [strString] - String to remove terminator from
    '
    ' Returns: The value of the string passed in minus any
    '          terminating zero.
    '==========================================================
    
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'------------------------------------------------------------------------------------
' 리소스 관련 함수 끝
'------------------------------------------------------------------------------------
