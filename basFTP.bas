Attribute VB_Name = "basFTP"
Option Explicit
Option Base 0

Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

Public Const MAX_PATH = 260
Public Const NO_ERROR = 0
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000


Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Public Const ERROR_NO_MORE_FILES = 18

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean



Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const INTERNET_FLAG_PASSIVE = &H8000000

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_MULTIPART = &H200000

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetWriteFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal sBuffer As String, _
        ByVal lNumberOfBytesToRead As Long, _
        lNumberOfBytesRead As Long) As Integer

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFileName As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean

Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer

Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000



Public Enum TranType
    emUpload = 0&
    emDownload = 1&
End Enum

Public Enum emmMinGAnSa
    emKBP = 1
    emKIS = 2
    emNICE = 3
End Enum

Public Enum emmFileGubun
    emISSUE = 1
    emBOND = 2
    emCDCP = 3
    emYTM = 4
    emWarrant = 5
End Enum

Public Const FL_SUCCESS = 0
Public Const FL_ALREADY = 1
Public Const FL_NOEXIST = 2
Public Const FL_ERROR = 3
Public Const FL_KILL = 4


Function FTP_FileCopy(ByVal CopyMode As TranType, _
    IP As String, _
    ID As String, _
    PWD As String, _
    szFileRemote As String, _
    szFileLocal As String, _
    Optional force As Boolean _
    ) As Long

'변수선언

On Error GoTo errHandler

Dim dwType As Long
Dim nFlag As Long
Dim hOpen As Long
Dim hConnection As Long
Dim bActiveSession As Boolean
Dim bRet As Boolean

FTP_FileCopy = FL_SUCCESS

dwType = FTP_TRANSFER_TYPE_BINARY
nFlag = INTERNET_FLAG_PASSIVE

'인터넷 세션 생성
hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

'사이트 연결

hConnection = InternetConnect(hOpen, IP, INTERNET_INVALID_PORT_NUMBER, _
        ID, PWD, INTERNET_SERVICE_FTP, nFlag, 0)
        If hConnection = 0 Then
            '연결실패
            FTP_FileCopy = FL_ERROR
            Exit Function
        Else
            '연결성공
            bActiveSession = True
        End If
If CopyMode = emUpload Then
'파일 보내기
        If isFile(szFileLocal) Then
            If force Or Not isFtpFile(hConnection, szFileRemote) Then     '파일이 있든없든 강제로 재전송하기 옵션이거나 송신한파일이 없으면
                Screen.MousePointer = vbHourglass
                bRet = FtpPutFile(hConnection, szFileLocal, szFileRemote, _
                    dwType, 0)
                Screen.MousePointer = vbDefault
                If bRet = False Then
                    '파일보내기 실패
                    FTP_FileCopy = err.LastDllError
                    Exit Function
                Else
                    '파일보내기 성공
'                    If FileCompare(hConnection, szFileRemote, FileLen(szFileLocal)) Then
                        FTP_FileCopy = FL_SUCCESS
'                    Else
'                        '받았던 파일 지움
''                        LogPut szfilelocla & "파일 Size가 틀립니다. 나중에 다시 받도록하겠습니다."
'                        Kill (szFileLocal)
'                        FTP_FileCopy = FL_KILL
'                    End If
                End If
            Else
                    FTP_FileCopy = FL_ALREADY
            End If
        Else
                    FTP_FileCopy = FL_NOEXIST
        End If
Else
'파일 받기

        If isFtpFile(hConnection, szFileRemote) Then
            If force Or Not isFile(szFileLocal) Then '강제로 파일받기를 하는경우 혹은 파일이 없으면
                Screen.MousePointer = vbHourglass
                bRet = FtpGetFile(hConnection, szFileRemote, szFileLocal, False, _
                   INTERNET_FLAG_RELOAD, dwType, 0)
                Screen.MousePointer = vbDefault
                If bRet = False Then
                    '파일받기 실패
                    '?err.LastDllError = 12003  에러면 파일없음
                    FTP_FileCopy = err.LastDllError
                    Exit Function
                Else
                    '파일받기 성공
                     
'                    If FileVerify(hConnection, szFileRemote, FileLen(szFileLocal)) Then
                        FTP_FileCopy = FL_SUCCESS
'                    Else
'                        '받았던 파일 지움
''                        LogPut szfilelocla & "파일 Size가 틀립니다. 나중에 다시 받도록하겠습니다."
'                        Kill (szFileLocal)
'                        FTP_FileCopy = FL_KILL
'                    End If
                End If
            Else
                    FTP_FileCopy = FL_ALREADY
            End If
        Else
                    FTP_FileCopy = FL_NOEXIST
        End If
        
End If

'사이트 끊기
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0

'인터넷 세션 소멸
If hOpen <> 0 Then InternetCloseHandle (hOpen)

Exit Function

errHandler:
 
FTP_FileCopy = err.LastDllError

End Function

Function FTP_FileDelete( _
    IP As String, _
    ID As String, _
    PWD As String, _
    szFileRemote As String _
    ) As Long

'변수선언

On Error GoTo errHandler

Dim dwType As Long
Dim nFlag As Long
Dim hOpen As Long
Dim hConnection As Long
Dim bActiveSession As Boolean
Dim bRet As Boolean

FTP_FileDelete = FL_SUCCESS

dwType = FTP_TRANSFER_TYPE_BINARY
nFlag = INTERNET_FLAG_PASSIVE

'인터넷 세션 생성
hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

'사이트 연결
hConnection = InternetConnect(hOpen, IP, INTERNET_INVALID_PORT_NUMBER, _
        ID, PWD, INTERNET_SERVICE_FTP, nFlag, 0)
        If hConnection = 0 Then
            '연결실패
            FTP_FileDelete = FL_ERROR
            Exit Function
        Else
            '연결성공
            bActiveSession = True
        End If
'파일 삭제

bRet = FtpDeleteFile(hConnection, szFileRemote)

    If bRet = False Then
    '파일삭제 실패
    'err.LastDllError = 12003  에러면 파일없음
        FTP_FileDelete = err.LastDllError
        'Exit Function
    Else
    '파일삭제 성공
    
    '                    If FileVerify(hConnection, szFileRemote, FileLen(szFileLocal)) Then
        FTP_FileDelete = FL_SUCCESS
    '                    Else
    '                        '받았던 파일 지움
    ''                        LogPut szfilelocla & "파일 Size가 틀립니다. 나중에 다시 받도록하겠습니다."
    '                        Kill (szFileLocal)
    '                        FTP_FileCopy = FL_KILL
    '                    End If
    End If



'사이트 끊기
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0

'인터넷 세션 소멸
If hOpen <> 0 Then InternetCloseHandle (hOpen)

Exit Function

errHandler:
 
FTP_FileDelete = err.LastDllError

End Function


'Function InterNet_ErrorOut(dError As Long, szCallFunction As String) As String
'    Dim dwIntError As Long, dwLength As Long
'    Dim strBuffer As String
'    If dError = ERROR_INTERNET_EXTENDED_ERROR Then
'        InternetGetLastResponseInfo dwIntError, vbNullString, dwLength
'        strBuffer = String(dwLength + 1, 0)
'        InternetGetLastResponseInfo dwIntError, strBuffer, dwLength
'        ErrorOut = szCallFunction & " Extd Err: " & dwIntError & " " & strBuffer
'        Debug.Print "인터넷 오류번호:", dwIntError
'        Debug.Print "인터넷 오류내용:", strBuffer
'    End If
'End Function

'Function IDEMODE() As Boolean
'On Error GoTo errHandler
'    Debug.Print 1 / 0
'Exit Function
'errHandler:
'IDEMODE = True
'End Function

Private Function isFtpFile(hConnection As Long, ftp_filename As String) As Boolean

    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long
    Dim nLastError As Long
    
    hFind = FtpFindFirstFile(hConnection, ftp_filename, pData, 0, 0)
    
    nLastError = err.LastDllError
    
    If (hFind = 0 And nLastError = ERROR_NO_MORE_FILES) Then
        isFtpFile = False
    ElseIf hFind <> 0 Then
        isFtpFile = True
    ElseIf nLastError = 12003& Then
        isFtpFile = False
    Else
        Debug.Assert False '다른 원인에 의한 에러
    End If
    
'    If isFtpFile And ftp_filename <> StripNulls(pData.cFileName) Then
'
'        Do
'            Call InternetFindNextFile(hFind, pData)
'        Loop While ftp_filename <> Trim(StripNulls(pData.cFileName))
'        Debug.Print pData.nFileSizeLow
'    End If
    
    InternetCloseHandle (hFind)
    
End Function
Public Function GetFTPFileSize(hConnection As Long, ftp_filename As String) As Long

    Dim isFtpFile As Boolean
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long
    Dim nLastError As Long
    
    hFind = FtpFindFirstFile(hConnection, ftp_filename, pData, 0, 0)
    
    nLastError = err.LastDllError
    
    If (hFind = 0 And nLastError = ERROR_NO_MORE_FILES) Then
        isFtpFile = False
    ElseIf hFind <> 0 Then
        isFtpFile = True
    ElseIf nLastError = 12003& Then
        isFtpFile = False
    Else
        Debug.Assert False '다른 원인에 의한 에러
    End If
    
    If isFtpFile And ftp_filename <> StripNulls(pData.cFileName) Then
        
        Do
            Call InternetFindNextFile(hFind, pData)
        Loop While ftp_filename <> Trim(StripNulls(pData.cFileName))
        GetFTPFileSize = pData.nFileSizeLow
    ElseIf isFtpFile Then
        GetFTPFileSize = pData.nFileSizeLow
    End If
    
    InternetCloseHandle (hFind)
    
End Function
Private Function FileCompare(hConnection As Long, ftp_filename As String, FilesizeLocal As Long) As Boolean
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long
    Dim nLastError As Long
    
    hFind = FtpFindFirstFile(hConnection, ftp_filename, pData, 0, 0)
    If pData.nFileSizeLow = FilesizeLocal Then
        FileCompare = True 'File 크기 같음
    Else
        FileCompare = False ' File 다름
    End If
    
End Function


Public Function isFile(strfilefullNM) As Boolean

    Dim oFileSystem As Object
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
   
    isFile = oFileSystem.FileExists(strfilefullNM)
    Set oFileSystem = Nothing
    
End Function

Public Function isFolder(strFolderfullNM) As Boolean

    Dim oFileSystem As Object
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
   
    isFolder = oFileSystem.FolderExists(strFolderfullNM)
    Set oFileSystem = Nothing
    
End Function


Public Function DeleteFile(strfilefullNM) As Boolean
    Dim oFileSystem As Object
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    If oFileSystem.FileExists(strfilefullNM) Then
    DeleteFile = oFileSystem.DeleteFile(strfilefullNM)
    End If
    Set oFileSystem = Nothing
End Function

Public Function GetTempDirW() As String
    
    Dim tempDirW As String
    
    tempDirW = Environ$("temp")
    If tempDirW = "" Then
        tempDirW = App.Path
    End If
    If Not isFolder(tempDirW) Then
        tempDirW = "C:\"
    End If
    If Right(tempDirW, 1) <> "\" Then
        tempDirW = tempDirW & "\"
    End If

    GetTempDirW = tempDirW
End Function

Public Sub AlwaysOnTop(f As Form, OnTop As Boolean)
    'hwndInsertAfter values
    Const HWND_TOP = 0
    Const HWND_BOTTOM = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    'wFlags values
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER = &H4
    Const SWP_NOREDRAW = &H8
    Const SWP_NOACTIVATE = &H10
    Const SWP_FRAMECHANGED = &H20           'The frame changed: send WM_NCCALCSIZE
    Const SWP_SHOWWINDOW = &H40
    Const SWP_HIDEWINDOW = &H80
    Const SWP_NOCOPYBITS = &H100
    Const SWP_NOOWNERZORDER = &H200    'Don't do owner Z ordering
    Const SWP_DRAWFRAME = SWP_FRAMECHANGED
    Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

    If OnTop = True Then
        'Turn on the TopMost attribute.
        SetWindowPos f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    ElseIf OnTop = False Then
        'Turn off the TopMost attribute.
        SetWindowPos f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub

Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function

