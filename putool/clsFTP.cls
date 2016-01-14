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

Option Explicit

Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private szErrorMessage As String

'레지스트리관련 API
'/*************************************************/
'/* Specify constants to registry data types.     */
'/* These are declared Public for outside module  */
'/* usage in the GetAppRegValue() function.       */
'/*************************************************/
 Const REG_NONE = 0
 Const REG_SZ = 1
 Const REG_EXPAND_SZ = 2
 Const REG_BINARY = 3
 Const REG_DWORD = 4
 Const REG_DWORD_LITTLE_ENDIAN = REG_DWORD '4
 Const REG_DWORD_BIG_ENDIAN = 5
 Const REG_LINK = 6
 Const REG_MULTI_SZ = 7
 Const REG_RESOURCE_LIST = 8
 Const INIFILE = "/POCKETQUIZ.INI"

Private Const BUFFERSIZE = 100

Private dwType As Long


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' 레지스트리 롱형 키값 쿼리..
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Long, _
    lpcbData As Long _
) As Long

' 레지스트리 널형 키값 쿼리..
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As Long, _
    lpcbData As Long _
) As Long


' 레지스트리 문자열 키값 쿼리..
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As String, _
    lpcbData As Long _
) As Long

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259


Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006


Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000



Private m_phonecount As Integer

Private names(1 To 10) As String
Private phones(1 To 10) As String

Private Const KEY_ALL_ACCESS = _
    ((STANDARD_RIGHTS_ALL Or _
    KEY_QUERY_VALUE Or _
    KEY_SET_VALUE Or _
    KEY_CREATE_SUB_KEY Or _
    KEY_ENUMERATE_SUB_KEYS Or _
    KEY_NOTIFY Or KEY_CREATE_LINK) And _
    (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&

'프로파일 관련 API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'FTP 관련 API
Private Declare Function WSAStartup Lib "wsock32" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long

Private Declare Function inet_addr Lib "wsock32" _
   (ByVal s As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
   

Private Const IP_SUCCESS As Long = 0
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const WS_VERSION_REQD As Long = &H101
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const PING_TIMEOUT As Long = 500

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

'사용자정의
Private Tel1(1 To 3) As String
Private Tel2(1 To 3) As String 'KBP =1 , KIS =2 , NICE = 3
'===============================================================================

'속성값을 가질 지역 변수
Private mvarGateWay1 As String '지역 복사
Private mvarGateWay2 As Variant '지역 복사
Private mvarStatus As Integer '지역 복사
Private mvarisOutNet As Boolean '지역 복사
Private mvarisInNet As Boolean '지역 복사
Private mvarDataConfirm As Boolean '지역 복사

'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent changedcomplete[(arg1, arg2, ... , argn)]
'Public Event changedcomplete(ByRef isOutterMang As Integer) '1:내부, 2:외부, 3:On Change

Const conKBPIP  As String = "211.233.42.100"
Const conKBPUID As String = "hncb"
Const conKBPPass As String = "hncb3344"


Const conKISIP  As String = "211.41.63.3"
Const conKISUID As String = "kmbank"
Const conKISPass As String = "bankkm"


Const conNICEIP As String = "211.193.45.2"
Const conNICEUID As String = "kmb"
Const conNICEPass As String = "kmb"

'Const ipchangerEXE As String = "C:\Program Files\Fnet\IPChanger-H\IPChanger.exe"
'
'Const GateWay1 As String = "172.17.23.1" '내부망게이트웨이
'Const GateWay2 As String = "172.17.23.21" '외부망게이트웨이

Dim ipchangerEXE As String ' = "C:\Program Files\Fnet\IPChanger-H\IPChanger.exe"

Dim GateWay1 As String  '= "172.17.23.1" '내부망게이트웨이
Dim GateWay2 As String  '= "172.17.23.21" '외부망게이트웨이

'Public WithEvents Profile As clsProfile    '환경변수 컬렉션

Public Event FileTransferProgress(ByVal lCurrentBytes As Long, ByVal lTotalBytes As Long)


Public Function SetOutNet() As Boolean


   Dim testData As String
   Dim anyOk As Boolean
    
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long

   Dim success1 As Long
   Dim success2 As Long
    
   Dim saveTime  As Double
   
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
     
     saveTime = CDbl(Now) * 100000
     
      Do Until anyOk
'        DoEvents
        success1 = Ping((GateWay1), testData, ECHO)
        success2 = Ping((GateWay2), testData, ECHO)
        
        If success1 = IP_SUCCESS And success2 = IP_SUCCESS Then
            anyOk = True
            SetOutNet = True
            Exit Do
        End If
        
        
        If success1 = IP_SUCCESS And success2 <> IP_SUCCESS Then
          'anyOk = True
          ChangeMang
          success1 = Ping((GateWay1), testData, ECHO)
          success2 = Ping((GateWay2), testData, ECHO)
        End If
        
        
        If success2 = IP_SUCCESS And success1 <> IP_SUCCESS Then
          anyOk = True
          SetOutNet = True
        End If
        
        
        If CDbl(Now) * 100000 - saveTime > 120 Then
            Exit Do
        End If
        
      Loop
      
      SocketsCleanup
      
   Else
   
    SetOutNet = False
   End If

End Function

Public Function SetInNet() As Boolean

   Dim testData As String
   Dim anyOk As Boolean
    
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long

   Dim success1 As Long
   Dim success2 As Long
   
   Dim saveTime  As Double
    
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
     
     saveTime = CDbl(Now) * 100000
     
      Do Until anyOk
        success1 = Ping((GateWay1), testData, ECHO)
        success2 = Ping((GateWay2), testData, ECHO)
        
        
        If success1 = IP_SUCCESS And success2 = IP_SUCCESS Then
            anyOk = True
            SetInNet = True
            Exit Do
        End If
        
        If success1 = IP_SUCCESS And success2 <> IP_SUCCESS Then
          anyOk = True
          SetInNet = True
        End If
        
        If success2 = IP_SUCCESS And success1 <> IP_SUCCESS Then
          'anyOk = True
          ChangeMang
        End If
        
        If CDbl(Now) * 100000 - saveTime > 120 Then
            Exit Do
        End If
        
      Loop
      
      SocketsCleanup
      
   Else
   
    SetInNet = False
   End If
End Function


'Public Function GetRemoteFile(curDate, MinGanSa As emmMinGAnSa, FileGubn As emmFileGubun) As String
'    Dim stryyyy
'    Dim strmm
'    Dim strdd
'    stryyyy = Year(curDate)
'    strmm = Right("0" & Month(curDate), 2)
'    strdd = Right("0" & Day(curDate), 2)
'
'    Select Case FileGubn
'    Case emISSUE '보유CDCP
'    '보유CDCP
'    'KBP  :/hncb/KMB20030307.txt
'    'KIS  :/kmbank/2003/03/issue07.txt
'    'NICE :KMB20030307.txt
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetRemoteFile = "/hncb/KMB" & stryyyy & strmm & strdd & ".txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetRemoteFile = "/kmbank/" & stryyyy & "/" & strmm & "/issue" & strdd & ".txt"
'        Else
'        'NICE
'            GetRemoteFile = "KMB" & stryyyy & strmm & strdd & ".txt"
'        End If
'
'    Case emBOND '채권
'
'    '채권
'    'KBP  :/hncb/KBP20030307.txt
'    'KIS  :/kmbank/2003/03/price07.txt
'    'NICE :NIC20030307.txt
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetRemoteFile = "/hncb/KBP" & stryyyy & strmm & strdd & ".txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetRemoteFile = "/kmbank/" & stryyyy & "/" & strmm & "/price" & strdd & ".txt"
'        Else
'        'NICE
'            GetRemoteFile = "NIC" & stryyyy & strmm & strdd & ".txt"
'        End If
'
'    Case emCDCP 'CDCP
'    'CDCP
'    'KBP  :/hncb/KBP20030307C.txt
'    'KIS  :/kmbank/2003/03/cdcp07.txt
'    'NICE :NIC20030307C.txt
'        If MinGanSa = emKBP Then
'        'KBP
'            GetRemoteFile = "/hncb/KBP" & stryyyy & strmm & strdd & "C.txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetRemoteFile = "/kmbank/" & stryyyy & "/" & strmm & "/cdcp" & strdd & ".txt"
'        Else
'        'NICE
'            GetRemoteFile = "NIC" & stryyyy & strmm & strdd & "C.txt"
'        End If
'
'    Case emYTM 'YTM
'    'YTM
'    'KBP  :/hncb/KBP20030307YTM.txt
'    'KIS  :/kmbank/2003/03/ytm07.txt
'    'NICE :NIC20030307YTM.txt
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetRemoteFile = "/hncb/KBP" & stryyyy & strmm & strdd & "YTM.txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetRemoteFile = "/kmbank/" & stryyyy & "/" & strmm & "/ytm" & strdd & ".txt"
'        Else
'        'NICE
'            GetRemoteFile = "NIC" & stryyyy & strmm & strdd & "YTM.txt"
'        End If
'
'    Case emWarrant 'Warrant
'    'Warrant
'    'KBP  :/hncb/KBP20030307W.txt
'    'KIS  :/kmbank/2003/03/warrant07.txt
'    'NICE :NIC20030307W.txt
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetRemoteFile = "/hncb/KBP" & stryyyy & strmm & strdd & "W.txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetRemoteFile = "/kmbank/" & stryyyy & "/" & strmm & "/warrant" & strdd & ".txt"
'        Else
'        'NICE
'            GetRemoteFile = "NIC" & stryyyy & strmm & strdd & "W.txt"
'        End If
'
'    End Select
'
'Debug.Print "RemoteFile", GetRemoteFile
'
'End Function
'Public Function GetLocalFile(curDate, MinGanSa As emmMinGAnSa, FileGubn As emmFileGubun) As String
'    Dim strYMD
'
'    strYMD = Format(curDate, "YYYYMMDD")
'
'    Select Case FileGubn
'    Case emISSUE '보유CDCP
'    '보유CDCP
'    'KBP  :C:\민간사시가\보유cdcp송신\KMB20030307.txt
'    'KIS  :C:\민간사시가\보유cdcp송신\KMB20030307.txt
'    'NICE :C:\민간사시가\보유cdcp송신\KMB20030307.txt
'
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetLocalFile = "C:\민간사시가\보유cdcp송신\KMB" & strYMD & ".txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetLocalFile = "C:\민간사시가\보유cdcp송신\KMB" & strYMD & ".txt"
'        Else
'        'NICE
'            GetLocalFile = "C:\민간사시가\보유cdcp송신\KMB" & strYMD & ".txt"
'        End If
'
'    Case emBOND '채권
'
'    '채권
'    'KBP  :C:\민간사시가\시가수신\KBP\KBP20030307.txt
'    'KIS  :C:\민간사시가\시가수신\KIS\KIS20030307.txt
'    'NICE :C:\민간사시가\시가수신\NIC\NIC20030307.txt
'
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetLocalFile = "C:\민간사시가\시가수신\KBP\KBP" & strYMD & ".txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetLocalFile = "C:\민간사시가\시가수신\KIS\KIS" & strYMD & ".txt"
'        Else
'        'NICE
'            GetLocalFile = "C:\민간사시가\시가수신\NIC\NIC" & strYMD & ".txt"
'        End If
'
'    Case emCDCP 'CDCP
'    'CDCP
'    'KBP  :C:\민간사시가\시가수신\KBP\KBP20030307C.txt
'    'KIS  :C:\민간사시가\시가수신\KIS\KIS20030307C.txt
'    'NICE :C:\민간사시가\시가수신\NIC\NIC20030307C.txt
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetLocalFile = "C:\민간사시가\시가수신\KBP\KBP" & strYMD & "C.txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetLocalFile = "C:\민간사시가\시가수신\KIS\KIS" & strYMD & "C.txt"
'        Else
'        'NICE
'            GetLocalFile = "C:\민간사시가\시가수신\NIC\NIC" & strYMD & "C.txt"
'        End If
'
'    Case emYTM 'YTM
'    'YTM
'    'KBP  :C:\민간사시가\시가수신\KBP\KBP20030307YTM.txt
'    'KIS  :C:\민간사시가\시가수신\KIS\20030307YTM.txt
'    'NICE :C:\민간사시가\시가수신\NIC\NIC20030307YTM.txt
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetLocalFile = "C:\민간사시가\시가수신\KBP\KBP" & strYMD & "YTM.txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetLocalFile = "C:\민간사시가\시가수신\KIS\" & strYMD & "YTM.txt"
'        Else
'        'NICE
'            GetLocalFile = "C:\민간사시가\시가수신\NIC\NIC" & strYMD & "YTM.txt"
'        End If
'    Case emWarrant
'
'        If MinGanSa = emKBP Then
'        'KBP
'            GetLocalFile = "C:\민간사시가\시가수신\KBP\KBP" & strYMD & "W.txt"
'        ElseIf MinGanSa = emKIS Then
'        'KIS
'            GetLocalFile = "C:\민간사시가\시가수신\KIS\KIS" & strYMD & "W.txt"
'        Else
'        'NICE
'            GetLocalFile = "C:\민간사시가\시가수신\NIC\NIC" & strYMD & "W.txt"
'        End If
'
'    End Select
'Debug.Print "LocalFile", GetLocalFile
'
'End Function
'
'
Public Property Get KBPIP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    KBPIP = conKBPIP
End Property

Public Property Get KISIP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    KISIP = conKISIP
End Property

Public Property Get NICEIP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    NICEIP = conNICEIP
End Property

Public Property Get KBPUSERNM() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    KBPUSERNM = conKBPUID
End Property

Public Property Get KISUSERNM() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    KISUSERNM = conKISUID
End Property

Public Property Get NICEUSERNM() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    NICEUSERNM = conNICEUID
End Property

Public Property Get KBPPASS() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    KBPPASS = conKBPPass
End Property

Public Property Get KISPASS() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    KISPASS = conKISPass
End Property

Public Property Get NICEPASS() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.isInNet
    NICEPASS = conNICEPass
End Property


Private Sub ChangeMang()
    Call Shell(ipchangerEXE, vbHide)
End Sub

Private Function Ping(sAddress As String, _
                     sDataToSend As String, _
                     ECHO As ICMP_ECHO_REPLY) As Long

  'If Ping succeeds :
  '.RoundTripTime = time in ms for the ping to complete,
  '.Data is the data returned (NULL terminated)
  '.Address is the Ip address that actually replied
  '.DataSize is the size of the string in .Data
  '.Status will be 0
  '
  'If Ping fails .Status will be the error code
   
   Dim hPort As Long
   Dim dwAddress As Long
   
  'convert the address into a long representation
   dwAddress = inet_addr(sAddress)
   
  'if a valid address..
   If dwAddress <> INADDR_NONE Then
   
     'open a port
      hPort = IcmpCreateFile()
      
     'and if successful,
      If hPort Then
      
        'ping it.
         Call IcmpSendEcho(hPort, _
                           dwAddress, _
                           sDataToSend, _
                           Len(sDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

        'return the status as ping succes and close
         Ping = ECHO.status
         Call IcmpCloseHandle(hPort)
      
      End If
      
   Else:
        'the address format was probably invalid
         Ping = INADDR_NONE
         
   End If
  
End Function

Private Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


Public Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub

'Public Function GetTel(KiKwan As emmMinGAnSa, idx As Integer) As String
'    If idx = 1 Then
'        GetTel = Tel1(KiKwan)
'    Else
'        GetTel = Tel2(KiKwan)
'    End If
'End Function
'
'
Public Function LenH(ByVal strString As String) As Long
'용도: 한글과 영문이 혼합된 경우에 한글은 2바이트, 영문은 1바이트로 계산

    LenH = LenB(StrConv(strString, vbFromUnicode))
    
End Function

Public Function LeftH(ByVal strString As String, ByVal lngLength As Long) As String
'용도: 한글과 영문이 혼합된 경우에 한글은 2바이트, 영문은 1바이트로 계산
    
    LeftH = StrConv(LeftB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
    
End Function

Public Function RightH(ByVal strString As String, lngLength As Long) As String
'용도: 한글과 영문이 혼합된 경우에 한글은 2바이트, 영문은 1바이트로 계산

    RightH = StrConv(RightB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
    
End Function

Public Function MidH(ByVal strString As String, ByVal lngStart As Long, _
        Optional ByVal lngLength) As String
'용도: 한글과 영문이 혼합된 경우에 한글은 2바이트, 영문은 1바이트로 계산
    
    If IsMissing(lngLength) Then
        MidH = StrConv(MidB(StrConv(strString, vbFromUnicode), lngStart), _
            vbUnicode)
    Else
        MidH = StrConv(MidB(StrConv(strString, vbFromUnicode), lngStart, _
            lngLength), vbUnicode)
    End If
            
End Function

Public Function SearchH(ByVal strString As String, ByVal strCharacter As String) As Integer
'용도 : 한글과 영문이 혼합된 경우에 문자열에서 특정한 1글자의 위치를 리턴
    Dim lngCounter As Long
    
    lngCounter = 0
    For lngCounter = 1 To LenH(strString)
        If strCharacter = MidH(strString, lngCounter, 1) Then
            Exit For
        End If
    Next
    SearchH = lngCounter
End Function
Public Function ConvertStr(ByVal strString As String) As String
' 택스트 필드 에서 작은따옴표 처리하여 DB에 입력
' 한글도 1바이트 영문도 1바이트
    Dim lngCounter As Long
    Dim strReturn As String
    
    lngCounter = 0
    For lngCounter = 1 To Len(strString)
        If Mid(strString, lngCounter, 1) = "'" Then
            strReturn = strReturn + "''"
        Else
            strReturn = strReturn + Mid(strString, lngCounter, 1)
        End If
    Next
    ConvertStr = strReturn
End Function



Private Function GetKeyValue(ByVal selected_hkey As Long, ByVal key As String, ByVal subkey As String) As String
Dim hKey As Long
Dim value As String
Dim length As Long
Dim value_type As Long

    ' Open the key.
    If RegOpenKeyEx(selected_hkey, key, 0&, _
        KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS _
    Then
        GetKeyValue = "Error opening key."
        Exit Function
    End If

    ' Get the subkey's value.
    length = 256
    value = Space$(length)
    If RegQueryValueEx(hKey, subkey, _
        0&, value_type, ByVal value, length) _
            <> ERROR_SUCCESS _
    Then
        GetKeyValue = "Error getting subkey value."
    Else
        ' Remove the trailing null character.
        GetKeyValue = Left$(value, length - 1)
    End If
    
    ' Close the key.
    RegCloseKey hKey
End Function

''Public Function GetIP(KiKwan As emmMinGAnSa) As String
''Select Case KiKwan
''Case emKBP
''    GetIP = KBPIP
''Case emKIS
''    GetIP = KISIP
''Case emNICE
''    GetIP = NICEIP
''Case Else
''    Debug.Assert False
''End Select
''End Function
''
'Public Function GetUserNM(KiKwan As emmMinGAnSa) As String
'Select Case KiKwan
'Case emKBP
'    GetUserNM = KBPUSERNM
'Case emKIS
'    GetUserNM = KISUSERNM
'Case emNICE
'    GetUserNM = NICEUSERNM
'Case Else
'    Debug.Assert False
'End Select
'End Function
'
'
''
'

'Public Function GetPass(KiKwan As emmMinGAnSa) As String
'Select Case KiKwan
'Case emKBP
'    GetPass = KBPPASS
'Case emKIS
'    GetPass = KISPASS
'Case emNICE
'    GetPass = NICEPASS
'Case Else
'    Debug.Assert False
'End Select
'End Function


Public Property Get isDataConfirm() As Boolean
    isDataConfirm = mvarDataConfirm
End Property

Public Property Let isDataConfirm(ByVal vNewValue As Boolean)
    mvarDataConfirm = vNewValue
    If mvarDataConfirm Then
        WritePrivateProfileString "ENV", "DATACONFIRM", "True", App.Path & "/tmCL21101.ini"
    Else
        WritePrivateProfileString "ENV", "DATACONFIRM", "False", App.Path & "/tmCL21101.ini"
    End If
End Property

Public Function isInNet() As Boolean

   Dim testData As String
   Dim anyOk As Boolean
    
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long

   Dim success1 As Long
   
   Dim saveTime  As Double
    
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
    'ping the ip passing the address, text
    'to send, and the ECHO structure.
    
    saveTime = CDbl(Now) * 100000
    
    success1 = Ping((GateWay1), testData, ECHO)
    
    If success1 = IP_SUCCESS Then
        isInNet = True
    Else
        isInNet = False
    End If
    
    SocketsCleanup
    End If
End Function

Public Function isOutNet() As Boolean

   Dim testData As String
   Dim anyOk As Boolean
    
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long

   Dim success1 As Long
   
   Dim saveTime  As Double
    
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
    'ping the ip passing the address, text
    'to send, and the ECHO structure.
    
    saveTime = CDbl(Now) * 100000
    
    success1 = Ping((GateWay2), testData, ECHO)
    
    If success1 = IP_SUCCESS Then
        isOutNet = True
    Else
        isOutNet = False
    End If
    
    SocketsCleanup
    End If
End Function


Public Property Get getPhoneCNT() As Variant
    getPhoneCNT = m_phonecount
End Property


Public Function getName(i As Integer) As String
    getName = names(i)
End Function

Public Function getPhone(i As Integer) As String
    getPhone = phones(i)
End Function


' Example : GetRegValue HKEY_LOCAL_MACHINE, "SOFTWARE\RegTest\Key1", "Val1", value ' value 변수에 적용되어서 받음
Private Function GetRegValue(ByVal lhKey As Long, sKeyName As String, sValueName As String, Optional ByVal sDefaultValue As Variant = "") As Variant
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long            'handle of opened key
    Dim vValue As Variant       'setting of queried value

    lRetVal = RegOpenKeyEx(lhKey, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)

'    If vValue = Null Or vValue = Empty Then
    If lRetVal <> ERROR_NONE Then
        GetRegValue = sDefaultValue ' 디폴트
    Else
        Dim i%, J%
        J% = 0
        For i% = 1 To Len(vValue)
            If Asc(Mid(vValue, i%, 1)) < 0 Then J% = J% + 1
        Next i%
        If J% > 0 Then vValue = Left(vValue, Len(vValue) - 1)
        'Else: vValue = vValue
        'End If

        GetRegValue = vValue
    End If

    RegCloseKey (hKey)
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String

   On Error GoTo QueryValueExError

   ' Determine the size and type of data to be read
   lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
   If lrc <> ERROR_NONE Then Error 5

   Select Case lType
      ' For strings
      Case REG_SZ:
         sValue = String(cch, 0)
      
         lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
         If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch - 1)
         Else
            vValue = Empty
         End If
      ' For DWORDS
      Case REG_DWORD:
         lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
         If lrc = ERROR_NONE Then vValue = Hex(lValue)
         
         Dim lCount As Long, lcnt As Long
         Dim vtemp1 As String, vtemp2 As String
         
         For lCount = 1 To Len(vValue) / 2
            lcnt = lCount * 2 - 1
            vtemp1 = Mid(vValue, lcnt, 2)
            vtemp2 = vtemp2 & val("&h" & vtemp1) & "."
         Next
         vValue = Mid(vtemp2, 1, Len(vtemp2) - 1)
      Case Else
         'all other data types not supported
         lrc = -1
   End Select

QueryValueExExit:
   QueryValueEx = lrc
   Exit Function

QueryValueExError:
   Resume QueryValueExExit
End Function



Public Function FTPUploadFile(hConnection As Long, sLocal As String, sRemote As String) As Boolean
    Dim Data(BUFFERSIZE - 1) As Byte
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    Dim hFile As Long
    
    Sum = 0
    lBlock = 0
    sLocal = Trim(sLocal)
    sRemote = Trim(sRemote)
    
    If sLocal <> "" And sRemote <> "" Then
      hFile = FtpOpenFile(hConnection, sRemote, GENERIC_WRITE, dwType, 0)
      If hFile = 0 Then
          ErrorOut Err.LastDllError, "FtpOpenFile:PutFile"
          FTPUploadFile = False
          Exit Function
      End If
      
      Open sLocal For Binary Access Read As #1
      Size = LOF(1)
      For lBlock = 1 To Size \ BUFFERSIZE
          Get #1, , Data
          If (InternetWriteFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
              FTPUploadFile = False
              ErrorOut Err.LastDllError, "InternetWriteFile"
              Exit Function
          End If
          DoEvents
          Sum = Sum + BUFFERSIZE
          RaiseEvent FileTransferProgress(Sum, Size)
      Next lBlock
      
      Get #1, , Data
      If (InternetWriteFile(hFile, Data(0), Size Mod BUFFERSIZE, Written) = 0) Then
          FTPUploadFile = False
          ErrorOut Err.LastDllError, "InternetWriteFile2"
          Exit Function
      End If
      
      Sum = Sum + (Size Mod BUFFERSIZE)
      Size = Sum
      RaiseEvent FileTransferProgress(Sum, Size)

      Close #1
      InternetCloseHandle (hFile)
      FTPUploadFile = True
   End If

End Function

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
    Dim dwRet As Long
    Dim dwTemp As Long
    Dim szString As String * 2048
    dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                      GetModuleHandle("wininet.dll"), dwError, 0, _
                      szString, 256, 0)
    szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
    If (dwError = 12003) Then
        ' Extended error information was returned
        dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
        szErrorMessage = szString
    End If
End Sub


Public Function FTPDownloadFile(ip As String, id As String, pwd As String, ByVal sRemote As String, sLocal As String) As Long
    Dim Data(BUFFERSIZE - 1) As Byte ' array of 100 elements 0 to 99
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    Dim hFile As Long
    
Dim dwType As Long
Dim nFlag As Long
Dim hOpen As Long
Dim hConnection As Long
Dim bActiveSession As Boolean
Dim sBuffer()             As Byte
Dim i As Integer
    
'Dim bRet As Boolean

'FTP_FileCopy = FL_SUCCESS

dwType = FTP_TRANSFER_TYPE_BINARY
nFlag = INTERNET_FLAG_PASSIVE

'인터넷 세션 생성
hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

'사이트 연결

hConnection = InternetConnect(hOpen, ip, INTERNET_INVALID_PORT_NUMBER, _
        id, pwd, INTERNET_SERVICE_FTP, nFlag, 0)
        If hConnection = 0 Then
            '연결실패
            FTPDownloadFile = FL_ERROR
'            FTPDownloadFile = False
            Exit Function
        Else
            '연결성공
            bActiveSession = True
        End If
        
        
    
    FTPDownloadFile = FL_SUCCESS
            
    Sum = 0
    lBlock = 0
    
    sLocal = Trim(sLocal)
    sRemote = Trim(sRemote)
    
    If sLocal <> "" And sRemote <> "" Then
      Size = GetFTPFileSize(hConnection, sRemote)
      DoEvents
      If Size > 0 Then
          hFile = FtpOpenFile(hConnection, sRemote, GENERIC_READ, dwType, 0)
          DoEvents
          If hFile = 0 Then
              ErrorOut Err.LastDllError, "FtpOpenFile:GetFile"
              Exit Function
          End If
          
          Dim sReadBuffer(2048)         As Byte ' *
          Dim lNumberOfBytesRead  As Long
        Dim bRet                As Boolean
        Dim bDoLoop             As Boolean
        sBuffer = ""
            Sum = 0
            bDoLoop = True
'            ReDim sBuffer(Size)
Open sLocal For Binary Access Write As #1
'Seek #1, 1
            While bDoLoop
'                sReadBuffer(0) = 0 'vbNullString ' String(UBound(sReadBuffer), vbNullString)
                bRet = InternetReadFileBinary(hFile, sReadBuffer(0), UBound(sReadBuffer), lNumberOfBytesRead)
'                sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)

                ReDim sBuffer(lNumberOfBytesRead)
                
                For i = 0 To lNumberOfBytesRead
                    sBuffer(i) = sReadBuffer(i)
                Next
                'sBuffer(0) = sReadBuffer
'                CopyMemory2 sBuffer, sReadBuffer, lNumberOfBytesRead  ' sBuffer = sReadBuffer
                Seek #1, Sum + 1
                Put #1, , sBuffer
                If bRet Then
                    
                    'CopyMemory2
                    Sum = Sum + lNumberOfBytesRead
                    RaiseEvent FileTransferProgress(Sum, Size)
                    
                    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                Else
                    bDoLoop = False
                End If
            Wend
    
    
'    Put #1, , sBuffer
    Close #1
    
          
          'InternetCloseHandle (hFile)
          DoEvents
          FTPDownloadFile = FL_SUCCESS
          DoEvents
      End If
   End If
   
   
   If hFile <> 0 Then InternetCloseHandle hFile
   '사이트 끊기
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0

'인터넷 세션 소멸
If hOpen <> 0 Then InternetCloseHandle (hOpen)

   
   
   
End Function


