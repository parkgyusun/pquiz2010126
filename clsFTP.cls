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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Option Base 0

Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800

Private Declare Function GetModuleHandle Lib "KERNEL32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
Private Declare Function FormatMessage Lib "KERNEL32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
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

Private Const BUFFERSIZE = 255

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
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "KERNEL32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'FTP 관련 API
Private Declare Function WSAStartup Lib "wsock32" _
   (ByVal wVersionRequired As Long, _
    lpWSAData As WSADATA) As Long

Private Declare Function inet_addr Lib "wsock32" _
   (ByVal S As String) As Long
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
   wversion As Integer
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
    Flags           As Byte
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
Public Event changedcomplete(ByRef isOutterMang As Integer) '1:내부, 2:외부, 3:On Change

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

Public WithEvents Profile As clsProfile    '환경변수 컬렉션
Attribute Profile.VB_VarHelpID = -1

Public Event FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)


Public Function SetOutNet() As Boolean


   Dim testData As String
   Dim anyOk As Boolean
    
   Dim ECHO As ICMP_ECHO_REPLY
   Dim Pos As Long
   Dim success As Long

   Dim success1 As Long
   Dim success2 As Long
    
   Dim saveTime  As Double
   
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
     
     saveTime = CDbl(Now) * 100000#
     
      Do Until anyOk
        DoEvents
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
        
        
        If CDbl(Now) * 100000# - saveTime > 120 Then
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
   Dim Pos As Long
   Dim success As Long

   Dim success1 As Long
   Dim success2 As Long
   
   Dim saveTime  As Double
    
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
     
     saveTime = CDbl(Now) * 100000#
     
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
        
        If CDbl(Now) * 100000# - saveTime > 120 Then
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

Public Function getINIFILE() As String
    getINIFILE = INIFILE
End Function

Public Sub Class_Initialize()


'Dim i As Integer

Dim Ret As Long
Dim RetStr As String * 255
Dim SelectedHkey As Long
'
'SelectedHkey = HKEY_LOCAL_MACHINE
'
'ipchangerEXE = GetKeyValue(SelectedHkey, "SOFTWARE\Fnet\IPChanger-H", "InstalledPath")
'
'If Right(ipchangerEXE, 1) = "\" Then
'    If isFile(ipchangerEXE & "LDIPChng.exe") Then
'        ipchangerEXE = ipchangerEXE & "LDIPChng.exe"
'    Else
'        ipchangerEXE = ipchangerEXE & "IPChanger.exe"
'    End If
'Else
'    If isFile(ipchangerEXE & "\LDIPChng.exe") Then
'        ipchangerEXE = ipchangerEXE & "\LDIPChng.exe"
'    Else
'        ipchangerEXE = ipchangerEXE & "\IPChanger.exe"
'    End If
'End If
'
'If Right(ipchangerEXE, 12) = "LDIPChng.exe" Then
'    GateWay1 = GetRegValue(SelectedHkey, "SOFTWARE\Fnet\IPChanger-H\IPCfg\info0", "GateWay")
'    GateWay2 = GetRegValue(SelectedHkey, "SOFTWARE\Fnet\IPChanger-H\IPCfg\info1", "GateWay")
'Else
'    GateWay1 = GetKeyValue(SelectedHkey, "SOFTWARE\Fnet\IPChanger-H", "GateWay1")
'    GateWay2 = GetKeyValue(SelectedHkey, "SOFTWARE\Fnet\IPChanger-H", "GateWay2")
'End If
'

'==============================================================================
'       환경변수 Set
'==============================================================================
Dim retVal As String
Dim retVal1 As String
Dim retVal2 As String

Set Profile = New clsProfile

If gUserid <> "" Then

'1
Ret = GetPrivateProfileString(gUserid, "fontsize", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "fontsize", "15", App.Path & INIFILE
    Profile.FontSize = 15
Else
    Profile.FontSize = retVal
End If


'2
Ret = GetPrivateProfileString(gUserid, "bSetTimeOutStudy", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bSetTimeOutStudy", "False", App.Path & INIFILE
    Profile.bSetTimeOutStudy = False
Else
    Profile.bSetTimeOutStudy = CBool(retVal)
End If


'3
Ret = GetPrivateProfileString(gUserid, "TimeOutSecStudy", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "TimeOutSecStudy", "60", App.Path & INIFILE
    Profile.TimeOutSecStudy = 60
Else
    Profile.TimeOutSecStudy = retVal
End If

'4
Ret = GetPrivateProfileString(gUserid, "TimeOutSec", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "TimeOutSec", "60", App.Path & INIFILE
    Profile.TimeOutSec = 60
Else
    Profile.TimeOutSec = retVal
End If

'5
Ret = GetPrivateProfileString(gUserid, "bSetTimeOut", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bSetTimeOut", "True", App.Path & INIFILE
    Profile.bSetTimeOut = CBool("True")
Else
    Profile.bSetTimeOut = CBool(retVal)
End If

'6
Ret = GetPrivateProfileString(gUserid, "bStreateOutCntCheck", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bStreateOutCntCheck", "True", App.Path & INIFILE
    Profile.bStreateOutCntCheck = CBool("True")
Else
    Profile.bStreateOutCntCheck = CBool(retVal)
End If

'7
Ret = GetPrivateProfileString(gUserid, "CntOfStreatOutSetting", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "CntOfStreatOutSetting", "1", App.Path & INIFILE
    Profile.CntOfStreatOutSetting = 1
Else
    Profile.CntOfStreatOutSetting = retVal
End If

'8
Ret = GetPrivateProfileString(gUserid, "bAlarm1", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bAlarm1", "False", App.Path & INIFILE
    Profile.bAlarm1 = False
Else
    Profile.bAlarm1 = retVal
End If

'9
Ret = GetPrivateProfileString(gUserid, "bAlarm2", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bAlarm2", "False", App.Path & INIFILE
    Profile.bAlarm2 = False
Else
    Profile.bAlarm2 = retVal
End If

'10
Ret = GetPrivateProfileString(gUserid, "bAlarm5", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bAlarm5", "False", App.Path & INIFILE
    Profile.bAlarm5 = False
Else
    Profile.bAlarm5 = retVal
End If

'11
Ret = GetPrivateProfileString(gUserid, "CntOfAlarm1", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "CntOfAlarm1", "10", App.Path & INIFILE
    Profile.CntOfAlarm1 = 10
Else
    Profile.CntOfAlarm1 = retVal
End If

'12
Ret = GetPrivateProfileString(gUserid, "CntOfAlarm2", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "CntOfAlarm2", "20", App.Path & INIFILE
    Profile.CntOfAlarm2 = 20
Else
    Profile.CntOfAlarm2 = retVal
End If

'13
Ret = GetPrivateProfileString(gUserid, "CntOfAlarm5", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "CntOfAlarm5", "30", App.Path & INIFILE
    Profile.CntOfAlarm5 = 30
Else
    Profile.CntOfAlarm5 = retVal
End If

'14
Ret = GetPrivateProfileString(gUserid, "fontBold", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "fontBold", "False", App.Path & INIFILE
    Profile.FontBold = False
Else
    Profile.FontBold = CBool(retVal)
End If


'15
Ret = GetPrivateProfileString(gUserid, "DelayTime1", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "DelayTime1", "10", App.Path & INIFILE
    Profile.DelayTime1 = 10
Else
    Profile.DelayTime1 = retVal
End If


'16
Ret = GetPrivateProfileString(gUserid, "DelayTime2", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "DelayTime2", "2", App.Path & INIFILE
    Profile.DelayTime2 = 2
Else
    Profile.DelayTime2 = retVal
End If

'17
Ret = GetPrivateProfileString(gUserid, "nansu", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "nansu", "2", App.Path & INIFILE
    Profile.nansu = 2
Else
    Profile.nansu = CLng(retVal)
End If

'18 기본값 false
Ret = GetPrivateProfileString(gUserid, "compactDB", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "compactDB", "False", App.Path & INIFILE
    Profile.compactDB = False
Else
    Profile.compactDB = CBool(retVal)
End If

'19'삭제시 물어봄
Ret = GetPrivateProfileString(gUserid, "setPoP1", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "setPoP1", "False", App.Path & INIFILE
    Profile.setPoP1 = False
Else
    Profile.setPoP1 = CBool(retVal)
End If

'20 경미한 에러에서 에러 메시지 보임 '
Ret = GetPrivateProfileString(gUserid, "setPoP2", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "setPoP2", "False", App.Path & INIFILE
    Profile.setPoP2 = False
Else
    Profile.setPoP2 = CBool(retVal)
End If

'20-1 닫을때 질문
Ret = GetPrivateProfileString(gUserid, "setPoP3", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "setPoP3", "False", App.Path & INIFILE
    Profile.setPoP3 = False
Else
    Profile.setPoP3 = CBool(retVal)
End If


'21
Ret = GetPrivateProfileString(gUserid, "ShowWhiteSpot", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "ShowWhiteSpot", "False", App.Path & INIFILE
    Profile.ShowWhiteSpot = False
Else
    Profile.ShowWhiteSpot = CBool(retVal)
End If

'22
Ret = GetPrivateProfileString(gUserid, "checkWhenResize", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "checkWhenResize", "False", App.Path & INIFILE
    Profile.checkWhenResize = False
Else
    Profile.checkWhenResize = CBool(retVal)
End If


'24
'fontName
Ret = GetPrivateProfileString(gUserid, "FontName", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "FontName", "", App.Path & INIFILE
    Profile.FontName = ""
Else
    Profile.FontName = retVal
End If

'25
'endSettingLogin
Ret = GetPrivateProfileString(gUserid, "endSettingLogin", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "endSettingLogin", "2", App.Path & INIFILE
    Profile.endSettingLogin = "2"
Else
    Profile.endSettingLogin = retVal
End If


'26
'endSettingLoginDefaultMin
Ret = GetPrivateProfileString(gUserid, "endSettingLoginDefaultMin", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "endSettingLoginDefaultMin", "60", App.Path & INIFILE
    Profile.endSettingLoginDefaultMin = "60"
Else
    Profile.endSettingLoginDefaultMin = retVal
End If


'27
'endSettingQuiz
Ret = GetPrivateProfileString(gUserid, "endSettingQuiz", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "endSettingQuiz", "2", App.Path & INIFILE
    Profile.endSettingQuiz = "2"
Else
    Profile.endSettingQuiz = retVal
End If


'28
'endSettingQuizDefaultMin
Ret = GetPrivateProfileString(gUserid, "endSettingQuizDefaultMin", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Or Not IsNumeric(retVal) Then
    WritePrivateProfileString gUserid, "endSettingQuizDefaultMin", "30", App.Path & INIFILE
    Profile.endSettingQuizDefaultMin = "30"
Else
    Profile.endSettingQuizDefaultMin = retVal
End If


'29
'endSettingSetMin
Ret = GetPrivateProfileString(gUserid, "endSettingSetMin", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "endSettingSetMin", "30", App.Path & INIFILE
    Profile.endSettingSetMin = "30"
Else
    Profile.endSettingSetMin = retVal
End If


'30
'endSettingQuestion
Ret = GetPrivateProfileString(gUserid, "endSettingQuestion", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "endSettingQuestion", "True", App.Path & INIFILE
    Profile.endSettingQuestion = True
Else
    Profile.endSettingQuestion = CBool(retVal)
End If


'31
'endSettingTimer
Ret = GetPrivateProfileString(gUserid, "endSettingTimer", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "endSettingTimer", "True", App.Path & INIFILE
    Profile.endSettingTimer = True
Else
    Profile.endSettingTimer = CBool(retVal)
End If

''32
''bTTSbeta (기본값 true임에 유의)
'Ret = GetPrivateProfileString(gUserid, "bTTSbeta", "", RetStr, 255, App.Path & INIFILE)
'retVal = LeftH(RetStr, Ret)
'
'If retVal = "" Then
'    WritePrivateProfileString gUserid, "bTTSbeta", "True", App.Path & INIFILE
'    Profile.bTTSbeta = True
'Else
'    Profile.bTTSbeta = retVal
'End If

'33
'bTTSuse (기본값 true임에 유의)
Ret = GetPrivateProfileString(gUserid, "bChkTTSuse", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bChkTTSuse", "True", App.Path & INIFILE
    Profile.bChkTTSuse = True
Else
    Profile.bChkTTSuse = retVal
End If

'34
'bTTSuse1 (기본값 true임에 유의)
Ret = GetPrivateProfileString(gUserid, "bChkTTSuse1", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bChkTTSuse1", "True", App.Path & INIFILE
    Profile.bChkTTSuse1 = True
Else
    Profile.bChkTTSuse1 = retVal
End If

'35
'bTTSuse2 (기본값 true임에 유의)
Ret = GetPrivateProfileString(gUserid, "bChkTTSuse2", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bChkTTSuse2", "True", App.Path & INIFILE
    Profile.bChkTTSuse2 = True
Else
    Profile.bChkTTSuse2 = retVal
End If

'36
'bTTSuse3 (기본값 true임에 유의)
Ret = GetPrivateProfileString(gUserid, "bChkTTSuse3", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bChkTTSuse3", "True", App.Path & INIFILE
    Profile.bChkTTSuse3 = True
Else
    Profile.bChkTTSuse3 = retVal
End If

'37
'bTTSuse4 (기본값 false임에 유의)
Ret = GetPrivateProfileString(gUserid, "bChkTTSuse4", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "bChkTTSuse4", "False", App.Path & INIFILE
    Profile.bChkTTSuse4 = False
Else
    Profile.bChkTTSuse4 = retVal
End If


'User24
Ret = GetPrivateProfileString(gUserid, "PenColor", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString gUserid, "PenColor", CStr(rgb(0, 255, 0)), App.Path & INIFILE
    Profile.PenColor = rgb(0, 255, 0)
Else
    Profile.PenColor = CLng(retVal)
End If

'++++++++++++++++++++


'clsProfile 로 이동해서 속성 추가 (클래스 편집기 이용)
'이후 clfFTP.saveProfile 부분기능 정의
End If '120 라인 위의 if와 쌍

'Main23
' Main에 설정하였음
Ret = GetPrivateProfileString("MAIN", "LastUser", "", RetStr, 255, App.Path & INIFILE)
retVal = LeftH(RetStr, Ret)

If retVal = "" Then
    WritePrivateProfileString "MAIN", "LastUser", "", App.Path & INIFILE
    Profile.LastUser = ""
Else
    Profile.LastUser = retVal
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
Dim Value As String
Dim Length As Long
Dim value_type As Long

    ' Open the key.
    If RegOpenKeyEx(selected_hkey, key, 0&, _
        KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS _
    Then
        GetKeyValue = "Error opening key."
        Exit Function
    End If

    ' Get the subkey's value.
    Length = 256
    Value = Space$(Length)
    If RegQueryValueEx(hKey, subkey, _
        0&, value_type, ByVal Value, Length) _
            <> ERROR_SUCCESS _
    Then
        GetKeyValue = "Error getting subkey value."
    Else
        ' Remove the trailing null character.
        GetKeyValue = Left$(Value, Length - 1)
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
   Dim Pos As Long
   Dim success As Long

   Dim success1 As Long
   
   Dim saveTime  As Double
    
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
    'ping the ip passing the address, text
    'to send, and the ECHO structure.
    
    saveTime = CDbl(Now) * 100000#
    
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
   Dim Pos As Long
   Dim success As Long

   Dim success1 As Long
   
   Dim saveTime  As Double
    
    testData = "12345678901234567890"
    If SocketsInitialize() Then
    
    'ping the ip passing the address, text
    'to send, and the ECHO structure.
    
    saveTime = CDbl(Now) * 100000#
    
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
        Dim i%, j%
        j% = 0
        For i% = 1 To Len(vValue)
            If Asc(Mid(vValue, i%, 1)) < 0 Then j% = j% + 1
        Next i%
        If j% > 0 Then vValue = Left(vValue, Len(vValue) - 1)
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



Private Sub Class_Terminate()
Set Profile = Nothing
End Sub

Private Sub Profile_saveProfile()
    Call saveProfile
End Sub

Private Sub saveProfile()
    
'==============================================================================
'       환경변수 Set
'==============================================================================
Dim retVal As String
Dim retVal1 As String
Dim retVal2 As String

'Set Profile = New clsProfile

'1
    WritePrivateProfileString gUserid, "fontsize", CStr(Profile.FontSize), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "bSetTimeOutStudy", CStr(Profile.bSetTimeOutStudy), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "TimeOutSecStudy", CStr(Profile.TimeOutSecStudy), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "TimeOutSec", CStr(Profile.TimeOutSec), App.Path & INIFILE
    WritePrivateProfileString gUserid, "bSetTimeOut", CStr(Profile.bSetTimeOut), App.Path & INIFILE
    WritePrivateProfileString gUserid, "bStreateOutCntCheck", CStr(Profile.bStreateOutCntCheck), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "bAlarm1", CStr(Profile.bAlarm1), App.Path & INIFILE
    WritePrivateProfileString gUserid, "bAlarm2", CStr(Profile.bAlarm2), App.Path & INIFILE
    WritePrivateProfileString gUserid, "bAlarm5", CStr(Profile.bAlarm5), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "CntOfAlarm1", CStr(Profile.CntOfAlarm1), App.Path & INIFILE
    WritePrivateProfileString gUserid, "CntOfAlarm2", CStr(Profile.CntOfAlarm2), App.Path & INIFILE
    WritePrivateProfileString gUserid, "CntOfAlarm5", CStr(Profile.CntOfAlarm5), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "fontBold", CStr(Profile.FontBold), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "DelayTime1", CStr(Profile.DelayTime1), App.Path & INIFILE
    WritePrivateProfileString gUserid, "DelayTime2", CStr(Profile.DelayTime2), App.Path & INIFILE
    
    WritePrivateProfileString gUserid, "nansu", CStr(Profile.nansu), App.Path & INIFILE
     
    WritePrivateProfileString gUserid, "setPoP1", CStr(Profile.setPoP1), App.Path & INIFILE
    WritePrivateProfileString gUserid, "setPoP2", CStr(Profile.setPoP2), App.Path & INIFILE
    WritePrivateProfileString gUserid, "compactDB", CStr(Profile.compactDB), App.Path & INIFILE
 
    WritePrivateProfileString gUserid, "ShowWhiteSpot", CStr(Profile.ShowWhiteSpot), App.Path & INIFILE
    WritePrivateProfileString gUserid, "checkWhenResize", CStr(Profile.checkWhenResize), App.Path & INIFILE
    WritePrivateProfileString gUserid, "setPoP3", CStr(Profile.setPoP3), App.Path & INIFILE
    
   '24
    WritePrivateProfileString gUserid, "FontName", CStr(Profile.FontName), App.Path & INIFILE
    
   '23
    WritePrivateProfileString "MAIN", "LastUser", CStr(Profile.LastUser), App.Path & INIFILE
    
    '25
    WritePrivateProfileString gUserid, "endSettingLogin", CStr(Profile.endSettingLogin), App.Path & INIFILE
    
    '26
    WritePrivateProfileString gUserid, "endSettingLoginDefaultMin", CStr(Profile.endSettingLoginDefaultMin), App.Path & INIFILE
    
    '27
    WritePrivateProfileString gUserid, "endSettingQuiz", CStr(Profile.endSettingQuiz), App.Path & INIFILE
    
    '28
    WritePrivateProfileString gUserid, "endSettingQuizDefaultMin", CStr(Profile.endSettingQuizDefaultMin), App.Path & INIFILE
    
    '29
    WritePrivateProfileString gUserid, "endSettingSetMin", CStr(Profile.endSettingSetMin), App.Path & INIFILE
    
    '30
    WritePrivateProfileString gUserid, "endSettingQuestion", CStr(Profile.endSettingQuestion), App.Path & INIFILE
    
    '31
    WritePrivateProfileString gUserid, "endSettingTimer", CStr(Profile.endSettingTimer), App.Path & INIFILE
    
'    '32
'    WritePrivateProfileString gUserid, "TTSbeta", CStr(Profile.bTTSbeta), App.Path & INIFILE
    
    '33
    WritePrivateProfileString gUserid, "bChkTTSuse", CStr(Profile.bChkTTSuse), App.Path & INIFILE
    
    '34
    WritePrivateProfileString gUserid, "bChkTTSuse1", CStr(Profile.bChkTTSuse1), App.Path & INIFILE
    '35
    WritePrivateProfileString gUserid, "bChkTTSuse2", CStr(Profile.bChkTTSuse2), App.Path & INIFILE
    '36
    WritePrivateProfileString gUserid, "bChkTTSuse3", CStr(Profile.bChkTTSuse3), App.Path & INIFILE
    '37
    WritePrivateProfileString gUserid, "bChkTTSuse4", CStr(Profile.bChkTTSuse4), App.Path & INIFILE
    
    'WritePrivateProfileString gUserid, "PenColor", Clng(Profile.PenColor    ), App.Path & INIFILE
    
End Sub


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
          errorOut err.LastDllError, "FtpOpenFile:PutFile"
          FTPUploadFile = False
          Exit Function
      End If
      
      Open sLocal For Binary Access Read As #1
      Size = LOF(1)
      For lBlock = 1 To Size \ BUFFERSIZE
          Get #1, , Data
          If (InternetWriteFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
              FTPUploadFile = False
              errorOut err.LastDllError, "InternetWriteFile"
              Exit Function
          End If
          DoEvents
          Sum = Sum + BUFFERSIZE
          RaiseEvent FileTransferProgress(Sum, Size)
      Next lBlock
      
      Get #1, , Data
      If (InternetWriteFile(hFile, Data(0), Size Mod BUFFERSIZE, Written) = 0) Then
          FTPUploadFile = False
          errorOut err.LastDllError, "InternetWriteFile2"
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

Private Sub errorOut(ByVal dwError As Long, ByRef szFunc As String)
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


Public Function FTPDownloadFile(hConnection As Long, sLocal As String, ByVal sRemote As String) As Boolean
    Dim Data(BUFFERSIZE - 1) As Byte ' array of 100 elements 0 to 99
    Dim Written As Long
    Dim Size As Long
    Dim Sum As Long
    Dim lBlock As Long
    Dim hFile As Long
    
    FTPDownloadFile = False
            
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
              errorOut err.LastDllError, "FtpOpenFile:GetFile"
              Exit Function
          End If
          
          Open sLocal For Binary Access Write As #1
          DoEvents
          Seek #1, 1
          DoEvents
          Sum = 1
          DoEvents
          For lBlock = 1 To Size \ BUFFERSIZE
              If (InternetReadFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
                  errorOut err.LastDllError, "InternetReadFile"
                  Close #1
                  Exit Function
              End If
              Put #1, , Data
              DoEvents
              Sum = Sum + BUFFERSIZE
              DoEvents
              RaiseEvent FileTransferProgress(Sum, Size)
              DoEvents
              DoEvents
          Next lBlock
          
          DoEvents
          ReDim Data2((Size Mod BUFFERSIZE) - 1) As Byte
          If (InternetReadFile(hFile, Data2(0), Size Mod BUFFERSIZE, Written) = 0) Then
              errorOut err.LastDllError, "InternetReadFile2"
              Close #1
              Exit Function
          End If
                      
          Put #1, , Data2
          
          Sum = Sum + (Size Mod BUFFERSIZE)
          DoEvents
          Size = Sum
          DoEvents
          RaiseEvent FileTransferProgress(Sum, Size)
                        DoEvents
              DoEvents
          Close #1
          
          InternetCloseHandle (hFile)
          DoEvents
          FTPDownloadFile = True
          DoEvents
      End If
   End If
End Function

Public Function GetProfiles() As String()
    Dim arr() As String
    
    Dim i As Long
    
    Dim strIni As String
    Dim strName$
    strIni = App.Path & INIFILE
    
    Dim FileHandle As Long
    Dim TextLine$
    
    FileHandle = FreeFile()
    Open strIni For Input As #FileHandle

    Do While Not EOF(FileHandle)        ' Loop until end of file
    Line Input #FileHandle, TextLine$  ' Read line into variable
         
        strName$ = Trim(TextLine$)
        If Left(strName$, 1) = "[" Then
        
            If Right(strName$, 1) = "]" Then
                i = i + 1
                ReDim Preserve arr(i)
                arr(i - 1) = Mid(strName$, 2, Len(strName$) - 2)
            End If
        End If
    Loop

    Close #FileHandle
    
    GetProfiles = arr
    
End Function

Public Sub deleteProfile(ByVal strUName As String)
    Call WritePrivateProfileSection(strUName, "", App.Path & INIFILE)
    Call WritePrivateProfileString(strUName, 0&, 0&, App.Path & INIFILE)

End Sub
