Attribute VB_Name = "Module1"
'==============================================================================
' 유의: mysql> select * from tp01 where pocketnm not in (SELECT a.subj from ts01 a inner join ts02 b on a.subj=b.subj) and cond like 'subj=%' and cond not like '%select%';  //==> 이렇게 쿼리를 수행하면 결과가 나온다면
' ts01과 잘못된 tp01 의 잘못된 관계의 DB가 있다는 것이다.
' 그리고 ts01 의 subj는 tp01의 pocketnm이다. 그리고 내부적으로 저장된 subj는 tp01의 cond에 해당된다.
'
'select * from ts01 where subj like '영속담%';
'+---------+--------------+
'| subj    | subjnm       |
'+---------+--------------+
'| 영속담1 | 영속담(중학) |
'+---------+--------------+
'1 row in set (0.00 sec)
'
'select * from tp01 where cond like '%영속담1%';
'+------+----------+---------------------+------+------+
'| seq  | pocketnm | cond                | xm   | chkm |
'+------+----------+---------------------+------+------+
'|   43 | 영어속담 | subj='영속담1'|tq02 |    0 |    0 |
'+------+----------+---------------------+------+------+
'1 row in set (0.00 sec)
'
'여기서 tp01의 pocketnm을 [영어속담]에서 영속담으로 바꿔야 한다. 아니면 [영속담(중학)] 으로 바꿔야 한다.
'update tp01 set pocketnm='영속담(중학)' where seq=43;
'
'select * from tp01 where pocketnm not in (SELECT a.subj from ts01 a
'
'
'select concat('update tp01 set pocketnm=''',subjnm,''' where seq=',seq,';') k from (select bbbb.subjnm,aaaa.seq from (select seq,substring(cond,aa+1,bb-aa-2)
' strip_subj from (select instr(cond,'''') aa,instr(cond,'|') bb,cond,seq,pocketnm from tp01
' where pocketnm not in (SELECT a.subj from ts01 a inner join ts02 b on a.subj=b.subj)
' and cond like 'subj=%' and cond not like '%select%') aaa ) aaaa inner join ts01 bbbb on (aaaa.strip_subj=bbbb.subj)) aaaaa
'
' 위에서 만들어진 쿼리문을 실행한다.
'-------------------------------------
' Fn_ <=== 여기서쿼리 실행됨(shift+f2) 키를 누르면 정의로 감.
' 추가할 기능
' windows7 에서는 ms agent를 별도로 설치해야 한다.
' 휴대전화를 통한 인증 프로세스 및 일주일 마다 인증번호 변경하여 인증번호 확인(유도) 단 미접속시
'  매일 접속한 사용자나 인증번호를 어렵게 바꾼경우는 5자리 이상의 알파벳(은) 전화하지 않는다.
'  휴대기기가 변경된 경우 그러한 사실을 이전 번호로도 쏴준다. 이는 무단 변조되지 않았음을 알려주는 것이다.
'  개인정보 변경에서 그러한 기능을 설정한다.
'  이웃추가에서 나를 이웃으로 한 사람 알기
'  이웃설정에서 내가 이웃으로 설정할 수 있도록 하며 메시지도 보낼 수 있다.
'  가입시 이메일주소를 다시한번 확인한다.
' 배경색, 글자색 설정 가능 하게
' 오류발생 2003서버에서
' 341 런타임 오류발생 컨트롤 배열 인덱스가 잘못되었습니다.==>
  '회원정보데이터가 잘못 생성된 경우 시험보기를 시작하면 그런 현상이 발생된다.
  'DB를 export import 하면 오류 생기고 폴더체로 풀어야 한다.
  
' pic1 등을 다시 붙여넣기 해도 마찬가지고 다시만들어도 마찬가지
' option base 0 을 하면 어떨까?
' 2003에서 컴파일한것은 2003에서 돌아가고 그렇지 않은것은 안돌아가고 한다.
' 각각의 머신에서 컴파일하는 수밖에 없다. 배타적으로는 실행파일이 작동되지 않는다.

' 0. 순위(전체 사용자)
' 0.1 인터넷 서비스
' 0.2 암호등록
' 0.3 연결계정의 암호화 후 저장(암호화)
'
' 1. 힌트입력기능(관리자인경우)
' 2. 부동산 관련 문제에 비트맵추가
' 3. 멀티미디어 입력에 따른 문제 추가..(문제네....)
' 4. 실적 새로운 문제를 푼것과 복습문제 풀이수가 많게 나오는 문제
' 5. global application association
' 6. disign and color skin
' 7. 달력에서 전체일정의 중복라인이 안생긴는경우(명암차이를 두어야 할것같음.) hsld함수에서...
' 8. 힌트의 권한 설정(공개및 비공개)
' 9. 알고있는 지식수 및 성적조회(등수)
' 10. 퀴즈수정및 승인처리기능
' 11. 나의 퀴즈
' 12. ocp 퀴즈등록
' 13. 과목추가
' 14. 유료컨텐츠로의 전환 기능
' 15. 비트맵을 ftp로 다운로드하기 (내장 res 이용하지 않음)
' 16. 사용자 정의 사전 link
' 17. 통계기능 추가
' 18. Integer 자료를 썼을경우 숫자가 커짐으로 인한 오류발생 예측해야함 고쳐야함
'mysql> select a.* from tu02 a,vq01 b where a.userid='영어' and a.subj=b.subj and
' a.seq=b.seq and b.quiz='irritate' and a.reserve_ymd!='99999999';
'+----------+------+--------+------+------+------+------------+-------------+----
'-----+
'| subj     | seq  | userid | o    | x    | chk  | update_ymd | reserve_ymd | gan
'gyek |
'+----------+------+--------+------+------+------+------------+-------------+----
'-----+
'| 영단어1  |  698 | 영어   |    0 |    0 |    0 | 20040614   | 99999999    |
'   0 |
'| 중3단어  |  662 | 영어   |    0 |    0 |    0 | 20040723   | 99999999    |
'   0 |
'| 토익Voca |  232 | 영어   |    1 |    3 |    1 | 20050407   | 20050407    |
'   0 |
'+----------+------+--------+------+------+------+------------+-------------+----
'-----+
'3 rows in set (0.12 sec)
'
'18. 학습계획시 실현가능성 있게 (작심3일형으로 해서 복습하는 스케쥴을 고려하여 일자 계획)
'    하는 방법 추가
'
'19. 시험계획시 문제수 추리는 과정 필요
' 예를 들어 2000문항이라고 하면 20개만 하겠다고 설정
'
'20. 종료시간 설정 기능(iqup한자 기능처럼)
'
'21. 닌텐도 기능추가(더하기,빼기 기능 과 집중력 증가를 위한 워밍업기능)
'
'22. 선생님용 어플리케이션은 별도의 어플리케이션으로 시리얼을 가짐
'기능 추가(선생님은 해당 아이디의 회원을 지명하여 자신의
'제자로 지명 가능하며 선생님이 산 과목을 아이들에게 지명할 수 있다.
'즉 선생님은 과목롤과 학생수 라이센스를 가지는데 과목의 설정 가능한 년도를
'1년이면 1년 2년이면 2년 그렇게 설정 할 수 있다.
'만약 1년 동안 과목을 임대하면 제자들의 최대 기간도 그 과목의 종료일로만 가능하다
'학생이 시험지를 만들어 놓은경우 어떻게 하지?... 명확히 답이 없네...
'
'==============================================================================
'개발환경 구축 관련
'------------------------------------------------------------------------------
'mysqldump -uroot -pnextedu --default-character-set=latin1 pocket_znlwm_000001>pocket_znlwm_000001.sql
'(백업)C:\Program Files\MySQL\MySQL Server 5.1\bin>mysqldump -uroot -pnextedu --default-character-set=latin1 --databases pocket_znlwm_000001>pocket.sql
'(복원)C:\Program Files\MySQL\MySQL Server 5.1\bin>mysql -uroot -pnextedu mysql
'Welcome to the MySQL monitor.  Commands end with ; or \g.
'Your MySQL connection id is 1
'Server version: 5.1.22-rc-community MySQL Community Server (GPL)
'
'Type 'help;' or '\h' for help. Type '\c' to clear the buffer.
'
'mysql> create database pocket_znlwm_000001;
'Query OK, 1 row affected (0.03 sec)
'
'mysql> exit
'Bye
'
'C:\Program Files\MySQL\MySQL Server 5.1\bin>mysql -uroot -pnextedu pocket_znlwm_000001 <pocket.sql
'
'C:\Program Files\MySQL\MySQL Server 5.1\bin>mysql -uroot -pnextedu mysql
'Welcome to the MySQL monitor.  Commands end with ; or \g.
'Your MySQL connection id is 3
'Server version: 5.1.22-rc-community MySQL Community Server (GPL)
'
'Type 'help;' or '\h' for help. Type '\c' to clear the buffer.
'
'mysql> create user changeit;
'Query OK, 0 rows affected (0.00 sec)
'
'mysql> grant all privileges on pocket_znlwm_000001.* to changeit@"%" identified By '' with grant option;
'Query OK, 0 rows affected (0.00 sec)
'
'프로젝트 속성>만들기>명령줄인수> /noupdate --host 127.0.0.1 -uchangeit -P3306 pocket_znlwm_000001 라고 넣는다.(db명은 적용 안된상태)
'
'==========================================================
'최근 수행한 쿼리 (쓸데 없이 들어간 줄바꿈 문자 및 공백 문자 제거)
'update vq01 set quiz=replace(replace(replace(replace(quiz,concat(char(13),char(10)),' '),'    ',' '),'   ',' '),'  ',' ')
',a=replace(replace(replace(replace(a,concat(char(13),char(10)),' '),'    ',' '),'   ',' '),'  ',' ')
',b=replace(replace(replace(replace(b,concat(char(13),char(10)),' '),'    ',' '),'   ',' '),'  ',' ')
',c=replace(replace(replace(replace(c,concat(char(13),char(10)),' '),'    ',' '),'   ',' '),'  ',' ')
',d=replace(replace(replace(replace(d,concat(char(13),char(10)),' '),'    ',' '),'   ',' '),'  ',' ')
'  where subj = '토익990' and instr(quiz,concat(char(13),char(10))) and instr(quiz,'(A)')>0 ;
'
'update vq01 set quiz=replace(replace(replace(replace(quiz,concat(char(13),char(10)),' '),'    ',' '),'   ',' '),'  ',' ')
'where subj = '토익990' and instr(quiz,concat(char(13),char(10))) and instr(quiz,'__')>0 ;
'
'===========================================================
'사용자 삭제를 위한 쿼리
'delete from tg01 where userid='';
'delete from tp02 where userid='';
'delete from tp03 where userid='';
'delete from ts02 where userid='';
'delete from tt01 where userid='';
'delete from tt02 where userid='';
'delete from tt03 where userid='';
'delete from tu01 where userid='';
'delete from tu02 where userid='';/*big size */
'delete from tu03 where userid='';
'update th01 set  userid= concat("#",userid) where userid='';

'mysql> delete from tg01 where userid='새테스트';
'Query OK, 3 rows affected (0.05 sec)
'
'mysql> delete from tp02 where userid='새테스트';
'Query OK, 6 rows affected (0.00 sec)
'
'mysql> delete from tp03 where userid='새테스트';
'Query OK, 162 rows affected (0.06 sec)
'
'mysql> delete from ts02 where userid='새테스트';
'Query OK, 31 rows affected (0.03 sec)
'
'mysql> delete from tt01 where userid='새테스트';
'Query OK, 1 row affected (0.00 sec)
'
'mysql> delete from tt02 where userid='새테스트';
'Query OK, 36 rows affected (0.00 sec)
'
'mysql> delete from tt03 where userid='새테스트';
'Query OK, 36 rows affected (0.00 sec)
'
'mysql> delete from tu01 where userid='새테스트';
'Query OK, 1 row affected (0.00 sec)
'
'mysql> delete from tu02 where userid='새테스트';
'Query OK, 81 rows affected (0.01 sec)
'
'mysql> delete from tu03 where userid='새테스트';
'Query OK, 6 rows affected (0.00 sec)
'
'mysql> update th01 set  userid= concat("#",userid) where userid='새테스트';
'Query OK, 0 rows affected (0.01 sec)
'Rows matched: 0  Changed: 0  Warnings: 0
'================================================================
'과목추가 쿼리(과목구입시)
'insert into ts02 values('구약:23' ,'박규선','20070701','21000701');
'구약:23 은 ts01 의 subj 코드이다.
'혹은 tp01 의 pocketname 일수도 있는데
'tp01의 cond의 subj는 2가지 의미가 있는데 select 가 없는 경우는 ts01의 subj가 cond.subj이다.
'그리고 like가 있는 경우의 cond.subj는 suject이름혹은 설명이다.(긴이름)
'확실하지는 않다. 명확한 분석이 더 필요하다. 해놓고도 해놓은 사람이 모르는 경우이다.
'
'================================================================
Option Explicit
Option Base 0

Global Const BUILD_NUMBER = "2016-01-08" '####

'Global Const TTS_GLOBAL_URL = "http://www.oddcast.com/demos/tts/tts_example.php?clients"
'Global Const TTS_GLOBAL_URL = "https://www.vocalware.com/index/demo"
'Global Const TTS_GLOBAL_URL = "http://speech.diotek.com/m41.php#contents"
Global Const TTS_GLOBAL_URL = "http://speech.diotek.com/m41.php#edit"
Global Const TTS_GLOBAL_URL2 = "https://translate.google.com/translate_tts?ie=UTF-8&q=$$$&tl=MMM&client=tw-ob"

Global TreeViewCache As Collection '트리뷰의 성능을 개선하기 위한 용도로 사용

'pttsnet_r.php

Private ZF As New Cls_GetFileType

'****mp3 play******
'http://www.vbforfree.com/easy-way-to-play-mp3-wav-wma-files-vb/
'CommandString = "open """ & FileName & """ type mpegvideo alias " & FileName
'retVal = mciSendString(CommandString, vbNullString, 0, 0)
'Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
'(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal _
'uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

'********registry**********
'http://www.windowsdevcenter.com/pub/a/windows/2004/06/15/VB_Registry_Keys.html?page=2

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Const REG_OPTION_BACKUP_RESTORE = 4 ' open for backup or restore
Const REG_OPTION_VOLATILE = 1 ' Key is not preserved when system is rebooted
Const REG_OPTION_NON_VOLATILE = 0 ' Key is preserved when system is rebooted
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
        Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
        Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL

        Public Const STANDARD_RIGHTS_ALL = &H1F0000

        Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
Const REG_SZ = 1

Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1

Public Const SYNCHRONIZE = &H100000

Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE _
                        Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
                        And (Not SYNCHRONIZE))
Public Const KEY_SET_VALUE = &H2
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE _
                         Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE _
                              Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY _
                              Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY _
                              Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
                              
Public Const ERROR_SUCCESS = 0&
'Among the constants representing error codes are the following:
Public Const ERROR_FILE_NOT_FOUND = 2&         ' Registry path does not exist
Public Const ERROR_ACCESS_DENIED = 5&          ' Requested permissions not available
Public Const ERROR_INVALID_HANDLE = 6&         ' Invalid handle or top-level key
Public Const ERROR_BAD_NETPATH = 53            ' Network path not found
Public Const ERROR_INVALID_PARAMETER = 87      ' Bad parameter to a Win32 API function
Public Const ERROR_CALL_NOT_IMPLEMENTED = 120& ' Function valid only in WinNT/2000?XP
Public Const ERROR_INSUFFICIENT_BUFFER = 122   ' Buffer too small to hold data
Public Const ERROR_BAD_PATHNAME = 161          ' Registry path does not exist
Public Const ERROR_NO_MORE_ITEMS = 259&        ' Invalid enumerated value
Public Const ERROR_BADDB = 1009                ' Corrupted registry
Public Const ERROR_BADKEY = 1010               ' Invalid registry key
Public Const ERROR_CANTOPEN = 1011&            ' Cannot open registry key
Public Const ERROR_CANTREAD = 1012&            ' Cannot read from registry key
Public Const ERROR_CANTWRITE = 1013&           ' Cannot write to registry key
Public Const ERROR_REGISTRY_RECOVERED = 1014&  ' Recovery of part of registry successful
Public Const ERROR_REGISTRY_CORRUPT = 1015&    ' Corrupted registry
Public Const ERROR_REGISTRY_IO_FAILED = 1016&  ' Input/output operation failed
Public Const ERROR_NOT_REGISTRY_FILE = 1017&   ' Input file not in registry file format
Public Const ERROR_KEY_DELETED = 1018&         ' Key already deleted
Public Const ERROR_KEY_HAS_CHILDREN = 1020&    ' Key has subkeys & cannot be deleted


'Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
'       (ByVal hKey As Long, _                ' Handle of already open key
'       ByVal lpSubKey As String, _           ' Path from hkey to key to open/create
'       ByVal Reserved As Long, _             ' Reserved, must be 0
'       ByVal lpClass As String, _            ' Reserved, must be a null string
'       ByVal dwOptions As Long, _            ' Type of key, or backup/restore
'       ByVal samDesired As Long, _           ' SAM constant(s)
'       lpSecurityAttributes As SECURITY_ATTRIBUTES, _
'       phkResult As Long, _                  ' Handle of opened/created key
'       lpdwDisposition As Long) As Long      '
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
       (ByVal hKey As Long, _
       ByVal lpSubKey As String, _
       ByVal Reserved As Long, _
       ByVal lpClass As String, _
       ByVal dwOptions As Long, _
       ByVal samDesired As Long, _
       lpSecurityAttributes As SECURITY_ATTRIBUTES, _
       phkResult As Long, _
       lpdwDisposition As Long) As Long
       
       Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
 "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
 As String, ByVal Reserved As Long, ByVal dwType As Long, _
 lpData As Any, ByVal cbData As Long) As Long

Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Global TmrAfterTTS_exit As Boolean
'**********************

Global refCquiz As Object
Global gbZipStop As Boolean
Global cmd As String '입력 파라메타
Global gMutex As Boolean  '뮤텍스
Global gQuizOnResize As Boolean '리사이즈를 하는경우에는 시간 흐름이 다시 처음부터 되지 않게하기 위해서이다.
Global gMainOnResize As Boolean '퀴즈리사이즈의경우 visible설정에 의해 호출되기도하나 이경우 resize는 실제 UI에의한 resize이다.
Global gPgBarSaveValue As Single '위와 같이 사용된다.
Global gTimer As Double 'Last Queuried Time[Sec]
Global gChangUser As Boolean
Global gStrIP As String
Global gStrPort As String
Global gArrUsers() As String
Global vDataGlobal As String ' 복습문제를 풀게 되는 경우 셋팅
Global vDataGlobal2 As String '틀린문제를 다시 푸는 경우

Global Const PI = 3.1415926536
Global Const ALLOW_AFFECT_DATE As Long = 30 '복습예약일이 오늘로 부터 미래의 한달30일 이내인 경우에는 유효학습으로 본다(예정일검토용) 양수여야한다.
Global Const ALLOW_AFFECT_DATE30 As Long = -90 '석달이내 최근에 학습한 것인지 체크한다.(update일자검토용)음수여야한다.
'****************************************************************
' 코딩으로 ODBC설정
'****************************************************************
'Constant Declaration

Private Const ODBC_ADD_DSN = 1 ' Add data source
Private Const ODBC_CONFIG_DSN = 2 ' Configure (edit) data source
Private Const ODBC_REMOVE_DSN = 3 ' Remove data source
Private Const vbAPINull As Long = 0& ' NULL Pointer

'Function Declare
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
(ByVal hwndParent As Long, ByVal fRequest As Long, _
ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

'****************************************************************


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
'Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Public Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'  대화 모드에 대한 비트 필드
Global Const IME_CMODE_ALPHANUMERIC = &H0
Global Const IME_CMODE_NATIVE = &H1
Global Const IME_CMODE_CHINESE = IME_CMODE_NATIVE
Global Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Global Const IME_CMODE_JAPANESE = IME_CMODE_NATIVE
Global Const IME_CMODE_KATAKANA = &H2                   '  only effect under IME_CMODE_NATIVE
Global Const IME_CMODE_LANGUAGE = &H3
Global Const IME_CMODE_FULLSHAPE = &H8
Global Const IME_CMODE_ROMAN = &H10
Global Const IME_CMODE_CHARCODE = &H20
Global Const IME_CMODE_HANJACONVERT = &H40
Global Const IME_CMODE_SOFTKBD = &H80
Global Const IME_CMODE_NOCONVERSION = &H100
Global Const IME_CMODE_EUDC = &H200
Global Const IME_CMODE_SYMBOL = &H400
Global Const IME_CMODE_FIXED = &H800


Global Const IME_SMODE_NONE = &H0
Global Const IME_SMODE_PLAURALCLAUSE = &H1
Global Const IME_SMODE_SINGLECONVERT = &H2
Global Const IME_SMODE_AUTOMATIC = &H4
Global Const IME_SMODE_PHRASEPREDICT = &H8
Global Const IME_SMODE_CONVERSATION = &H10

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public status As Panel

Global Const LISTVIEW_MODE0 = "큰 아이콘"
Global Const LISTVIEW_MODE1 = "작은 아이콘"
Global Const LISTVIEW_MODE2 = "목록 보기"
Global Const LISTVIEW_MODE3 = "자세히 보기"

Public gScreenHourGLS As Boolean

  
Public Type HSL
  hue As Long '1-240
  Saturation As Long '0-100
  Luminance As Long '0-100
End Type

Public STRCON   As String
Public Const GPWD2 = "d=zk"

Global gUserid As String
Global gSessionId As String

Global gLogonCnt As Long '나의 방문횟수(로그온 횟수 아님,하루한번 기록갱신됨)
Global gPassWord As String
Global gbIsSuperAdmin As Boolean

Global gPocket As String
Global gOrder As Long
Global gLastNew As Long
Global gHangSu As Long
Global gIsNew As Boolean
Global gNowNum As Long
Global gbReadTu03 As Boolean
Global gChasu As Long
Global gbMnuExamClick As Boolean '시험보기메뉴를클릭하고있는동안에 quiz_disp 가 호출되면 skip설정
Global gnBangmoonConstraint As Long '새로운 회원초대를 위한 방문제한 횟수

Public fMainForm As frmMain
Public Con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sSql As String

Public SELFLAG As Boolean

Global cFTP As New clsFTP
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

'//결과값을 여러가지 형태로 보여주는 result
Public Type ut_bRecordSet
   result As Boolean       ' Query의 결과 성공( True) /  실패 ( False)
   Error As Boolean        ' Error 유무   Error발생 (TRUE) / Error미 발생 (FALSE)
   nrow As Long
   rs As ADODB.Recordset   ' Query의 결과에 의한 RecordSet
End Type

'Dim conHandle As ADODB.Connection

Public SQLCA As ut_bRecordSet
'======================================================================
'폰트 리스트 얻기 API
'======================================================================
Public Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" _
     (ByVal hdc As Long, ByVal lpszFamily As String, _
     ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
     ByVal hdc As Long) As Long
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4
Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

'html 템플릿 문자열변수는 처음에 한번 셋팅된 값이 유지되는 전역변수이므로유의할것
Private dic_kr_str As String
Private dic_modum_str As String
Private dic_str As String
Private dic_ee_str As String
Private dic_hanja_str As String '처음에 한번 셋팅된 값이 유지되는 전역변수이므로유의할것
Private dic_jp_str As String
Private dic_cn_str As String

Private tts_ko_str As String
Private tts_ko_str2 As String
Private tts_ko_str6 As String
Private tts_eng_str As String

'------------
'UTF8관련
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Byte, Source As Byte, ByVal Length As Long)

'UTF8을 ANSI ASCII 코드로 변환
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Declare Function GetACP Lib "kernel32" () As Long

Private Declare Function WideCharToMultiByteArray Lib "kernel32" Alias "WideCharToMultiByte" _
(ByVal CodePage As Long, _
ByVal dwFlags As Long, _
ByRef lpWideCharStr As Byte, _
ByVal cchWideChar As Long, _
ByRef lpMultiByteStr As Byte, _
ByVal cchMultiByte As Long, _
ByVal lpDefaultChar As Long, _
ByVal lpUsedDefaultChar As Long) As Long

'Private Declare Function MultiByteToWideChar Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long


Private Type abyteBOM 'Byte Order Mark
    byte1 As Byte
    Byte2 As Byte
    byte3 As Byte
    byte4 As Byte
End Type

Private Const CP_UTF8 As Long = 65001

Private Const CP_ACP = 0 ' default to ANSI code page

'++++++++++++++++++++++++
Private mGetRndName As String


'이상UTF8
'------------


'http://www.google.com/search?tbm=vid&hl=ko&source=hp&biw=1302&bih=789&q=%ED%99%94&gbv=2&oq=%ED%99%94#hl=ko&gbv=2&tbm=vid&sclient=psy-ab&
'q=TEST&oq=TEST&gs_l=serp.3..0l4.23343.25841.1.26233.5.5.0.0.0.0.163.430.3j1.4.0...0.0...1c.1j4.oWNfDNzSKig&pbx=1&bav=on.2,or.r_gc.r_pw.r_qf.&fp=5ed5d632c557d8a0&bpcl=35466521&biw=1302&bih=761
'https://www.google.co.kr/#hl=ko&newwindow=1&output=search&sclient=psy-ab&q=화산&oq = 화산&gs_l=hp.3..0l4.2806.5386.0.6445.2.2.0.0.0.0.91.182.2.2.0...0.0...1c.1j4.qX1r9edXsKg&pbx=1&bav=on.2,or.r_gc.r_pw.r_qf.&fp=6ab7330b8f11c38d&bpcl=35466521&biw=1920&bih=955

Sub Main()
On Error Resume Next
'FileSystem.RmDir ("tmp")
FileSystem.MkDir ("tmp")
On Error GoTo 0

dic_modum_str = "<html><head> "
dic_modum_str = dic_modum_str + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'> "
dic_modum_str = dic_modum_str + "</head> "
dic_modum_str = dic_modum_str + "<body onload='show_dic()'> "
dic_modum_str = dic_modum_str + "<script> "
dic_modum_str = dic_modum_str + "function show_dic() "
dic_modum_str = dic_modum_str + "{ "
dic_modum_str = dic_modum_str + " document.form.query.value=encodeURIComponent(document.getElementById('div_xmp').innerHTML);"
dic_modum_str = dic_modum_str + " var obj_query = document.getElementById('query'); "
dic_modum_str = dic_modum_str + " if(obj_query.value==''){ "
dic_modum_str = dic_modum_str + "  alert('단어를 입력하세요.'); "
dic_modum_str = dic_modum_str + "  obj_query.focus(); "
dic_modum_str = dic_modum_str + "  return; "
dic_modum_str = dic_modum_str + " } "
'dic_modum_str = dic_modum_str + " hwin = window.open('http://dic.naver.com/search.nhn?target=dic&ie=utf8&query_utf=&isOnlyViewEE=&query='+obj_query.value+'','_pocket_quiz_pop_','left=0,top=0,width=1024,height=768,status=yes,scrollbars=yes,toolbar=yes,address=yes,resizable=yes'); "
dic_modum_str = dic_modum_str + " hwin = window.open('http://dic.naver.com/search.nhn?target=dic&ie=utf8&query_utf=&isOnlyViewEE=&query='+obj_query.value+'','_pocket_quiz_pop_','left='+(screen.availWidth-1024)+',top=0,width=1024,height='+screen.availHeight+',status=yes,scrollbars=yes,toolbar=yes,address=yes,resizable=yes'); "
dic_modum_str = dic_modum_str + " hwin.focus();"
dic_modum_str = dic_modum_str + "} "
dic_modum_str = dic_modum_str + "</script> "
dic_modum_str = dic_modum_str + "<form name='form'><xmp id='div_xmp'>MZM</xmp> "
dic_modum_str = dic_modum_str + "<input type='text' name='query' width='100px' value=''> "
dic_modum_str = dic_modum_str + "<input type='button' value='찾기' onclick='show_dic()'> "
dic_modum_str = dic_modum_str + "</form> "
dic_modum_str = dic_modum_str + "</body> "
dic_modum_str = dic_modum_str + "</html>"

dic_kr_str = "<html><head> "
dic_kr_str = dic_kr_str + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'> "
dic_kr_str = dic_kr_str + "</head> "
dic_kr_str = dic_kr_str + "<body onload='show_dic()'> "
dic_kr_str = dic_kr_str + "<script> "
dic_kr_str = dic_kr_str + "function show_dic() "
dic_kr_str = dic_kr_str + "{ "
dic_kr_str = dic_kr_str + " document.form.query.value=encodeURIComponent(document.getElementById('div_xmp').innerHTML);"
dic_kr_str = dic_kr_str + " var obj_query = document.getElementById('query'); "
dic_kr_str = dic_kr_str + " if(obj_query.value==''){ "
dic_kr_str = dic_kr_str + "  alert('단어를 입력하세요.'); "
dic_kr_str = dic_kr_str + "  obj_query.focus(); "
dic_kr_str = dic_kr_str + "  return; "
dic_kr_str = dic_kr_str + " } "
dic_kr_str = dic_kr_str + " window.open('http://krdic.naver.com/small_search.nhn?query='+obj_query.value+'','_pocket_quiz_pop_','width=400,height=500,status=yes,scrollbars=no,toolbar=no'); "
dic_kr_str = dic_kr_str + "} "
dic_kr_str = dic_kr_str + "</script> "
dic_kr_str = dic_kr_str + "<form name='form'><xmp id='div_xmp'>MZM</xmp> "
dic_kr_str = dic_kr_str + "<input type='text' name='query' width='100px' value=''> "
dic_kr_str = dic_kr_str + "<input type='button' value='찾기' onclick='show_dic()'> "
dic_kr_str = dic_kr_str + "</form> "
dic_kr_str = dic_kr_str + "</body> "
dic_kr_str = dic_kr_str + "</html>"

dic_str = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" + vbCrLf
dic_str = dic_str + "<html lang='ko'><head> "
dic_str = dic_str + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'> "
dic_str = dic_str + "<script type='text/javascript'> "
dic_str = dic_str + "function show_dic() "
dic_str = dic_str + "{ "
dic_str = dic_str + " document.form.query.value=encodeURIComponent(document.getElementById('div_xmp').innerHTML);"
dic_str = dic_str + "    var obj_query = document.getElementById('query'); "
dic_str = dic_str + " if(obj_query.value==''){ "
dic_str = dic_str + "  alert('단어를 입력하세요.'); "
dic_str = dic_str + "  obj_query.focus(); "
dic_str = dic_str + "  return; "
dic_str = dic_str + " } "
dic_str = dic_str + " window.open('http://endic.naver.com/popManager.nhn?m=search&searchOption=&query='+obj_query.value+'','pocket_quiz','width=300,height=400,status=yes,scrollbars=no,toolbar=no'); "
dic_str = dic_str + "} " 'http://endic.naver.com/search.nhn?query=
dic_str = dic_str + "</script> "
dic_str = dic_str + "</head> "
dic_str = dic_str + "<body onload='show_dic()'> "
dic_str = dic_str + "<form name='form'><xmp id='div_xmp'>MZM</xmp> "
dic_str = dic_str + "<input type='text' name='query' width='100px' value=''> "
dic_str = dic_str + "<input type='button' value='찾기' onclick='show_dic()'> "
dic_str = dic_str + "</form> "
dic_str = dic_str + "</body> "
dic_str = dic_str + "</html>"

dic_ee_str = "<html><head> "
dic_ee_str = dic_ee_str + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'> "
dic_ee_str = dic_ee_str + "</head> "
dic_ee_str = dic_ee_str + "<body onload='show_dic()'> "
dic_ee_str = dic_ee_str + "<script> "
dic_ee_str = dic_ee_str + "function show_dic() "
dic_ee_str = dic_ee_str + "{ "
dic_ee_str = dic_ee_str + " document.form.query.value=encodeURIComponent(document.getElementById('div_xmp').innerHTML);"
dic_ee_str = dic_ee_str + "    var obj_query = document.getElementById('query'); "
dic_ee_str = dic_ee_str + " if(obj_query.value==''){ "
dic_ee_str = dic_ee_str + "  alert('단어를 입력하세요.'); "
dic_ee_str = dic_ee_str + "  obj_query.focus(); "
dic_ee_str = dic_ee_str + "  return; "
dic_ee_str = dic_ee_str + " } "
dic_ee_str = dic_ee_str + " window.open('http://eedic.naver.com/small.naver?query='+obj_query.value+'','_pocket_quiz_pop_','width=400,height=500,status=yes,scrollbars=no,toolbar=no'); "
dic_ee_str = dic_ee_str + "} "
dic_ee_str = dic_ee_str + "</script> "
dic_ee_str = dic_ee_str + "<form name='form'><xmp id='div_xmp'>MZM</xmp> "
dic_ee_str = dic_ee_str + "<input type='text' name='query' width='100px' value=''> "
dic_ee_str = dic_ee_str + "<input type='button' value='찾기' onclick='show_dic()'> "
dic_ee_str = dic_ee_str + "</form> "
dic_ee_str = dic_ee_str + "</body> "
dic_ee_str = dic_ee_str + "</html>"

dic_hanja_str = "<html><head> "
dic_hanja_str = dic_hanja_str + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'> "
dic_hanja_str = dic_hanja_str + "</head> "
dic_hanja_str = dic_hanja_str + "<body onload='show_dic()'> "
dic_hanja_str = dic_hanja_str + "<script> "
dic_hanja_str = dic_hanja_str + "function show_dic() "
dic_hanja_str = dic_hanja_str + "{ "
dic_hanja_str = dic_hanja_str + " document.form.query.value=encodeURIComponent(document.getElementById('div_xmp').innerHTML);"
dic_hanja_str = dic_hanja_str + "    var obj_query = document.getElementById('query'); "
dic_hanja_str = dic_hanja_str + " if(obj_query.value==''){ "
dic_hanja_str = dic_hanja_str + "  alert('단어를 입력하세요.'); "
dic_hanja_str = dic_hanja_str + "  obj_query.focus(); "
dic_hanja_str = dic_hanja_str + "  return; "
dic_hanja_str = dic_hanja_str + " } "
dic_hanja_str = dic_hanja_str + " window.open('http://hanja.naver.com/small/search?query='+obj_query.value+'','_pocket_quiz_pop_','width=300,height=400,status=yes,scrollbars=no,toolbar=no'); "
dic_hanja_str = dic_hanja_str + "} "
dic_hanja_str = dic_hanja_str + "</script> "
dic_hanja_str = dic_hanja_str + "<form name='form'><xmp id='div_xmp'>MZM</xmp> "
dic_hanja_str = dic_hanja_str + "<input type='text' name='query' width='100px' value=''> "
dic_hanja_str = dic_hanja_str + "<input type='button' value='찾기' onclick='show_dic()'> "
dic_hanja_str = dic_hanja_str + "</form> "
dic_hanja_str = dic_hanja_str + "</body> "
dic_hanja_str = dic_hanja_str + "</html>"

dic_jp_str = "<html><head> "
dic_jp_str = dic_jp_str + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'> "
dic_jp_str = dic_jp_str + "</head> "
dic_jp_str = dic_jp_str + "<body onload='show_dic()'> "
dic_jp_str = dic_jp_str + "<script> "
dic_jp_str = dic_jp_str + "function show_dic() "
dic_jp_str = dic_jp_str + "{ "
dic_jp_str = dic_jp_str + " document.form.query.value=encodeURIComponent(document.getElementById('div_xmp').innerHTML);"
dic_jp_str = dic_jp_str + "    var obj_query = document.getElementById('query'); "
dic_jp_str = dic_jp_str + " if(obj_query.value==''){ "
dic_jp_str = dic_jp_str + "  alert('단어를 입력하세요.'); "
dic_jp_str = dic_jp_str + "  obj_query.focus(); "
dic_jp_str = dic_jp_str + "  return; "
dic_jp_str = dic_jp_str + " } "
dic_jp_str = dic_jp_str + " window.open('http://jpdic.naver.com/mini_search.nhn?query='+obj_query.value+'','_pocket_quiz_pop_','width=300,height=400,status=yes,scrollbars=no,toolbar=no'); "
dic_jp_str = dic_jp_str + "} "
dic_jp_str = dic_jp_str + "</script> "
dic_jp_str = dic_jp_str + "<form name='form'><xmp id='div_xmp'>MZM</xmp> "
dic_jp_str = dic_jp_str + "<input type='text' name='query' width='100px' value=''> "
dic_jp_str = dic_jp_str + "<input type='button' value='찾기' onclick='show_dic()'> "
dic_jp_str = dic_jp_str + "</form> "
dic_jp_str = dic_jp_str + "</body> "
dic_jp_str = dic_jp_str + "</html>"

dic_cn_str = "<html><head> "
dic_cn_str = dic_cn_str + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'> "
dic_cn_str = dic_cn_str + "</head> "
dic_cn_str = dic_cn_str + "<body onload='show_dic()'> "
dic_cn_str = dic_cn_str + "<script> "
dic_cn_str = dic_cn_str + "function show_dic() "
dic_cn_str = dic_cn_str + "{ "
dic_cn_str = dic_cn_str + " document.form.query.value=encodeURIComponent(document.getElementById('div_xmp').innerHTML);"
dic_cn_str = dic_cn_str + "    var obj_query = document.getElementById('query'); "
dic_cn_str = dic_cn_str + " if(obj_query.value==''){ "
dic_cn_str = dic_cn_str + "  alert('단어를 입력하세요.'); "
dic_cn_str = dic_cn_str + "  obj_query.focus(); "
dic_cn_str = dic_cn_str + "  return; "
dic_cn_str = dic_cn_str + " } "
dic_cn_str = dic_cn_str + " window.open('http://cndic.naver.com/mini/search/all?q='+obj_query.value+'','_pocket_quiz_pop_','width=300,height=400,status=yes,scrollbars=no,toolbar=no'); "
dic_cn_str = dic_cn_str + "} "
dic_cn_str = dic_cn_str + "</script> "
dic_cn_str = dic_cn_str + "<form name='form'><xmp id='div_xmp'>MZM</xmp> "
dic_cn_str = dic_cn_str + "<input type='text' name='query' width='100px' value=''> "
dic_cn_str = dic_cn_str + "<input type='button' value='찾기' onclick='show_dic()'> "
dic_cn_str = dic_cn_str + "</form> "
dic_cn_str = dic_cn_str + "</body> "
dic_cn_str = dic_cn_str + "</html>"

tts_ko_str = ""
tts_ko_str = tts_ko_str + vbNewLine + "<html lang=""ko-KR"">"
tts_ko_str = tts_ko_str + vbNewLine + "    <head>"
tts_ko_str = tts_ko_str + vbNewLine + "        <meta charset=""UTF-8"">"
'tts_ko_str = tts_ko_str + vbNewLine + "        <meta http-equiv=""refresh"" content=""0;url=http://translate.google.com/translate_tts?ie=UTF-8&tl=ko&q=MZM"">"
tts_ko_str = tts_ko_str + vbNewLine + "        <meta http-equiv=""refresh"" content=""0;url=" & Module1.TTS_GLOBAL_URL & """>"
tts_ko_str = tts_ko_str + vbNewLine + "        <title></title>"
tts_ko_str = tts_ko_str + vbNewLine + "    </head>"
tts_ko_str = tts_ko_str + vbNewLine + "    <body>"
tts_ko_str = tts_ko_str + vbNewLine + "    </body>"
tts_ko_str = tts_ko_str + vbNewLine + "</html>"
  'document.location.href=;

tts_eng_str = ""
tts_eng_str = tts_eng_str + "<html> <head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"
tts_eng_str = tts_eng_str + "</head><script>"
tts_eng_str = tts_eng_str + "function body_onload()"
tts_eng_str = tts_eng_str + "{   location.href=""http://202.131.25.111/tts/tts.cgi?spk_id=100&text_fmt=0&pitch=100&volume=100&speed=100&wrapper=0&enc=0&text=MZM%3c%76%74%6d%6c%5f%70%61%75%73%65%20%74%69%6d%65%3d%22%35%30%30%22%2f%3e"";}"
tts_eng_str = tts_eng_str + "</script> <body onload='body_onload()'> <body> </html>"




    Dim fLogin As New frmLogin
    cmd = Command
If App.PrevInstance Then
    MsgBox "이미 프로그램이 실행중입니다.", vbExclamation
    Exit Sub
End If



If InStr(LCase(Command), "/noupdate") = 0 Then
    CheckFtp Command
End If
'RegisterDataSource

    GPWD = GPWD1 & GPWD2 & GPWD3
    
    
    'STRCON = "Provider=MySqlProv;Data Source=test;Integrated Security="""";Password="""";User ID=root;Location=localhost;Extended Properties="""""

    TmrAfterTTS_exit = True
    fLogin.Show vbModal
    TmrAfterTTS_exit = False
        
'        If InStr(LCase(Command), "/mdb") > 0 Then
'            If Right(App.Path, 1) = "\" Then
'                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
'
'            Else
'                '본사 계정 학교마다 2개를 바꿀것. db명과 user명 (중요!) 암호는 넣으면 암됨(왜냐. ODBC드라이버가 지원해주지 못함)
'                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & "\" & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
'            End If
'        Else
'            STRCON = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mysql;DESC=;DATABASE=" + getParamValue(Command, "d", "pocket_znlwm_000001") + ";SERVER=" + getParamValue(Command, "h", "127.0.0.1") + ";UID=" + getParamValue(Command, "u", "changeit") + ";PASSWORD=" + getParamValue(Command, "p", "") + ";PORT=" + getParamValue(Command, "P", "3306") + ";SOCKET=;OPTION=" + getParamValue(Command, "o", "2097155") + ";STMT=;"""
'
'        End If
        If Len(STRCON) > 0 Then
            If dbversioncheck Then
                If Not fLogin.OK Then
                    '로그온에 실패하여 응용 프로그램을 종료합니다.
                    If gChangUser = False Then
                        End
                    End If
                End If
                Unload fLogin
            
                If fMainForm Is Nothing Then
                    Set fMainForm = New frmMain
                    fMainForm.Show
                Else
                    cFTP.Profile.setPoP3 = False '종료하시겠습니까 라는 메시지가 안나오게 한다.
                    Unload fMainForm
                    Set fMainForm = Nothing
                    Set fMainForm = New frmMain
                    fMainForm.Show
                    
                    Con_Open '다른사용자를 등록시키는 동안 잠깐동안 세션이 끊기지만 복구된다.
                    sSql = "update tu01 set islogon=1 where userid='" & gUserid & "' and con_time='" & gSessionId & "'"
                    Fn_SQLExec (sSql)
                    
                End If
            Else
                MsgBox "새로운 버전의 프로그램을 설치하세요.(호환성없음 으로 인한 문제 발생)", vbCritical
            End If
        End If
End Sub


Sub CheckFtp(ByVal cmd As String)

Dim addr As String
Dim usr As String
Dim pass As String

Dim vData As Variant
Dim ocmd As String '오리지날 cmd

ocmd = cmd

cmd = LCase(cmd)
'If InStr(cmd, "/ftp") > 0 Then
    'cmd = Trim(Replace(cmd, "/ftp ", "", 1, 1))
    
    'vData = Split(cmd, " ")
    
    'If UBound(vData) >= 2 Then
        addr = getParamValue(cmd, "x", getParamValue(cmd, "h", "127.0.0.1"))
        usr = getParamValue(cmd, "y", "pquiz")
        pass = getParamValue(cmd, "z", "user")
'    ElseIf UBound(vData) = 1 Then
'        addr = vData(0)
'        usr = vData(1)
'        pass = "user"
'
'    ElseIf UBound(vData) = 0 Then
'        addr = vData(0)
'        usr = "pquiz"
'        pass = "user"
'
'    End If

'Else
'    If InStr(cmd, "/local") > 0 Then
'
'        addr = "localhost"'hosts 파일에 도메인을 등록하면 필요가 없다.
'    Else
'        addr = getParamValue(cmd, "h", "59.150.9.39") ' "59.150.9.39" '"pquiz.codns.com"
    
'    End If
        
'        usr = "pquiz"
'        pass = "user"

'End If

'------------------------------------------------------------------------------
'FTP를 통해서 업데이트기초정보(pupdate.zip) 파일을 다운로드한다.(서버:pupdate.ds 로컬:pupdate.dl 압축:pupdate.zip==>pupdate.ds)
'------------------------------------------------------------------------------
If FTP_FileCopy(emDownload, addr, usr, pass, "pupdate.zip", getPerfectFile(App.Path, "pupdate.zip"), True) = FL_SUCCESS Then
    Call DeleteFile(getPerfectFile(App.Path, "pupdate.ds"))
    Call UnZipfile("pupdate.zip")
Else
    'MsgBox "ftp download 오류"
    Debug.Assert False
    Exit Sub '오류
End If


'------------------------------------------------------------------------------
'버젼을 비교한다.
'------------------------------------------------------------------------------
'1. 업데이트 프로그램을 다운로드할 필요가 있는지를 비교한다.(pUpdate.zip: pUpdate.exe)
'2. 공지사항 파일을 다운로드할 필요가 있는지를 비교한다.(pNotice.zip: notice.txt)
'3. 응용프로그램을 다운로드할 필요가 있는지를 비교한다.(pocketquiz.zip: pocketquiz.exe)

Dim RetStr As String * 255
Dim Ret As Long

Dim LOCALUPDATEFILE As String
Dim REMOTEUPDATEFILE As String

LOCALUPDATEFILE = "pupdate.dl"

REMOTEUPDATEFILE = "pupdate.ds"


Dim retVal_LU As String
Dim retVal_LN As String
Dim retVal_LA As String

Dim retVal_RU As String
Dim retVal_RN As String
Dim retVal_RA As String

If isFile(getPerfectFile(App.Path, "pupdate.dl")) Then
    
    Ret = GetPrivateProfileString("VERSION", "ver_update", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LU = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_notice", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LN = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_app", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LA = LeftH(RetStr, Ret)
 
    Ret = GetPrivateProfileString("VERSION", "ver_update", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RU = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_notice", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RN = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_app", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RA = LeftH(RetStr, Ret)

    If retVal_LU < retVal_RU Or (Not isFile(getPerfectFile(App.Path, "putool.exe"))) Then
        
        '------------------------------------------------------------------------------
        'FTP를 통해서 putool.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
        
        If FTP_FileCopy(emDownload, addr, usr, pass, "putool.zip", getPerfectFile(App.Path, "putool.zip"), True) = FL_SUCCESS Then
'            Call Kill("putool.exe")
            Call DeleteFile(getPerfectFile(App.Path, "putool.exe"))
            If UnZipfile("putool.zip") Then
                WritePrivateProfileString "VERSION", "ver_update", retVal_RU, getPerfectFile(App.Path, LOCALUPDATEFILE)
            End If
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
    
    End If

    If retVal_LN < retVal_RN Or (Not isFile(getPerfectFile(App.Path, "pnotice.txt"))) Then
        
        '------------------------------------------------------------------------------
        'FTP를 통해서 pnotice.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
        If FTP_FileCopy(emDownload, addr, usr, pass, "pnotice.zip", getPerfectFile(App.Path, "pnotice.zip"), True) = FL_SUCCESS Then
'            Call Kill("pnotice.txt")
            Call DeleteFile(getPerfectFile(App.Path, "pnotice.txt"))
            Call UnZipfile("pnotice.zip")
            Call Shell("notepad " & getPerfectFile(App.Path, "pnotice.txt"), vbNormalFocus)
            WritePrivateProfileString "VERSION", "ver_notice", retVal_RN, getPerfectFile(App.Path, LOCALUPDATEFILE)
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
        
    End If

    If retVal_LA < retVal_RA Then
        Call Shell(getPerfectFile(App.Path, "putool.exe -x" + getParamValue(Command, "x", "127.0.0.1") & " -y" & usr & " -z" & pass & " " & ocmd), vbNormalFocus)
        End
    End If

Else
    Ret = GetPrivateProfileString("VERSION", "ver_update", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LU = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_notice", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LN = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_app", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LA = LeftH(RetStr, Ret)
 
    Ret = GetPrivateProfileString("VERSION", "ver_update", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RU = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_notice", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RN = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_app", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RA = LeftH(RetStr, Ret)
    
        '------------------------------------------------------------------------------
        'FTP를 통해서 putool.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
        
        If FTP_FileCopy(emDownload, addr, usr, pass, "putool.zip", getPerfectFile(App.Path, "putool.zip"), True) = FL_SUCCESS Then
'            Call Kill("putool.exe")
            Call DeleteFile(getPerfectFile(App.Path, "putool.exe"))
            Call UnZipfile("putool.zip")
            WritePrivateProfileString "VERSION", "ver_update", retVal_RU, getPerfectFile(App.Path, LOCALUPDATEFILE)
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
        
                
        '------------------------------------------------------------------------------
        'FTP를 통해서 pnotice.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
        If FTP_FileCopy(emDownload, addr, usr, pass, "pnotice.zip", getPerfectFile(App.Path, "pnotice.zip"), True) = FL_SUCCESS Then
'            Call Kill("pnotice.txt")
            Call DeleteFile(getPerfectFile(App.Path, "pnotice.txt"))
            Call UnZipfile("pnotice.zip")
            Call Shell("notepad " & getPerfectFile(App.Path, "pnotice.txt"), vbNormalFocus)
            WritePrivateProfileString "VERSION", "ver_notice", retVal_RN, getPerfectFile(App.Path, LOCALUPDATEFILE)
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
        '/ftp
        Call Shell(getPerfectFile(App.Path, "putool.exe") & " -x" & addr & " -y" & usr & " -z" & pass & " " & ocmd, vbNormalFocus)
        End

End If


'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

End Sub

Function UnZipfile(fname As String) As Boolean
'    Dim pZ As New PSZLib
'    Dim pzFileCol As PSZipFileCol
'    Dim PZFile As PSZipFile
    UnZipfile = Show_ZipContents2(getPerfectFile(App.Path, fname), App.Path)
'    pZ.InFileName = getPerfectFile(App.Path, fname)
'    pZ.RootBasePath = App.Path
'
'    pZ.GetZipFileInfos
'
'    Set pzFileCol = pZ.Item
'    Dim i As Long
'    For i = 1 To pzFileCol.Count
'        UnZipfile = True
'        Set PZFile = pzFileCol.Item(i)
'        pZ.ResetFileList
'        pZ.AddFileList PZFile.FileName
'
'
'        pZ.InFileName = getPerfectFile(App.Path, fname)
'        pZ.RootBasePath = App.Path
'
'        Call pZ.Decompress(0)
'    Next

End Function
Private Function Show_ZipContents2(fnamefull As String, ToDir As String) As Boolean
    Dim X As Long
    Dim Enc As String
    Dim DirCnt As Long
    Dim FileCnt As Long
    Dim Temp As Long
    ZF.Get_Contents (fnamefull)
'    For X = 1 To lstInZip.ListItems.Count
'        lstInZip.ListItems(X).Selected = False
'    Next
    For X = 1 To 1
'        With lstInZip
            Enc = " "
            If ZF.Encrypted(X) Then Enc = "+"
            If Not ZF.IsDir(X) Then
                FileCnt = FileCnt + 1
'                .ListItems.Add X, , Enc & ZF.FileName(X)
'                .ListItems(X).SubItems(1) = ZF.Method(X)
                Temp = ZF.CRC32(X)
                If Temp = 0 Then
'                    .ListItems(X).SubItems(2) = "?"
                Else
'                    .ListItems(X).SubItems(2) = Hex(Temp)
                End If
                Temp = ZF.Compressed_Size(X)
'                If Temp = 0 Then
'                    .ListItems(X).SubItems(3) = "?"
'                Else
'                    .ListItems(X).SubItems(3) = Temp
'                End If
'                .ListItems(X).SubItems(4) = ZF.UnCompressed_Size(X)
'                .ListItems(X).SubItems(5) = ZF.FileDateTime(X)
            Else
                DirCnt = DirCnt + 1
'                .ListItems.Add X, , Enc & ZF.FileName(X)
'                .ListItems(X).SubItems(1) = ZF.Method(X)
'                .ListItems(X).SubItems(2) = "Directory Entry"
'                .ListItems(X).SubItems(3) = "Directory Entry"
'                .ListItems(X).SubItems(4) = "Directory Entry"
'                .ListItems(X).SubItems(5) = ZF.FileDateTime(X)
            End If
'        End With
    Next
    If ZF.FileCount > 0 Then
'        lblHeadLine.Caption = "Contents of " & Filetype(PackFileType) & " file " & _
'                              FileList.FileName & " -> " & _
'                              DirCnt & " directories and " & _
'                              FileCnt & " files"
    Else
'        lblHeadLine.Caption = "Not supported format"
    End If
'    If ZF.CanUnpack Then
'        btnUnzip.Enabled = True
'    Else
'        btnUnzip.Enabled = False
'    End If
    
    Call UnZipAll(ZF.FileCount, ToDir)
    Show_ZipContents2 = True
End Function

Private Sub UnZipAll(ByVal cnt As Long, ToDir As String)
    Dim FileUnzip() As Boolean
    Dim Sel As Boolean
    Dim X As Long
    Dim retVal As Boolean
        ReDim FileUnzip(cnt)
        For X = 1 To cnt
                FileUnzip(X) = True
        Next
    If ToDir = "" Then
        MsgBox "No path to store files", vbCritical
        Exit Sub
    End If
'    MousePointer = vbHourglass
    retVal = ZF.UnPack(FileUnzip, ToDir)
'    MousePointer = vbNormal
End Sub
Function dbversioncheck() As Boolean
    
    Dim sdbid As String
    Dim lver As Long
    
    sdbid = getTableVal("dbid", "ver", "")
    lver = getTableVal("ver", "ver", "")
    
    If sdbid = GDBID Or lver = GDBVER Then
        dbversioncheck = True
    Else
        dbversioncheck = False
        MsgBox "데이터와 프로그램의 버젼이 일치하지 않습니다.", vbCritical
    End If

End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next


    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Long


    '폼의 캡션을 설정합니다.
    frm.Caption = LoadResString(CInt(frm.Tag))
    

    '글꼴을 설정합니다.
    Set fnt = frm.Font
    fnt.Name = LoadResString(20)
    fnt.Size = CInt(LoadResString(21))
    

    '메뉴 항목의 Caption 속성과
    '다른 모든 컨트롤의 Tag 속성을 사용하여
    '컨트롤의 캡션을 설정합니다.
    For Each ctl In frm.Controls
        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next

End Sub

Public Function Con_Open() As Boolean
On Error GoTo ErrTrap
If Con.State = 0 Then


RegisterDataSource
'frmConnecting.Show vbModeless '모달폼에서 모달이아닌폼을 호출하면 오류가됨.
'DoEvents
'CenterForm frmConnecting
'DoEvents
'AlwaysOnTop frmConnecting, True
'DoEvents
'   MsgBox "연결중입니다. 30초정도 걸립니다... (확인 클릭)", vbExclamation
    Con.Open STRCON
   'MsgBox "연결완료               ", vbExclamation
'   Unload frmConnecting
DoEvents

'    Con.CommandTimeout = 300
End If
Exit Function
ErrTrap:
    Con_Open = True
    MsgBox err.Description, vbCritical
    If err.Number = 3146 Then
    '?err.Number
    '3146
    'Print err.Description
    'ODBC--호출이 실패하였습니다.
        If MsgBox("MySQL 커넥터를 설치하시겠습니까?" + vbNewLine + "설치후 프로그램을 다시 시작해 주세요.", vbExclamation + vbYesNo, "설치안내") = vbYes Then
            Call Shell(App.Path + "\mysql-connector-odbc-3.51.23-win32.exe", vbNormalFocus)
        Else
            MsgBox "MY SQL ODBC 설치가 올바로 되지 않았습니다. MYSQL ODBC를 (재)설치 하세요.(pc restart > install)", vbExclamation
        End If
        
    Else
        MsgBox "인터넷 연결에 문제가 있거나 서비스 점검중입니다.", vbExclamation
        If (IDEMODE) Then
            If MsgBox("종료하시겠습니까?", vbYesNo + vbQuestion) = vbNo Then
                Resume
            Else
                End
            End If
        End If
    End If
    End '프로그램 종료
End Function

Private Sub RegisterDataSource()
Dim sAttribs As String, lReturn As Long

On Error GoTo Err1

' 키워드 문자열을 작성합니다.
'sAttribs = "Description=포켓퀴즈" & Chr$(0) & _
'"DSN=mysql2" & Chr$(0) & _
'"DEFAULTDIR=" & App.Path & "\Data\" & Chr$(0) & _
'"EXTENSIONS=csv" & Chr$(0)

        Dim att As String
        Dim mydb As Database

        att = "Description =" & Chr$(13)
        att = att & "Database=" & Chr$(13)   ' Build keywords string.
        att = att & "Option=" & Chr$(13)
        att = att & "Port=" & Chr$(13)
        
        'If InStr(LCase(cmd), "/local") = 0 Then
           att = att & "Server=" & Chr$(13)
'        Else
'           att = att & "Server=localhost" & Chr$(13)
'        End If
'
        'att = att & "User=ctm000"
        att = att & "User="

        ' Update ODBC.INI.
        
    'Call err.Raise(vbObjectError + 100, "Module1.calGorgetColor", "입력범위오류")
        RegisterDatabase "mysql", "MySQL ODBC 3.51 Driver", True, att 'dao360.dll

Exit Sub
Err1:
On Error GoTo Err2
    '레지스트리에 등록한다
    '[HKEY_CURRENT_USER\Software\ODBC\ODBC.INI\ODBC Data Sources]"mysql"="MySQL ODBC 3.51 Driver"
    
'그래도 에러가 나면 무시한다. 왜냐하면 처음에 인스톨할 때 이미 등록 했기 때문이다.
'    Dim lResult As Long
'    Dim hTopKey As Long, hKey As Long, lDisposition As Long
    Dim sRegPath As String
'    Dim sa As SECURITY_ATTRIBUTES
'
'    hTopKey = HKEY_CURRENT_USER
    sRegPath = "Software\ODBC\ODBC.INI\ODBC Data Sources"
'
'    sa.nLength = Len(sa)
'    sa.bInheritHandle = CLng(True)
'
'    lResult = RegCreateKeyEx(hTopKey, sRegPath, 0, vbNullString, REG_OPTION_NON_VOLATILE, _
'                             KEY_ALL_ACCESS, sa, hKey, lDisposition)
'    If lResult = ERROR_SUCCESS Then
'       If lDisposition = REG_CREATED_NEW_KEY Then
'          ' Assign default values
'       ElseIf lDisposition = REG_OPENED_EXISTING_KEY Then
'          ' Retrieve values from existing keys
'       End If
'
       Dim str As String
    
       str = "MySQL ODBC 3.51 Driver" & vbNullChar
       
       Call SetRegValue(HKEY_CURRENT_USER, sRegPath, "mysql", REG_SZ, UnicodeStringToBytes(str), Len(str))

'    Else
'       'MsgBox "Error " & lResult
'    End If
    
Err2:
'lReturn = SQLConfigDataSource(0&, ODBC_ADD_DSN, "Microsoft Text Driver (*.txt; *.csv)", sAttribs)
End Sub

'http://www.vbforums.com/showthread.php?402133-RegSetValueEx-dword-problem
'1.Call SetRegLongValue(LocalMachine, "Software\MyProduct\SomeKey", "Some DWORD Value", 12345)
Private Function SetRegValue( _
  ByVal hive As Long, ByVal subKeyPath As String, _
  ByVal ValueName As String, ByVal dataType As Long, _
  ByRef Data() As Byte, ByVal dataLength As Long) As Boolean
  
  Dim lngKey As Long
  Dim lngResult As Long
  Dim lngSecurity As Long
  
  ' Request write access to the registry key
  lngSecurity = KEY_WRITE
  
  ' Open the key in the specified hive
  lngResult = RegOpenKeyEx( _
    hive, subKeyPath, 0, lngSecurity, lngKey)
  
  ' If we couldn't open the key, abort
  If lngResult <> ERROR_SUCCESS Then
    SetRegValue = False
    Exit Function
  End If
  
  ' Set the value in the registry
  lngResult = RegSetValueEx( _
    lngKey, ValueName, 0, dataType, Data(LBound(Data)), dataLength)
    
  ' If succesful, return True
  If lngResult = ERROR_SUCCESS Then
    SetRegValue = True
  End If
  
  ' Close the registry key
  Call RegCloseKey(lngKey)
End Function

'http://www.arcatapet.com/software/vbregset.cfm
Sub RegSetValue(H_KEY&, RSubKey$, ValueName$, RegValue$)
    'H_KEY must be one of the Key Constants
    Dim lRtn&         'returned by registry functions, should be 0&
    Dim hKey&         'return handle to opened key
    Dim lpDisp&
    Dim Sec_Att As SECURITY_ATTRIBUTES
    Sec_Att.nLength = 12&
    Sec_Att.lpSecurityDescriptor = 0&
    Sec_Att.bInheritHandle = False
    If RegValue = "" Then RegValue = " "
    
        lRtn = RegCreateKeyEx(H_KEY, RSubKey, 0&, "", 0&, KEY_WRITE, Sec_Att, hKey, lpDisp)
        If lRtn <> 0 Then
            Exit Sub       'No key open, so leave
        End If
        lRtn = RegSetValueEx(hKey, ValueName, 0&, REG_SZ, ByVal RegValue, CLng(Len(RegValue) + 1))
        lRtn = RegCloseKey(hKey)
End Sub


Public Function Con_Close() As Boolean

Exit Function

On Error GoTo ErrTrap
If Con.State = 1 Then
    Con.Close
End If

Exit Function

ErrTrap:
    Con_Close = True
    MsgBox err.Description, vbCritical
    
End Function

Public Function GETYMD() As String
GETYMD = Format(Now, "YYYYMMDD")
End Function

Public Function GETTIME() As String
GETTIME = Format(Now, "hhMMSS")
End Function

'=================================================
'시험지 만들기
'=================================================
Public Function insert_pocketinfo(pocketnm As String) As Boolean
Dim RS2 As New ADODB.Recordset
Dim cond As String
Dim OBJTABLE As String

Dim makecnt As Long

Dim SSQL1 As String
Dim SSQL2 As String
Dim SSQL3 As String

Dim i As Long
Dim affected As Long

Dim ymd As String

Screen.MousePointer = vbHourglass
ymd = GETYMD()


Dim chasu As Long

SSQL1 = "SELECT MAX(CHASU) FROM TP02 WHERE USERID='" & gUserid & "' AND POCKETNM='" & pocketnm & "' "
RS2.Open SSQL1, STRCON, 1, 3
If Not RS2.EOF Then
    If IsNull(RS2(0)) Then
        chasu = 0
    Else
        chasu = RS2(0) + 1
    End If
    
End If
RS2.Close


SSQL1 = "SELECT COND ,xm,chkm FROM TP01 WHERE POCKETNM='" & pocketnm & "'"
RS2.Open SSQL1, Con, 1, 3


Dim xm As Long
Dim chkm As Long
Dim img As Long
Dim selimg As Long


cond = RS2(0)
xm = RS2(1)
chkm = RS2(2)

If xm = 0 And chkm = 0 Then
    img = 1
    selimg = 2
    
ElseIf xm > 0 And chkm = 0 Then
    img = 3
    selimg = 4

ElseIf xm = 0 And chkm > 0 Then
    img = 5
    selimg = 6
Else
    img = 7
    selimg = 8
End If

OBJTABLE = "TU02" 'RS2(1)
RS2.Close

Dim vData As Variant
vData = Split(cond, "|")

If UBound(vData) = 1 Then
    SSQL1 = "select * from vq01 where " & vData(0) & " order by seq"
Else
    SSQL1 = "SELECT * FROM " & OBJTABLE & " WHERE userid='" & gUserid & "' and " & cond & " order by seq"
End If
Con.BeginTrans
RS2.Open SSQL1, Con, 1, 3

i = 0
makecnt = 0
Do Until RS2.EOF
    i = i + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pocketnm & "'," & chasu & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    RS2.MoveNext
Loop

RS2.Close

Screen.MousePointer = vbDefault
If makecnt < 1 Then
    Con.RollbackTrans
    If makecnt = 0 Then
        MsgBox "시험지 만들 데이터가 없습니다.", vbExclamation
'    Else
'        MsgBox "시험지 만들 데이터가 너무적습니다.", vbExclamation
    End If
    insert_pocketinfo = False
Else
    Dim maxCode As Long
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pocketnm & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "',0," & xm & " ," & chkm & "," & img & " ," & selimg & ",'0','0')"
    Fn_SQLExec (SSQL1)

    Dim makeOrder As Long
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    End If

    If makeOrder < 2 Then
        makeOrder = 2
    End If
    
    
    SSQL1 = "insert into tu03 values('" & pocketnm & "'," & chasu & ",'" & gUserid & "',1,1," & makeOrder & ")"
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0

    insert_pocketinfo = True
    Con.CommitTrans
    MsgBox makecnt & " 건을 만들었습니다", vbInformation
End If
End Function
Function getMaxTableVal(clm As String, tbl As String, strWhere) As Long
Dim lsql As String
Dim lRs As ADODB.Recordset

'##1
Dim conbef As Long
If Con.State = 1 Then
    conbef = 1
Else
    Con_Open
    conbef = 0
End If

lsql = "select max(" & clm & ") from " & tbl & " " & strWhere
Set lRs = Fn_SQLExec(lsql).rs '

If lRs.EOF Then
    getMaxTableVal = 1
Else
    If IsNull(lRs(0)) Then
        getMaxTableVal = 1
    Else
        getMaxTableVal = lRs(0) + 1
    End If
End If

'##2
If conbef = 1 Then
Else
    Con_Close
End If

End Function

Function getTableVal(clm As String, tbl As String, strWhere) As Variant
Dim lsql As String
Dim lRs As ADODB.Recordset

'##1
Dim conbef As Long
If Con.State = 1 Then
    conbef = 1
Else
    Con_Open
    conbef = 0
End If

lsql = "select " & clm & " from " & tbl & " " & strWhere
Set lRs = Fn_SQLExec(lsql, , , True).rs '

If Not lRs.EOF Then
    On Error Resume Next
    getTableVal = lRs(0).Value
    If err.Number <> 0 Then
    getTableVal = ""
    End If
End If


'##2
If conbef = 1 Then
Else
'    Con_Close
End If

End Function


'=================================================
'시험지 만들기2
'=================================================
Public Function insert_pocketinfo2(pn As String, cnt As Long) As Boolean
Dim RS2 As New ADODB.Recordset

Dim makecnt As Long

Dim SSQL1 As String
Dim SSQL2 As String
Dim SSQL3 As String

Dim ymd As String
Dim i As Long
Dim affected As Long

Dim OBJTABLE As String

ymd = GETYMD()


Dim chasu As Long

SSQL1 = "SELECT MAX(CHASU) FROM TP02 WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn & "' "
RS2.Open SSQL1, Con, 1, 3
If Not RS2.EOF Then
    If IsNull(RS2(0)) Then
        chasu = 0
    Else
        chasu = RS2(0) + 1
    End If
    
End If
RS2.Close

OBJTABLE = "TU02" 'RS2(1)


SSQL1 = "SELECT * FROM " & OBJTABLE & " WHERE O+X>0 AND USERID='" & gUserid & "' ORDER BY UPDATE_YMD ASC"

Con.BeginTrans
RS2.Open SSQL1, Con, 1, 3

i = 0
makecnt = 0
Do Until RS2.EOF Or i >= cnt
    i = i + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    RS2.MoveNext
Loop

RS2.Close

If makecnt < 1 Then
    Con.RollbackTrans
    If makecnt = 0 Then
        MsgBox "시험지 만들 데이터가 없습니다.", vbExclamation
'    Else
'        MsgBox "시험지 만들 데이터가 너무적습니다."
    End If
    insert_pocketinfo2 = False
Else
    Dim maxCode As Long
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "',0,0,0,1,2,'0','0')"
    Fn_SQLExec (SSQL1)

    Dim makeOrder As Long
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    End If

    If makeOrder < 2 Then
        makeOrder = 2
    End If

    SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "',1,1," & makeOrder & ")"
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0
    Con.CommitTrans
    insert_pocketinfo2 = True
    MsgBox makecnt & " 건을 만들었습니다", vbExclamation
End If
End Function


'=================================================
'시험지 만들기3
'=================================================
Public Function insert_pocketinfo3(pn As String, cnt As Long, STRO As String, STRX As String) As Boolean
Dim RS2 As New ADODB.Recordset

Dim makecnt As Long

Dim SSQL1 As String
Dim SSQL2 As String
Dim SSQL3 As String

Dim ymd As String
Dim OBJTABLE As String
Dim i As Long
Dim affected  As Long

ymd = GETYMD()


Dim chasu As Long

SSQL1 = "SELECT MAX(CHASU) FROM TP02 WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn & "' "
RS2.Open SSQL1, Con, 1, 3
If Not RS2.EOF Then
    If IsNull(RS2(0)) Then
        chasu = 0
    Else
        chasu = RS2(0) + 1
    End If
    
End If
RS2.Close

OBJTABLE = "TU02" 'RS2(1)


SSQL1 = "SELECT * FROM " & OBJTABLE & " WHERE USERID='" & gUserid & "' and o+x>0 ORDER BY x desc,o asc, UPDATE_YMD ASC"
Con.BeginTrans
RS2.Open SSQL1, Con, 1, 3

i = 0
makecnt = 0
Do Until RS2.EOF
    i = i + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    If i >= cnt Then Exit Do
    RS2.MoveNext
Loop

RS2.Close

If makecnt < 1 Then
    Con.RollbackTrans
    If makecnt = 0 Then
        MsgBox "시험지 만들 데이터가 없습니다.", vbExclamation
'    Else
'        MsgBox "시험지 만들 데이터가 너무적습니다."
    End If
    insert_pocketinfo3 = False
Else

    Dim maxCode As Long
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "',0,0,0,1,2,'0','0')"
    Fn_SQLExec (SSQL1)

    Dim makeOrder As Long
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    End If

    If makeOrder < 2 Then
        makeOrder = 2
    End If
   
    SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "',1,1," & makeOrder & ")"
        
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0

    insert_pocketinfo3 = True
    Con.CommitTrans
    MsgBox makecnt & " 건을 만들었습니다", vbExclamation
End If
End Function

'=================================================
'시험지 만들기4
'=================================================
Public Function insert_pocketinfo4(pn As String, cnt As Long, lst As ListBox) As Boolean
Dim RS2 As New ADODB.Recordset
Dim SSQL1 As String

Dim makecnt As Long

Dim ymd As String
Dim OBJTABLE As String
Dim i As Long
Dim affected  As Long

ymd = GETYMD()


Dim chasu As Long

SSQL1 = "SELECT MAX(CHASU) FROM TP02 WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn & "' "
RS2.Open SSQL1, Con, 1, 3
If Not RS2.EOF Then
    If IsNull(RS2(0)) Then
        chasu = 0
    Else
        chasu = RS2(0) + 1
    End If
    
End If
RS2.Close

OBJTABLE = "TU02" 'RS2(1)

SSQL1 = "SELECT x/(x+o), " & OBJTABLE & ".* FROM " & OBJTABLE & " WHERE USERID='" & gUserid & "' and o+x>0 "

If Not lst.Selected(0) Then
    SSQL1 = SSQL1 & vbNewLine & "and subj in (" & selectSeries(lst, vbTab) & ")"
End If
SSQL1 = SSQL1 & " ORDER BY 1 desc"
Con.BeginTrans
RS2.Open SSQL1, Con, 1, 3

i = 0
makecnt = 0
Do Until RS2.EOF
    i = i + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    If i >= cnt Then Exit Do
    RS2.MoveNext
Loop

RS2.Close

If makecnt < 1 Then
    Con.RollbackTrans
    If makecnt = 0 Then
        MsgBox "시험지 만들 데이터가 없습니다.", vbExclamation
'    Else
'        MsgBox "시험지 만들 데이터가 너무적습니다."
    End If
    insert_pocketinfo4 = False
Else
    Dim maxCode As Long
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "',0,0,0,1,2,'0','0')"
    Fn_SQLExec (SSQL1)

    Dim makeOrder As Long
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    End If

    If makeOrder < 2 Then
        makeOrder = 2
    End If

    SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "',1,1," & makeOrder & ")"
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0

    insert_pocketinfo4 = True
    Con.CommitTrans
    MsgBox makecnt & " 건을 만들었습니다", vbExclamation
End If
End Function

'=================================================
'시험지 만들기5
'=================================================
Public Function insert_pocketinfo5(pn As String, Ch As String, idx As Integer) As Boolean
Dim RS2 As New ADODB.Recordset
Dim SSQL1 As String

Dim makecnt As Long

Dim ymd As String
Dim OBJTABLE As String
Dim i As Long
Dim affected  As Long

ymd = GETYMD()

Dim code As Long
Dim pcode As Long


Dim lsql As String
Dim chasu As Long

lsql = "select code from tp02 where userid='" & gUserid & "' and pocketnm='" & pn & "' and chasu=" & Ch & " "

pcode = Fn_SQLExec(lsql).rs(0)

code = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")

chasu = getMaxTableVal("chasu", "tp02", "WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn & "'")

OBJTABLE = "Tp03" 'RS2(1)

SSQL1 = "SELECT * FROM " & OBJTABLE
SSQL1 = SSQL1 & " WHERE USERID='" & gUserid & "' and pocketnm='" & pn & "' and chasu=" & Ch

Dim pn2 As String

Select Case idx
Case 0
    SSQL1 = SSQL1 & " and x>0 "
    pn2 = "틀린문제"
Case 1
    SSQL1 = SSQL1 & " and chk>0 "
    pn2 = "체크문제"
Case 2
    SSQL1 = SSQL1 & " and o+x=0 "
    pn2 = "안푼문제"
Case 3
    SSQL1 = SSQL1 & " and o>0 and x=0 "
    pn2 = "맞은문제"
Case 4
    SSQL1 = SSQL1 & " and x+chk>0"
    pn2 = "틀린+체크"
Case Else
    SSQL1 = SSQL1 & " and (x+chk>0 or o+x=0) "
    pn2 = "틀린+체크+안품"
End Select

Con.BeginTrans
RS2.Open SSQL1, Con, 1, 3

i = 0
makecnt = 0

Dim ch2 As Long
Dim ch3 As Long


ch2 = getMaxTableVal("chasu", "tp02", "WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn2 & "'")
ch3 = getMaxTableVal("chasu", "tp03", "WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn2 & "'")

If ch3 > ch2 Then
    Debug.Assert False
    ch2 = ch3
    'Debug.Assert False
End If

Do Until RS2.EOF
    i = i + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn2 & "'," & ch2 & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    'If i >= cnt Then Exit Do
    RS2.MoveNext
Loop

RS2.Close

If makecnt < 1 Then
    Con.RollbackTrans
    If makecnt = 0 Then
        MsgBox "시험지 만들 데이터가 없습니다.", vbExclamation
'    Else
'        MsgBox "시험지 만들 데이터가 너무적습니다.", vbExclamation
    End If
    insert_pocketinfo5 = False
Else
    Dim maxCode As Long
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn2 & "'," & ch2 & ",'" & gUserid & "',0,'" & ymd & "'," & pcode & ",0,0,1,2,'0','0')"
    Fn_SQLExec (SSQL1)

    Dim makeOrder As Long
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    End If
    
    If makeOrder < 2 Then
        makeOrder = 2
    End If
    
    SSQL1 = "insert into tu03 values('" & pn2 & "'," & ch2 & ",'" & gUserid & "',1,1," & makeOrder & ")"
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0

    insert_pocketinfo5 = True
'    Debug.Assert False '아래 트랜잭션이 커밋이 유효한가 살펴보기 위해 디버그 모드로 전환됨
    Con.CommitTrans
    MsgBox makecnt & " 건을 만들었습니다", vbExclamation
    Module1.vDataGlobal2 = pn2 & " " & ch2
End If
End Function


'=================================================
'시험지 만들기6
'-------------------------------------------------
' 오늘의 필수 복습문제
'=================================================
Public Function insert_pocketinfo6(pn As String, cnt As Long, lst As ListBox) As Boolean
Dim RS2 As New ADODB.Recordset
Dim SSQL1 As String

Dim makecnt As Long

Dim ymd As String
Dim OBJTABLE As String
Dim i As Long
Dim affected  As Long

ymd = GETYMD()

Dim chasu As Long

SSQL1 = "SELECT MAX(CHASU) FROM TP02 WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn & "' "
RS2.Open SSQL1, Con, 1, 3
If Not RS2.EOF Then
    If IsNull(RS2(0)) Then
        chasu = 0
    Else
        chasu = RS2(0) + 1
    End If
    
End If
RS2.Close

OBJTABLE = "TU02" 'RS2(1)

SSQL1 = "SELECT " & OBJTABLE & ".* FROM " & OBJTABLE & " WHERE USERID='" & gUserid & "' and reserve_ymd<'" & ymd & "'"

If Not lst.Selected(0) Then
    SSQL1 = SSQL1 & vbNewLine & "and subj in (" & selectSeries(lst, vbTab) & ")"
End If
SSQL1 = SSQL1 & " ORDER BY reserve_ymd asc"
Con.BeginTrans
RS2.Open SSQL1, Con, 1, 3

i = 0
makecnt = 0
Do Until RS2.EOF
    i = i + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    If i >= cnt Then Exit Do
    RS2.MoveNext
Loop

RS2.Close

If makecnt < 1 Then
    Con.RollbackTrans
    If makecnt = 0 Then
        MsgBox "시험지 만들 데이터가 없습니다.", vbExclamation
'    Else
'        MsgBox "시험지 만들 데이터가 너무적습니다."
    End If
    insert_pocketinfo6 = False
Else
    Dim maxCode As Long
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "',0,0,0,1,2,'0','0')"
    Fn_SQLExec (SSQL1)

    Dim makeOrder As Long
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    End If

    If makeOrder < 2 Then
        makeOrder = 2
    End If

    SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "'," & makecnt & "," & makecnt + 1 & "," & makecnt & ")"
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0

    insert_pocketinfo6 = True
    Con.CommitTrans
    MsgBox makecnt & " 건을 만들었습니다", vbExclamation
End If
End Function

'=================================================
'시험지 만들기7
'-------------------------------------------------
' 선택한 문제지 내의 시험지내에서 필요한 복습문제 생성
'=====================================================
Public Function insert_pocketinfo7(pn As String, Ch As String, idx As Long) As Boolean
Dim RS2 As New ADODB.Recordset
Dim SSQL1 As String
Dim SSQL2 As String
Dim SSQL3 As String

Dim makecnt As Long
Dim dft_cnt As Long
Dim str_sum As Long

Dim ymd As String
Dim OBJTABLE As String
Dim i As Long
Dim affected  As Long

ymd = GETYMD()

Dim code As Long
Dim pcode As Long
Dim ret_cnt As Variant


Dim lsql As String
Dim chasu As Long

lsql = "select code from tp02 where userid='" & gUserid & "' and pocketnm='" & pn & "' and chasu=" & Ch & " "

pcode = Fn_SQLExec(lsql).rs(0)

code = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")

chasu = getMaxTableVal("chasu", "tp02", "WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn & "'")

'OBJTABLE = "Tp03" 'RS2(1)

SSQL1 = "select count(*) from tp03 a,tu02 b where  a.userid='" & gUserid & "' and a.pocketnm='" & pn & "' "
SSQL1 = SSQL1 & " and a.chasu=" & Ch & " and "
SSQL1 = SSQL1 & "a.subj=b.subj and a.seq=b.seq and a.userid=b.userid and b.update_ymd<>'" & ymd & "' and b.reserve_ymd<" & ymd & " and reserve_ymd<>99999999 order by b.gangyek,b.reserve_ymd,b.update_ymd"

SSQL2 = "select count(*) from tp03 a,tu02 b where  a.userid='" & gUserid & "' and a.pocketnm='" & pn & "' "
SSQL2 = SSQL2 & " and a.chasu=" & Ch & " and "
SSQL2 = SSQL2 & "a.subj=b.subj and a.seq=b.seq and a.userid=b.userid and b.update_ymd<>'" & ymd & "' and (b.reserve_ymd<" & date2Str(DateAdd("d", Module1.ALLOW_AFFECT_DATE, Now)) & " or b.update_ymd<" & date2Str(DateAdd("d", Module1.ALLOW_AFFECT_DATE30, Now)) & ") and reserve_ymd<>99999999 order by b.gangyek,b.reserve_ymd,b.update_ymd"

SSQL3 = "select count(*) from tp03 a,tu02 b where  a.userid='" & gUserid & "' and a.pocketnm='" & pn & "' "
SSQL3 = SSQL3 & " and a.chasu=" & Ch & " and "
SSQL3 = SSQL3 & "a.subj=b.subj and a.seq=b.seq and a.userid=b.userid and b.update_ymd<>'" & ymd & "' and reserve_ymd<>'99999999' order by b.gangyek,b.reserve_ymd,b.update_ymd"

dft_cnt = Fn_SQLExec(SSQL1).rs(0)
str_sum = Fn_SQLExec(SSQL3).rs(0)
Do
    ret_cnt = InputBox("문제수 <유효:" & Fn_SQLExec(SSQL2).rs(0) & "><최대:" & str_sum & ">", , dft_cnt)
    If Trim(ret_cnt) = "0" Or Trim(ret_cnt) = "" Then Exit Function
Loop Until IsNumeric(ret_cnt)

dft_cnt = ret_cnt

Dim pn2 As String

Select Case idx
Case 0
    pn2 = "복습문제" & ymd
Case Else
End Select

SSQL1 = "select a.*,if(b.update_ymd<" & date2Str(DateAdd("d", Module1.ALLOW_AFFECT_DATE30, Now)) & " or b.reserve_ymd<" & date2Str(DateAdd("d", Module1.ALLOW_AFFECT_DATE, Now)) & ",0,1) up2 from tp03 a,tu02 b where a.userid='" & gUserid & "' and a.pocketnm='" & pn & "' "
SSQL1 = SSQL1 & " and a.chasu=" & Ch & " and "
SSQL1 = SSQL1 & "a.subj=b.subj and a.seq=b.seq and a.userid=b.userid and b.update_ymd<>'" & ymd & "' and reserve_ymd<>99999999  order by up2,b.reserve_ymd asc "

If InStr(LCase(STRCON), "mysql") > 0 Then
    SSQL1 = SSQL1 & "limit " & dft_cnt
End If

Con.BeginTrans
RS2.Open SSQL1, Con, 1, 3

i = 0
makecnt = 0

Dim ch2 As Long
Dim ch3 As Long

ch2 = getMaxTableVal("chasu", "tp02", "WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn2 & "'")
ch3 = getMaxTableVal("chasu", "tp03", "WHERE USERID='" & gUserid & "' AND POCKETNM='" & pn2 & "'")

If ch3 > ch2 Then
    Debug.Assert False
    ch2 = ch3
    'Debug.Assert False
End If

Do Until RS2.EOF
    i = i + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn2 & "'," & ch2 & ",'" & gUserid & "'," & i & ","
    SSQL1 = SSQL1 & "'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    If i >= dft_cnt Then Exit Do
    RS2.MoveNext
Loop

RS2.Close

If makecnt < 1 Then
    Con.RollbackTrans
    If makecnt = 0 Then
        MsgBox "시험지 만들 데이터가 없습니다.", vbExclamation
'    Else
'        MsgBox "시험지 만들 데이터가 너무적습니다."
    End If
    insert_pocketinfo7 = False
Else
    Dim maxCode As Long
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn2 & "'," & ch2 & ",'" & gUserid & "',0,'" & ymd & "'," & pcode & ",0,0,1,2,'0','0')"
    Fn_SQLExec (SSQL1)

    Dim makeOrder As Long
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    End If

    If makeOrder < 2 Then
        makeOrder = 2
    End If

    SSQL1 = "insert into tu03 values('" & pn2 & "'," & ch2 & ",'" & gUserid & "'," & makecnt & "," & CStr(makecnt + 1) & "," & makecnt & ")"
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0

    insert_pocketinfo7 = True
    Con.CommitTrans
    MsgBox makecnt & " 건을 만들었습니다", vbExclamation
    
    If MsgBox("시험보기를 하시겠습니까? [" & pn2 & "]", vbQuestion + vbYesNo) = vbYes Then
        vDataGlobal = pn2 & " " & ch2
    Else
        vDataGlobal = ""
    End If
    
End If
End Function

'============================================================================
'총계정원장 등록 slow
'============================================================================

Public Sub insert_totalaccount(Optional ByRef pic As PictureBox)
Dim ymd As String
Dim affected  As Long
Dim URS As ut_bRecordSet
Dim allcnt As Long
'Dim cnt As Long
'Dim ppic As PictureBox

If pic Is Nothing Then
'    Set ppic = pic
Debug.Assert False
End If

Screen.MousePointer = vbHourglass
    Dim cnt As Long
    cnt = 0
    sSql = "SELECT SUBJ,SEQ FROM vq01 a where not exists (select subj,seq from tu02  b where userid = '" & gUserid & "' and a.subj=b.subj and a.seq=b.seq)"
    URS = Fn_SQLExec(sSql)
    Set rs = URS.rs
    allcnt = URS.nrow
    'Set rs = Fn_SQLExec(sSql, , , True).rs
    ymd = GETYMD()
    affected = 0
    On Error Resume Next
    
    
    
    
    If pic Is Nothing Then
    
        Do Until rs.EOF()
            sSql = "INSERT INTO TU02 VALUES('" & rs(0) & "'," & rs(1) & ",'" & gUserid & "',0,0,0,'" & ymd & "','99999999',0)"
            
            affected = Fn_SQLExec(sSql, "", True).nrow
    '        Debug.Assert affected = 1
            cnt = cnt + affected
            rs.MoveNext
        Loop
    Else
        Call UpdateStatus(pic, 0, True)
        Do Until rs.EOF()
'            cnt = cnt + 1
            Call UpdateStatus(pic, cnt / allcnt, False)
            sSql = "INSERT INTO TU02 VALUES('" & rs(0) & "'," & rs(1) & ",'" & gUserid & "',0,0,0,'" & ymd & "','99999999',0)"
            
            affected = Fn_SQLExec(sSql, "", True).nrow
    '        Debug.Assert affected = 1
            cnt = cnt + affected
            rs.MoveNext
            
            
            
        Loop
    End If
    On Error GoTo 0
Screen.MousePointer = vbNormal
    MsgBox cnt & " 건의 어카운트가 추가되었습니다.", vbExclamation, "어카운트 추가"

End Sub


'============================================================================
'총계정원장 등록 fast
'============================================================================

Public Sub insert_totalaccount2(Optional ByRef pic As PictureBox)
Dim ymd As String
Dim affected  As Long
Dim URS As ut_bRecordSet
Dim allcnt As Long
Dim cnt As Long

    Screen.MousePointer = vbHourglass
    
'    sSql = "INSERT INTO TU02(subj, seq, userid, o, x, chk, update_ymd, reserve_ymd, gangyek) (select subj,seq,'" & gUserid & "',0,0,0,'" & ymd & "','99999999',0 FROM vq01 a where not exists (select subj,seq from tu02  b where userid = '" & gUserid & "' and a.subj=b.subj and a.seq=b.seq))"
    sSql = "INSERT INTO TU02(subj, seq, userid, o, x, chk, update_ymd, reserve_ymd, gangyek) "
    sSql = sSql & vbCrLf & " (select a.subj,a.seq,'" & gUserid & "',0,0,0,'" & ymd & "','99999999',0 "
    sSql = sSql & vbCrLf & "FROM vq01 a , ts02 d where a.subj=d.subj and d.userid='" & gUserid & "' and d.subj in ('','') and not exists (select a.subj,a.seq from tu02  b,ts02 c where"
    sSql = sSql & vbCrLf & "b.userid = '" & gUserid & "' and a.subj=b.subj and a.seq=b.seq and b.userid=c.userid"
    sSql = sSql & vbCrLf & "and a.subj=c.subj and a.subj=b.subj))"
    
    cnt = Fn_SQLExec(sSql).nrow
    
    Screen.MousePointer = vbNormal
    MsgBox cnt & " 건의 어카운트가 추가되었습니다.", vbExclamation, "어카운트 추가"

End Sub


Public Sub deletePaper(ByVal pocketnm As String, ByVal chasu As String)
'    Con_Open
    Dim affected As Long
    Dim lRs As ADODB.Recordset
    Dim p1 As String, p2 As String
    
    sSql = "select pocketnm,chasu from tp02 where userid='" & gUserid & "' and pcode in (select code from  tp02  where pocketnm='" & pocketnm & "' and userid='" & gUserid & "' and chasu=" & chasu & ")"
    Set lRs = Fn_SQLExec(sSql, , , True).rs
    
    Do Until lRs.EOF
    
        p1 = lRs(0).Value
        p2 = lRs(1).Value
    
       Call deletePaper(p1, p2)
       
       lRs.MoveNext
    
    Loop
    lRs.Close
    
    '=======================================
    '삭제 하는게 좋은데 남겨둔 이유는 이것과 관련해서 지우고 싶은 데이터가 어딘가에 있는게 같이지워야 하기 때문이다.
    ' 2006-01-09
    '=======================================
    sSql = "update  tp02 set hidden=1 where pocketnm='" & pocketnm & "' and userid='" & gUserid & "' and chasu=" & chasu & ""
'    sSql = "delete from tp02 where pocketnm='" & pocketnm & "' and userid='" & gUserid & "' and chasu=" & chasu & ""
    
    affected = Fn_SQLExec(sSql).nrow
    Debug.Assert affected = 1
    
    sSql = "delete from tp03 where pocketnm='" & pocketnm & "' and userid='" & gUserid & "' and chasu=" & chasu & ""
    affected = Fn_SQLExec(sSql).nrow
'    Debug.Assert affected > 0
    

'    Con_Close
End Sub



Public Function GETjindo(pn As String, chasu As Long, State As Long) As Double

Dim selPocket As String
Dim rs0 As Long
Dim rs1 As Long
Dim RS2 As Long
Dim rs3 As Long
Dim rs4 As Long
Dim selChasu As Long
Dim RSTOTAL As Long
Dim Jindo As Double

Dim rsK As Long

'vData = tvTreeView.SelectedItem.Key
'POS = InStrRev(vData, " ")

selPocket = pn
selChasu = chasu

If State = 0 Then
'    Debug.Assert False
    Con_Open
Else
    
End If
sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and  o=x AND O=0 and chasu=" & selChasu
rs0 = Fn_SQLExec(sSql).rs(0)

sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and   o>x and chasu=" & selChasu
rs1 = Fn_SQLExec(sSql).rs(0)

sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and   o<x and chasu=" & selChasu
RS2 = Fn_SQLExec(sSql).rs(0)

sSql = "SELECT COUNT(*) FROM TP03 WHERE POCKETNM='" & selPocket & "' AND USERID='" & gUserid & "' and chasu=" & selChasu
RSTOTAL = Fn_SQLExec(sSql).rs(0)

rsK = RSTOTAL - rs0 - rs1 - RS2
If RSTOTAL > 0 Then
    Jindo = (rs1 + RS2 + rsK) / RSTOTAL * 100
Else
    Jindo = 0
End If

If State = 0 Then

    Con_Close
Else
    
End If

GETjindo = Jindo
End Function


Public Function GETCORRECTRATE(pn As String, chasu As Long, State As Long) As Double

Dim selPocket As String
Dim selChasu As Long

Dim rs0 As Long
Dim rs1 As Long
Dim RS2 As Long
Dim rs3 As Long
Dim rs4 As Long
Dim RSTOTAL As Long

Dim rsK As Long


'vData = tvTreeView.SelectedItem.Key
'POS = InStrRev(vData, " ")

selPocket = pn
selChasu = chasu

If State = 0 Then
'    Debug.Assert False
    Con_Open
Else
    
End If
sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and  o=x AND O=0 and chasu=" & selChasu
rs0 = Fn_SQLExec(sSql).rs(0)

sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and   o>x and chasu=" & selChasu
rs1 = Fn_SQLExec(sSql).rs(0)

sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and   o<x and chasu=" & selChasu
RS2 = Fn_SQLExec(sSql).rs(0)

sSql = "SELECT COUNT(*) FROM TP03 WHERE POCKETNM='" & selPocket & "' AND CHASU=" & selChasu & " AND USERID='" & gUserid & "' and chasu=" & selChasu
RSTOTAL = Fn_SQLExec(sSql).rs(0)

rsK = RSTOTAL - rs0 - rs1 - RS2

If (rs1 + RS2 + rsK) > 0 Then
    GETCORRECTRATE = rs1 / (rs1 + RS2 + rsK) * 100
Else
    GETCORRECTRATE = 0
End If

If State = 0 Then
    Con_Close
Else
    
End If
'GETCORRECTRATE = jindo
End Function

Public Function GETTotalCnt(pn As String, chasu As Long, State As Long) As Long

Dim selPocket As String
Dim selChasu As Long

selPocket = pn
selChasu = chasu

If State = 0 Then
    Con_Open
Else
    
End If

sSql = "SELECT COUNT(*) FROM TP03 WHERE POCKETNM='" & selPocket & "' AND CHASU=" & selChasu & " AND USERID='" & gUserid & "' and chasu=" & selChasu
GETTotalCnt = Fn_SQLExec(sSql).rs(0)

If State = 0 Then
    Con_Close
Else
    
End If

End Function

Public Sub selall(ctl As Control)
Call ctl.SetFocus
ctl.SelStart = 0
ctl.SelLength = Len(ctl)
End Sub

Public Function ONLYNUM(KeyAscii As Integer) As Long
If IsNumeric(Chr(KeyAscii)) Then
    ONLYNUM = KeyAscii
Else
    ONLYNUM = 0
End If
End Function
Function IDEMODE() As Boolean
On Error GoTo errHandler
    Debug.Print 1 / 0
Exit Function
errHandler:
IDEMODE = True
End Function

Public Function GETFIRSTTXT(SN As String, DL As String) As String
    Dim Pos As Long
    Pos = InStrRev(SN, DL)
    If Pos = 0 Then
       GETFIRSTTXT = ""
    Else
       GETFIRSTTXT = Left$(SN, Pos - 1)
    End If

End Function

Public Function GETLASTTXT(SN As String, DL As String) As String
    Dim Pos As Long
    Pos = InStrRev(SN, DL)
    If Pos = 0 Then
       GETLASTTXT = ""
    Else
        GETLASTTXT = Mid(SN, Pos + Len(DL))
    End If

End Function


'/////////////////////////////////////////////////////////////////////////////
'//  Name   : Fn_SQLExec(DBHandle As ADODB.Connection, sSql As String)
'//  Type   : Boolean
'// return  : ut_bRecordSet형식으로 return
'// comment : 해당 핸들의 DB에 SQL문을 실행
'/////////////////////////////////////////////////////////////////////////////
Public Function Fn_SQLExec(ByVal sSql As String, Optional Caption As String = "", Optional bIgnoreDup As Boolean = False, Optional bNotCounting As Boolean = False, Optional ByRef mutex_val As Boolean) As ut_bRecordSet

Dim str6 As String
Dim nrows As Long
Dim countsql As String
Dim tRs As ADODB.Recordset

Dim strMsg As String, strLog As String

Static mutex As Boolean
Static longtimeSql As String
Static light As Boolean
Static frmMsgBox As Form

Do While mutex
    gTimer = Timer
    strMsg = "실행중인 작업이 있습니다. 잠시후 다시 시도하세요"
    
    'If "127.0.0.1" = getParamValue(Command, "h", "127.0.0.1") Then
        strLog = "시도:" + sSql + vbNewLine + "지연사유:" + longtimeSql
        If gbIsSuperAdmin Then
            strMsg = strMsg + vbNewLine + vbNewLine + strLog + "확인클릭=클립보드복사"
        End If
    'End If
    
    Call LogPut(strLog, "tmp\" + CRC16(strLog) + ".txt")
    
    
    Set frmMsgBox = New frmMsgBox60sec
    Load frmMsgBox
        
    frmMsgBox.setMsg strMsg
    frmMsgBox.cnt = 60
    TmrAfterTTS_exit = True
    frmMsgBox.Show vbModal, frmMain
    TmrAfterTTS_exit = False
    
    If gbIsSuperAdmin Then
    If frmMsgBox.btn_flag = 0 Then
        Clipboard.Clear
        Clipboard.SetText strMsg
    End If
    End If

    light = True
    
    Fn_SQLExec.Error = True
    mutex_val = mutex
    Exit Function
Loop

longtimeSql = sSql

mutex = True
mutex_val = mutex
    'Set conHandle = Con
    err.Clear

On Error GoTo ErrTrap
gMutex = True
DoEvents
gMutex = False
If Con.State = 0 Then
    'If MsgBox("DB 연결을 하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
        Con_Open
    'Else
    '    err.Raise vbObjectError + 100
    'End If
End If

  If sSql <> "select 1" Then
    Screen.MousePointer = vbHourglass
  End If
  
  gScreenHourGLS = True 'HourGlassLunch
  
  str6 = Left(sSql, 6)
  
  str6 = Trim(UCase(str6))
  
  If Left(str6, 2) = "SE" Then
'    sSql = LCase(sSql)
    Set Fn_SQLExec.rs = Nothing
    Set Fn_SQLExec.rs = New ADODB.Recordset
    Fn_SQLExec.rs.Open sSql, Con, 1, 3
'
'    Fn_SQLExec.rs.Open sSql, Con, 0, 1 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 0, 2 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 0, 3 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 0, 4 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'
'    Fn_SQLExec.rs.Open sSql, Con, 1, 1 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 1, 2 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 1, 3 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 1, 4 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'
'    Fn_SQLExec.rs.Open sSql, Con, 2, 1 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
''    Fn_SQLExec.rs.Open sSql, Con, 2, 2 ' 1, 3
''    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
''    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 2, 3 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 2, 4 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'
'    Fn_SQLExec.rs.Open sSql, Con, 3, 1 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 3, 2 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
'
''    Fn_SQLExec.rs.Open sSql, Con, 3, 3 ' 1, 3
''    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
''    Fn_SQLExec.rs.Close
'
'    Fn_SQLExec.rs.Open sSql, Con, 3, 4 ' 1, 3
'    Debug.Assert Fn_SQLExec.rs.RecordCount = -1
'    Fn_SQLExec.rs.Close
    
    
    Fn_SQLExec.Error = False
    Fn_SQLExec.nrow = 0
    
    
    If Fn_SQLExec.rs.RecordCount = -1 Then
        
      If Not bNotCounting Then
         If Not (LCase(sSql) Like "select count*") And sSql <> "select 1" Then
            countsql = "select count(*) " & Mid(sSql, InStr(LCase(sSql), "from "))
            Set tRs = Con.Execute(countsql)
            If tRs.EOF Then
                Fn_SQLExec.nrow = 0
            Else
                Fn_SQLExec.nrow = tRs(0)
            End If
 
         Else
            Fn_SQLExec.nrow = 1
         End If
      Else
'         Debug.Assert False
      End If
    Else
      Fn_SQLExec.nrow = Fn_SQLExec.rs.RecordCount
    End If
    
    
'    Fn_SQLExec.nrow = Fn_SQLExec.rs.RecordCount
    If Fn_SQLExec.nrow > 0 Then
      Fn_SQLExec.result = True
    Else
      Fn_SQLExec.result = False
    End If
    
  Else
    Call Con.Execute(sSql, nrows)
    
    Fn_SQLExec.Error = False
    Fn_SQLExec.nrow = nrows
    If nrows > 0 Then
      Fn_SQLExec.result = True
    Else
      Fn_SQLExec.result = False
    End If
  End If
    
'  Screen.MousePointer = vbDefault
  gScreenHourGLS = False
  
  SQLCA = Fn_SQLExec
mutex = False

If light Then
    
    'MsgBox "긴 쿼리 작업을 완료했습니다.", vbExclamation
    
    Set frmMsgBox = New frmMsgBox60sec
    Load frmMsgBox
        
    frmMsgBox.setMsg "긴 쿼리 작업을 완료했습니다."
    frmMsgBox.cnt = 10
    TmrAfterTTS_exit = True
    frmMsgBox.Show vbModal, frmMain
    TmrAfterTTS_exit = False
    
    light = False
End If
  gTimer = Timer
  gMutex = False
  mutex_val = mutex
  Exit Function
ErrTrap:
gMutex = True '오류팝업이 뜬 후 타이머에 의해 쿼리가 호출 될 수 있다.
  Fn_SQLExec.Error = True
  
  Fn_SQLExec.result = False
  
'  Screen.MousePointer = vbDefault
  gScreenHourGLS = False
  If err.Number = -2147467259 And bIgnoreDup Then   '듀비 또는 chekc오류
     Fn_SQLExec.nrow = 0
     gMutex = False
     mutex_val = mutex
     Exit Function
  Else
     Fn_SQLExec.nrow = -1
  End If
  Dim err_src As Variant
  Dim err_des As String
  Dim err_num As Long
  
  err_src = err.Source
  err_num = err.Number
  err_des = err.Description
  
If IDEMODE Then
    ErrMsgProc "오류(module) - Fn_SQLExec " & sSql, err_src, err_des, err_num
 Else
  ErrMsgProc "오류(module) - Fn_SQLExec ", err_src, err_des, err_num
 End If
 
 Dim retVal As Long
 
 retVal = MsgBox("오류가 발생했습니다.", vbExclamation + vbRetryCancel)
 
 If retVal = vbRetry Then
   retVal = MsgBox("오류발생 위치부터 처리하시겠습니까?", vbQuestion + vbYesNo)
   If Con.State = 0 Then
      Con_Open
   End If
   If retVal = vbYes Then
      gMutex = False
      Resume
   ElseIf retVal = vbNo Then
      gMutex = False
If IDEMODE Then
      Resume Next
Else
    End '프로그램 종료
End If
   End If
 
 ElseIf retVal = vbCancel Then
  End '프로그램 종료
 End If
 
  SQLCA = Fn_SQLExec
mutex = False
gMutex = False
End Function

Public Sub ErrMsgProc(mMsg As String, src As Variant, des As String, num As Long)
'    MainForm.tmrAni = False
    Dim TempDayString As String
    Select Case DatePart("w", Now)  ' vbUseSystemDayOfWeek
        Case vbMonday: TempDayString = "월요일"
        Case vbTuesday: TempDayString = "화요일"
        Case vbWednesday: TempDayString = "수요일"
        Case vbThursday: TempDayString = "목요일"
        Case vbFriday: TempDayString = "금요일"
        Case vbSaturday: TempDayString = "토요일"
        Case vbSunday: TempDayString = "일요일"
    End Select
    Dim TempErrorText As String
'    Debug.Assert False
    TempErrorText = "Procedure :         " & mMsg & vbCrLf & vbCrLf & _
                    "Error Source :      " & src & vbCrLf & vbCrLf & _
                    "Error Number :      " & num & vbCrLf & vbCrLf & _
                    "Error Description : " & des & vbCrLf & vbCrLf & _
                    "Date : " & Format(Now, "yyyy" & "년 " & "m" & "월 " & "d" & "일 " & TempDayString) & vbCrLf & vbCrLf & _
                    "Time : " & Time
                                
    Form_Error_Message_Box.Text_View.Text = TempErrorText
    AlwaysOnTop Form_Error_Message_Box, True
    TmrAfterTTS_exit = True
    Form_Error_Message_Box.Show vbModal
    TmrAfterTTS_exit = False
    
'    MainForm.tmrAni = True
    Set Form_Error_Message_Box = Nothing
' EX)ErrMsgProcL "객체 이름 - Private Sub Command1_Click()"
End Sub


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
Public Function selectSeries(lst As Object, Optional dli As String) As String

Dim i As Long
Dim nLen As Long
Dim str As String

If IsMissing(dli) Then
    For i = 1 To lst.ListCount - 1
        If lst.Selected(i) Then
            str = str & "'" & lst.List(i) & "',"
        End If
    Next
Else
    For i = 1 To lst.ListCount - 1
        If lst.Selected(i) Then
            str = str & "'" & GETLASTTXT(lst.List(i), dli) & "',"
        End If
    Next
End If
str = str & "''"

selectSeries = str

End Function

Public Function selectSeriesArray(lst As Variant) As String

Dim i As Long
Dim nLen As Long
Dim str As String

Dim cnt As Long
cnt = 0

For i = LBound(lst) To UBound(lst)
    str = str & "'" & Trim(lst(i)) & "'"
    If i <> UBound(lst) Then
        str = str & ","
    End If
    cnt = cnt + 1
Next

If cnt = 0 Then
    str = str & "''"
End If

selectSeriesArray = str

End Function

Public Function itemSeries1(lst As Object, Optional dli As String) As String

Dim i As Long
Dim nLen As Long
Dim str As String
Dim ListText As String


If IsMissing(dli) Then
    For i = lst.ListCount - 1 To 0 Step -1
        ListText = lst.ListText(i)
        str = str & "'" & ListText & "',"
    Next
Else
    For i = lst.ListCount - 1 To 0 Step -1
        ListText = lst.ListText(i)
        str = str & "'" & GETFIRSTTXT(ListText, dli) & "',"
    Next
End If

str = str & "''"

itemSeries1 = str

End Function


Public Function itemSeries2(lst As Object, Optional dli As String) As String

Dim i As Long
Dim nLen As Long
Dim str As String
Dim ListText As String


If IsMissing(dli) Then
    For i = lst.ListCount - 1 To 0 Step -1
        ListText = lst.ListText(i)
        str = str & "'" & ListText & "',"
    Next
Else
    For i = lst.ListCount - 1 To 0 Step -1
        ListText = lst.ListText(i)
        str = str & "'" & GETLASTTXT(ListText, dli) & "',"
    Next
End If

str = str & "''"

itemSeries2 = str

End Function


Function str2Date(ByVal str As String) As Date

Dim ymd As String

If str = "" Then
str = GETYMD
End If

'Debug.Assert Len(str) = 8
If Len(str) <> 8 Then
    str = GETYMD
End If
ymd = Mid(str, 1, 4) & "-" & Mid(str, 5, 2) & "-" & Mid(str, 7, 2)

str2Date = CDate(ymd)

'Debug.Print str2Date

End Function

Function date2Str(ByVal dt As Date) As String

date2Str = Format(dt, "yyyyMMdd")

End Function

'====================================================
'버그수정 round round(2.5)=> 2 Error Not 3
' round(213.14,-1)=> 210 But RunTime Error 5
'====================================================
Public Function Round15( _
    ByVal dblNumber As Double, _
    Optional ByVal lngDecimals As Long = 0) _
    As Double
' By Gustav, gustav@cactus.dk, 20020509
' Modification of Round14
' by Donald, donald@xbeat.net, 20020419
' modification of Round10 inspired by Jost's Round13 (Variant = CDec!)

  Static dblNumberPrevious    As Double
  Static lngDecimalsPrevious  As Long
  Static varFactor            As Variant
  Static dblFactorInv         As Double
  Static dblValue             As Double

  Dim booNewDecimals          As Boolean
  Dim booNewNumber            As Boolean

  ' Ignore excessive values of lngDecimals and dblValue.
  On Error GoTo Err_Round15

  booNewNumber = dblNumber <> dblNumberPrevious
  booNewDecimals = lngDecimals <> lngDecimalsPrevious Or dblFactorInv = 0

  If booNewDecimals = True Then
    ' Calculate factor for this number of decimals.
    varFactor = CDec(10 ^ lngDecimals)
    dblFactorInv = 10 ^ -lngDecimals
    lngDecimalsPrevious = lngDecimals
  End If

  If booNewDecimals = True Or booNewNumber = True Then
    dblNumberPrevious = dblNumber
    If dblNumber = 0 Then
      dblValue = 0
    ElseIf dblNumber > 0 Then
      dblValue = Int(dblNumber * varFactor + 0.5)
      dblValue = dblValue * dblFactorInv
    Else
      dblValue = -Int(-dblNumber * varFactor + 0.5)
      dblValue = dblValue * dblFactorInv
    End If
  End If

Exit_Round15:
  Round15 = dblValue
  'mutex_val = mutex
  Exit Function

Err_Round15:
  ' Return input value unmodified.
  dblValue = dblNumber
  'mutex_val = mutex
  Resume Exit_Round15

End Function


Public Function HSLToRGB201( _
    ByVal hue As Long, _
    ByVal Saturation As Long, _
    ByVal Luminance As Long, _
    Optional Validate As Boolean _
    ) As Long
' by Donald, donald@xbeat.net, 20011126
' after seeing Thomas Kabir's code (contact@vbfrood.de, http://www.vbfrood.de)
  Dim rR As Single, rG As Single, rB As Single
  Dim rH As Single, rL As Single, rs As Single
  Dim rMin As Single, rMax As Single, rDiff As Single
  
'  If Validate Then ValidateHSL hue, Saturation, Luminance
  
  If Saturation = 0 Then
    ' CLng(CSng(...)) else 127.5 -> 127
    HSLToRGB201 = CLng(CSng(2.55 * Luminance)) * &H10101
  Else
    rH = hue / 60: rs = Saturation / 100: rL = Luminance / 100
    If rL <= 0.5 Then
      rMin = rL * (1 - rs)
    Else
      rMin = rL - rs * (1 - rL)
    End If
    rMax = 2 * rL - rMin
    rDiff = rMax - rMin
    
    Select Case hue \ 60
    Case 0
      rR = rMax
      rB = rMin
      rG = rH * rDiff + rMin
    Case 1
      rG = rMax
      rB = rMin
      rR = rMin - (rH - 2) * rDiff
    Case 2
      rG = rMax
      rR = rMin
      rB = (rH - 2) * rDiff + rMin
    Case 3
      rB = rMax
      rR = rMin
      rG = rMin - (rH - 4) * rDiff
    Case 4
      rB = rMax
      rG = rMin
      rR = (rH - 4) * rDiff + rMin
    Case Else
      rR = rMax
      rG = rMin
      rB = rMin - (rH - 6) * rDiff
    End Select
    HSLToRGB201 = CLng(rB * 255) * &H10000 + CLng(rG * 255) * &H100 + CLng(rR * 255)
  End If
  
End Function


Public Function RGBToHSL201(ByVal RGBValue As Long) As HSL
' by Donald, donald@xbeat.net, 20011126
  Dim R As Long, g As Long, B As Long
  Dim lMax As Long, lMin As Long, lDiff As Long, lSum As Long

  R = RGBValue And &HFF&
  g = (RGBValue And &HFF00&) \ &H100&
  B = (RGBValue And &HFF0000) \ &H10000

  If R > g Then lMax = R: lMin = g Else lMax = g: lMin = R
  If B > lMax Then lMax = B Else If B < lMin Then lMin = B

  lDiff = lMax - lMin
  lSum = lMax + lMin
  
  ' Luminance
  RGBToHSL201.Luminance = lSum / 5.1!
  
  If lDiff Then
    ' Saturation
    If RGBToHSL201.Luminance <= 50& Then
      RGBToHSL201.Saturation = 100 * lDiff / lSum
    Else
      RGBToHSL201.Saturation = 100 * lDiff / (510 - lSum)
    End If
    ' Hue
    Dim q As Single: q = 60 / lDiff
    Select Case lMax
    Case R
      If g < B Then
        RGBToHSL201.hue = 360& + q * (g - B)
      Else
        RGBToHSL201.hue = q * (g - B)
      End If
    Case g
      RGBToHSL201.hue = 120& + q * (B - R)
    Case B
      RGBToHSL201.hue = 240& + q * (R - g)
    End Select
  End If
  
End Function


Public Sub SplitRGB02(ByVal lColor As Long, _
    ByRef lRed As Long, _
    ByRef lGreen As Long, _
    ByRef lBlue As Long)
' by Donald, donald@xbeat.net, 20010922
  lRed = lColor And &HFF
  lGreen = (lColor And &HFF00&) \ &H100&
  lBlue = (lColor And &HFF0000) \ &H10000
End Sub



Public Function GrayScale09(ByVal lColor As Long) As Long
' by Donald, donald@xbeat.net, 20011123
  GrayScale09 = ((77& * (lColor And &HFF&) + _
                 152& * (lColor And &HFF00&) \ &H100& + _
                  28& * ((lColor And &HFF0000) \ &H10000)) \ 256&) * &H10101
End Function


'<<
Public Function ShiftLeft06(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011001
  Select Case ShiftCount
  Case 0&
    ShiftLeft06 = Value
  Case 1&
    If Value And &H40000000 Then
      ShiftLeft06 = (Value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If Value And &H20000000 Then
      ShiftLeft06 = (Value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If Value And &H10000000 Then
      ShiftLeft06 = (Value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If Value And &H8000000 Then
      ShiftLeft06 = (Value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If Value And &H4000000 Then
      ShiftLeft06 = (Value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If Value And &H2000000 Then
      ShiftLeft06 = (Value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If Value And &H1000000 Then
      ShiftLeft06 = (Value And &HFFFFFF) * &H80& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If Value And &H800000 Then
      ShiftLeft06 = (Value And &H7FFFFF) * &H100& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If Value And &H400000 Then
      ShiftLeft06 = (Value And &H3FFFFF) * &H200& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If Value And &H200000 Then
      ShiftLeft06 = (Value And &H1FFFFF) * &H400& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If Value And &H100000 Then
      ShiftLeft06 = (Value And &HFFFFF) * &H800& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFFF) * &H800&
    End If
  Case 12&
    If Value And &H80000 Then
      ShiftLeft06 = (Value And &H7FFFF) * &H1000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If Value And &H40000 Then
      ShiftLeft06 = (Value And &H3FFFF) * &H2000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If Value And &H20000 Then
      ShiftLeft06 = (Value And &H1FFFF) * &H4000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If Value And &H10000 Then
      ShiftLeft06 = (Value And &HFFFF&) * &H8000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If Value And &H8000& Then
      ShiftLeft06 = (Value And &H7FFF&) * &H10000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFF&) * &H10000
    End If
  Case 17&
    If Value And &H4000& Then
      ShiftLeft06 = (Value And &H3FFF&) * &H20000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFF&) * &H20000
    End If
  Case 18&
    If Value And &H2000& Then
      ShiftLeft06 = (Value And &H1FFF&) * &H40000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFF&) * &H40000
    End If
  Case 19&
    If Value And &H1000& Then
      ShiftLeft06 = (Value And &HFFF&) * &H80000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFF&) * &H80000
    End If
  Case 20&
    If Value And &H800& Then
      ShiftLeft06 = (Value And &H7FF&) * &H100000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FF&) * &H100000
    End If
  Case 21&
    If Value And &H400& Then
      ShiftLeft06 = (Value And &H3FF&) * &H200000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FF&) * &H200000
    End If
  Case 22&
    If Value And &H200& Then
      ShiftLeft06 = (Value And &H1FF&) * &H400000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FF&) * &H400000
    End If
  Case 23&
    If Value And &H100& Then
      ShiftLeft06 = (Value And &HFF&) * &H800000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFF&) * &H800000
    End If
  Case 24&
    If Value And &H80& Then
      ShiftLeft06 = (Value And &H7F&) * &H1000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7F&) * &H1000000
    End If
  Case 25&
    If Value And &H40& Then
      ShiftLeft06 = (Value And &H3F&) * &H2000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3F&) * &H2000000
    End If
  Case 26&
    If Value And &H20& Then
      ShiftLeft06 = (Value And &H1F&) * &H4000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1F&) * &H4000000
    End If
  Case 27&
    If Value And &H10& Then
      ShiftLeft06 = (Value And &HF&) * &H8000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &HF&) * &H8000000
    End If
  Case 28&
    If Value And &H8& Then
      ShiftLeft06 = (Value And &H7&) * &H10000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7&) * &H10000000
    End If
  Case 29&
    If Value And &H4& Then
      ShiftLeft06 = (Value And &H3&) * &H20000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3&) * &H20000000
    End If
  Case 30&
    If Value And &H2& Then
      ShiftLeft06 = (Value And &H1&) * &H40000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1&) * &H40000000
    End If
  Case 31&
    If Value And &H1& Then
      ShiftLeft06 = &H80000000
    Else
      ShiftLeft06 = &H0&
    End If
  End Select
End Function


'>>
Public Function ShiftRight08(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011010
  Select Case ShiftCount
  Case 0&:  ShiftRight08 = Value
  Case 1&:  ShiftRight08 = (Value And &HFFFFFFFE) \ &H2&
  Case 2&:  ShiftRight08 = (Value And &HFFFFFFFC) \ &H4&
  Case 3&:  ShiftRight08 = (Value And &HFFFFFFF8) \ &H8&
  Case 4&:  ShiftRight08 = (Value And &HFFFFFFF0) \ &H10&
  Case 5&:  ShiftRight08 = (Value And &HFFFFFFE0) \ &H20&
  Case 6&:  ShiftRight08 = (Value And &HFFFFFFC0) \ &H40&
  Case 7&:  ShiftRight08 = (Value And &HFFFFFF80) \ &H80&
  Case 8&:  ShiftRight08 = (Value And &HFFFFFF00) \ &H100&
  Case 9&:  ShiftRight08 = (Value And &HFFFFFE00) \ &H200&
  Case 10&: ShiftRight08 = (Value And &HFFFFFC00) \ &H400&
  Case 11&: ShiftRight08 = (Value And &HFFFFF800) \ &H800&
  Case 12&: ShiftRight08 = (Value And &HFFFFF000) \ &H1000&
  Case 13&: ShiftRight08 = (Value And &HFFFFE000) \ &H2000&
  Case 14&: ShiftRight08 = (Value And &HFFFFC000) \ &H4000&
  Case 15&: ShiftRight08 = (Value And &HFFFF8000) \ &H8000&
  Case 16&: ShiftRight08 = (Value And &HFFFF0000) \ &H10000
  Case 17&: ShiftRight08 = (Value And &HFFFE0000) \ &H20000
  Case 18&: ShiftRight08 = (Value And &HFFFC0000) \ &H40000
  Case 19&: ShiftRight08 = (Value And &HFFF80000) \ &H80000
  Case 20&: ShiftRight08 = (Value And &HFFF00000) \ &H100000
  Case 21&: ShiftRight08 = (Value And &HFFE00000) \ &H200000
  Case 22&: ShiftRight08 = (Value And &HFFC00000) \ &H400000
  Case 23&: ShiftRight08 = (Value And &HFF800000) \ &H800000
  Case 24&: ShiftRight08 = (Value And &HFF000000) \ &H1000000
  Case 25&: ShiftRight08 = (Value And &HFE000000) \ &H2000000
  Case 26&: ShiftRight08 = (Value And &HFC000000) \ &H4000000
  Case 27&: ShiftRight08 = (Value And &HF8000000) \ &H8000000
  Case 28&: ShiftRight08 = (Value And &HF0000000) \ &H10000000
  Case 29&: ShiftRight08 = (Value And &HE0000000) \ &H20000000
  Case 30&: ShiftRight08 = (Value And &HC0000000) \ &H40000000
  Case 31&: ShiftRight08 = CBool(Value And &H80000000)
  End Select
End Function


Public Function getLastDay(ByVal pYear As Long, ByVal pMonth As Long) As Long

Dim dt As Date

If pMonth = 12 Then
    pMonth = 1
    pYear = pYear + 1
Else
    pMonth = pMonth + 1
End If

getLastDay = CInt(Format(CDate(pYear & "-" & pMonth & "-01") - 1, "dd"))


End Function

Public Function MakeDbTxt(str1 As String) As String
MakeDbTxt = Replace(str1, "'", "''")

End Function


'==============================================================================
'min ~ max 사이의 랜덤한 값을 가져온다.
'==============================================================================
Public Function rndVal(Min As Double, Max As Double) As Double
    Randomize
    rndVal = (Max - Min) * Rnd + Min
End Function

'==============================================================================
'현재시간의 1000초의 값을 가지고 guid를 생성한다.(주의/ 빨리 수행되면 guid가 아니다.
'==============================================================================
Public Function getGUID(n진수) As String

Dim val As Double
Dim nVal As String
Dim stack1 As String
Dim Order As Long

Dim vArray  As Variant
Dim 몫 As Long
Dim 나머지 As Double
Dim stack2 As String

val = CDbl(Now) * 10000000000#

Order = 0

Do While (n진수 ^ Order) < val
    Order = Order + 1
Loop


Do
    몫 = Fix(Trim(val / (n진수 ^ Order)))
    나머지 = val - 몫 * (n진수 ^ Order)
    Order = Order - 1
    stack1 = stack1 & 몫 & "_"
    val = 나머지
Loop While Order >= 0


vArray = Split(stack1, "_")

stack2 = ""

Dim i As Long

Do While vArray(i) <> ""
    stack2 = stack2 & num2antilog_char(vArray(i))
    i = i + 1
Loop
getGUID = stack2
End Function

'==============================================================================
'guid 함수에서 콜한다.
'==============================================================================
Function num2antilog_char(ByVal C As Long) As String
    Dim val As String
    Select Case C
        Case 0: val = "0"
        Case 1: val = "1"
        Case 2: val = "2"
        Case 3: val = "3"
        Case 4: val = "4"
        Case 5: val = "5"
        Case 6: val = "6"
        Case 7: val = "7"
        Case 8: val = "8"
        Case 9: val = "9"
        Case 10: val = "A"
        Case 11: val = "B"
        Case 12: val = "C"
        Case 13: val = "D"
        Case 14: val = "E"
        Case 15: val = "F"
        Case 16: val = "G"
        Case 17: val = "H"
        Case 18: val = "I"
        Case 19: val = "J"
        Case 20: val = "K"
        Case 21: val = "L"
        Case 22: val = "M"
        Case 23: val = "N"
        Case 24: val = "O"
        Case 25: val = "P"
        Case 26: val = "Q"
        Case 27: val = "R"
        Case 28: val = "S"
        Case 29: val = "T"
        Case 30: val = "U"
        Case 31: val = "V"
        Case 32: val = "W"
        Case 33: val = "X"
        Case 34: val = "Y"
        Case 35: val = "Z"
        Case 36: val = "a"
        Case 37: val = "b"
        Case 38: val = "c"
        Case 39: val = "d"
        Case 40: val = "e"
        Case 41: val = "f"
        Case 42: val = "g"
        Case 43: val = "h"
        Case 44: val = "i"
        Case 45: val = "j"
        Case 46: val = "k"
        Case 47: val = "l"
        Case 48: val = "m"
        Case 49: val = "n"
        Case 50: val = "o"
        Case 51: val = "p"
        Case 52: val = "q"
        Case 53: val = "r"
        Case 54: val = "s"
        Case 55: val = "t"
        Case 56: val = "u"
        Case 57: val = "v"
        Case 58: val = "w"
        Case 59: val = "x"
        Case 60: val = "y"
        Case 61: val = "z"
        Case Else
            Debug.Assert False
    End Select
num2antilog_char = val
End Function

'==============================================================================
' 완전한 이름의 파일경로와 파일명을 리턴한다.
'==============================================================================
Public Function getPerfectFile(pth As String, fnm As String) As String
    If Right(pth, 1) = "\" Then
        getPerfectFile = pth & fnm
    Else
        getPerfectFile = pth & "\" & fnm
    End If
End Function

Public Function LenH(ByVal strString As String) As Long
    '용도: 한글과 영문이 혼합된 경우 한글을 2바이트, 영문은 1바이트로 계산
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

Public Function StripNulls(OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = Trim(OriginalStr)
End Function

'-----------------------------------------------------------
' SUB: UpdateStatus
'
' "Fill" (by percentage) inside the PictureBox and also
' display the percentage filled
'
' IN: [pic] - PictureBox used to bound "fill" region
'     [sngPercent] - Percentage of the shape to fill
'     [fBorderCase] - Indicates whether the percentage
'        specified is a "border case", i.e. exactly 0%
'        or exactly 100%.  Unless fBorderCase is True,
'        the values 0% and 100% will be assumed to be
'        "close" to these values, and 1% and 99% will
'        be used instead.
'
' Notes: Set AutoRedraw property of the PictureBox to True
'        so that the status bar and percentage can be auto-
'        matically repainted if necessary
'-----------------------------------------------------------
'
Public Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase As Boolean = False)
    Dim strPercent As String
    Dim intX As Long
    Dim intY As Long
    Dim intWidth As Long
    Dim intHeight As Long

    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &H800000 ' dark blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    Dim intPercent
    intPercent = Int(100 * sngPercent + 0.5)
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If

    pic.Refresh
End Sub

Public Sub CenterForm(anyForm As Form)
   '폼을 중앙에 위치
    anyForm.Top = (Screen.Height - anyForm.Height) / 2
    anyForm.Left = (Screen.Width - anyForm.Width) / 2
End Sub

Function autoCRLF(ByRef str As String, ByVal W As Single, ByRef pic As PictureBox, Optional ByVal bInit As Boolean = False) As String
Dim strNew As String
Dim strRemain As String
Dim strLineFront As String
Dim strLineTemp As String
Dim tempW As Single


Static firstWidth As Single
Static K As Long
If bInit Then
    K = 0
End If
K = K + 1
If K = 1 Then
    firstWidth = W
End If

Dim Pos As Long
Dim posVbcrlf As Long

If str = "" Then Exit Function

strRemain = str

Pos = InStr(strRemain, " ")
posVbcrlf = InStr(strRemain, vbCrLf)

If posVbcrlf <> 0 Then
    If posVbcrlf < Pos And posVbcrlf = 1 Then
        strRemain = Mid(strRemain, posVbcrlf + 2)
        Pos = posVbcrlf + 2
        autoCRLF = vbNewLine & autoCRLF(strRemain, firstWidth, pic)
        Exit Function
    ElseIf posVbcrlf < Pos Then 'at 20051114 Modifyed By Park Gyu Sun
'        Debug.Assert False
        strLineFront = Left(strRemain, posVbcrlf - 1)
        strRemain = Mid(strRemain, posVbcrlf + 2)
        Pos = posVbcrlf + 2
        autoCRLF = strLineFront & vbNewLine & autoCRLF(strRemain, firstWidth, pic)
        Exit Function
        
    End If
End If

strLineFront = ""

If Pos = 0 Then
    Pos = Len(strRemain)
End If
    
'=---------------------------------------------------------------------------
    strLineTemp = strLineFront & Left(strRemain, Pos)
    
    If (pic.TextWidth(strLineTemp) > W) Then
        If (W <> firstWidth) Then
            autoCRLF = vbNewLine & autoCRLF(strRemain, firstWidth, pic)
            Exit Function
        End If
        strLineFront = Left(strRemain, 1)
        strRemain = Mid(strRemain, 2)
    
        Do While pic.TextWidth(strLineFront) < W
            
            strLineTemp = strLineFront & Left(strRemain, 1)
            If pic.TextWidth(strLineTemp) > W Then
                Exit Do
            End If
            
            If strRemain = "" Then Exit Do
            
            strLineFront = strLineTemp
            strRemain = Mid(strRemain, 2)
            
        Loop
        
        strNew = strLineFront & vbNewLine & autoCRLF(strRemain, firstWidth, pic)
    Else
        strLineFront = strLineTemp
        strRemain = Mid(strRemain, Len(strLineFront) + 1)
        
        
        If strLineFront = "" Then
            strLineFront = strLineFront & Left(strRemain, 1)
            strRemain = Mid(strRemain, 2)
        
            Do While pic.TextWidth(strLineFront) < W
                
                strLineTemp = strLineFront & Left(strRemain, 1)
                If pic.TextWidth(strLineTemp) > W Then
                    Exit Do
                End If
                
                If strRemain = "" Then Exit Do
                
                strLineFront = strLineTemp
                strRemain = Mid(strRemain, 2)
                
            Loop
            
            strNew = strLineFront & vbNewLine & autoCRLF(strRemain, firstWidth, pic)
        Else
            
            tempW = W - pic.TextWidth(strLineFront)
            
            If tempW > pic.TextWidth("가") Then
                strNew = strLineFront & autoCRLF(strRemain, tempW, pic)
            Else
                If Left(strRemain, 1) = vbCr Then
                    strNew = strLineFront & autoCRLF(strRemain, firstWidth, pic) 'newline 2개를 방지한다.
                Else
                    strNew = strLineFront & vbNewLine & autoCRLF(strRemain, firstWidth, pic)
                End If
            End If
        End If
    End If
'=---------------------------------------------------------------------------
autoCRLF = strNew

End Function



'====================================================================
'망각 색상을 계산한다.
'--------------------------------------------------------------------
' 입력범위: 0~ 255
'
' 출력값: 색상 0에 대응하는 값은 검정색 ,중간색은 노란색, 255에 해당하는 색은 녹색
'
'====================================================================
Function calForgetColor(ByVal val As Integer) As Double
Dim R As Integer, g As Integer
Select Case val
Case 0 To 127
    g = 255 / (255 / 2) * val
    R = g
Case 128 To 255:
    g = 255
    R = -255 / (255 - (255 / 2)) * (val - 255)
Case Else
    Call err.Raise(vbObjectError + 100, "Module1.calGorgetColor", "입력범위오류")
End Select

calForgetColor = rgb(R, g, 0)

End Function

Function calForgetcolor2(ByVal oCnt As Integer, ByVal XCnt As Integer, ByVal update_ymd As String, ByVal reserve_ymd As String, ByVal dayFactor As Integer) As Double
Dim dt1 As Date
Dim dt2 As Date
Dim dtToday As Date

Dim delta1 As Long
Dim delta2 As Long
Dim deltaAll As Long

Dim ratio As Double

If reserve_ymd = "99999999" Then
    calForgetcolor2 = vbBlack
    Exit Function
End If

dt1 = str2Date(update_ymd)
dt2 = str2Date(reserve_ymd)
'dt1 = dt2 - dayFactor

dtToday = str2Date(GETYMD())

If dt2 < dtToday Then
    calForgetcolor2 = vbBlack
    Exit Function
End If


deltaAll = DateDiff("d", dt1, dt2) + 1
delta1 = DateDiff("d", dt1, dtToday)
delta2 = DateDiff("d", dtToday, dt2) + 1

ratio = delta2 / deltaAll * 255 '255분율로 했을 때의 기억정도

If (oCnt + XCnt > 0) Then
    ratio = ratio * (oCnt) / (oCnt + XCnt)  '기억정도에 신뢰도를 곱합
End If

calForgetcolor2 = calForgetColor(CInt(ratio))


End Function

'========================================================
'진동학의 중감쇠 적용
'--------------------------------------------------------
'
'x=A*exp(a*t)
't-x선도에서 t=0 일때 x=A 가되며 t가 무한대로 갈때 x는 0으로 수렴한다.
'a는 0보다 작은 수이며
' -0.005 일 경우를 일반상수로 하였으나 나중에는 사람에 따른 변수로 사용할 것임
' 만약 a가 -0.002 일경우는 약간 완만한 감쇠로 진행이 되는것이고
' a가 0에 가까울수록 느슨하게 감쇠함
' 처음에 -0.002로 하였으나 강하게 감쇠하게 적용하기 위해 -0.005로 하였음
' 아래는 엑셀의 결과이다.

'x=A*e^(a*t)
'a=-1
'a -0.005
'a 1
'
't X
'0.000000    1.000000
'1.000000    0.995012
'2.000000    0.990050
'3.000000    0.985112
'4.000000    0.980199
'5.000000    0.975310
'6.000000    0.970446
'...
'365.000000  0.161218
'...
'710.000000  0.028725
'...
'1000.000000     0.006738
'
'========================================================
Function heavydamping(a As Double, daydiff) As Double
    Dim rval As Double
    rval = Exp(a * Abs(daydiff))
    Debug.Assert rval <= 1 And rval > 0
    heavydamping = rval
End Function

Function iString(cnt As Integer, opt As Integer) As String
    Dim c1 As Integer, C2 As Integer
    Dim retVal As String
    
    c1 = Fix(cnt / 10)
    C2 = cnt - c1 * 10
    
    If opt = 1 Then
        retVal = String(c1, "★")
        retVal = retVal + Replace(String(C2, "■"), "■■■■■■", "■■■■■,■")
    Else
        retVal = String(c1, "☆")
        retVal = retVal + Replace(String(C2, "□"), "□□□□□□", "□□□□□,□")
    End If
    
    iString = retVal
    
End Function

Function getParamValue(strCmd As String, strParam As String, strDefault As String) As String
    Dim tmpArr As Variant
    
    Dim strCmdTmp As String
    
    strCmdTmp = Replace(strCmd, "/", "-")
        
    strCmdTmp = Replace(strCmdTmp, "--host", "-h ") 'Sever
    strCmdTmp = Replace(strCmdTmp, "--password", "-p ") 'Password
    strCmdTmp = Replace(strCmdTmp, "--uid", "-u ") 'Uid
    strCmdTmp = Replace(strCmdTmp, "--option", "-o ") 'Option
    strCmdTmp = Replace(strCmdTmp, "--port", "-P ") 'Port
    strCmdTmp = Replace(strCmdTmp, "--database", "-d ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "--ftpaddr", "-x ") 'ftp addr
    strCmdTmp = Replace(strCmdTmp, "--ftpuser", "-y ") 'ftp user
    strCmdTmp = Replace(strCmdTmp, "--ftppass", "-z ") 'ftp pass
    strCmdTmp = Replace(strCmdTmp, "--debug", "-D ") 'debug mode

    strCmdTmp = Replace(strCmdTmp, "-h", "-h ") 'Sever
    strCmdTmp = Replace(strCmdTmp, "-p", "-p ") 'Password
    strCmdTmp = Replace(strCmdTmp, "-u", "-u ") 'Uid
    strCmdTmp = Replace(strCmdTmp, "-o", "-o ") 'Option
    strCmdTmp = Replace(strCmdTmp, "-P", "-P ") 'Port
    strCmdTmp = Replace(strCmdTmp, "-d", "-d ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "-x", "-x ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "-y", "-y ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "-z", "-z ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "-D", "-D ") 'Debug mode
    
    strCmdTmp = Replace(strCmdTmp, "  ", " ") 'space handling
    strCmdTmp = Replace(strCmdTmp, "  ", " ") 'space handling
    
    tmpArr = Split(strCmdTmp, " ")
    getParamValue = strDefault
    
    Dim strParamTmp As String
    
    strParamTmp = "-" & strParam
    
    Dim i As Long
    For i = 0 To UBound(tmpArr)
        If tmpArr(i) = strParamTmp Then
            getParamValue = tmpArr(i + 1)
            Exit For
        End If
        
    Next
    
End Function

'==========================================
'폰트 목록 얻기 process
'사용예
'Call FillListWithFonts(List1)
'==========================================
Public Function EnumFontProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
     ByVal FontType As Long, lParam As ComboBox) As Long
  Dim FaceName As String
  Dim FullName As String

    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    If InStr(FaceName, "@") = 0 Then
    lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    End If
    EnumFontProc = 1

End Function

'==========================================
'폰트 목록 얻기 process
'사용예
'Call FillListWithFonts(List1)
'==========================================
Public Sub FillListWithFonts(lb As ComboBox)
  Dim hdc As Long

    lb.Clear
    hdc = GetDC(lb.hwnd)
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontProc, lb
    ReleaseDC lb.hwnd, hdc

End Sub

Public Sub make_html(URL As String, file_name As String)
    
    Dim str As String
    str = "<script>window.open('MZM','_iqup5000_','toolbar=no,menubar=no,scrollbars=yes,statusbar=yes,resizable=yes');</script>"
    
    str = Replace(str, "MZM", URL)
    
    Open file_name For Output As #1
    Print #1, str
    Close #1
    
End Sub

Public Sub make_html_pop(word As String, file_name As String)

    Dim str As String
    str = Replace(dic_str, "MZM", word)
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub

Public Sub make_html_pop_hanja(word As String, file_name As String)

    Dim str As String
    str = Replace(dic_hanja_str, "MZM", word)
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub

Public Sub make_html_pop_modum(word As String, file_name As String)

    Dim str As String
    str = Replace(dic_modum_str, "MZM", word)
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub

Public Sub make_html_pop_kr(word As String, file_name As String)

    Dim str As String
    str = Replace(dic_kr_str, "MZM", word)
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub

Public Sub make_html_pop_ee(word As String, file_name As String)

    Dim str As String
    str = Replace(dic_ee_str, "MZM", word)
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub

Public Sub make_html_pop_jp(word As String, file_name As String)

    Dim str As String
    str = Replace(dic_jp_str, "MZM", word)
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub

Public Sub make_html_pop_cn(word As String, file_name As String)

    Dim str As String
    str = Replace(dic_cn_str, "MZM", word)
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub

Public Sub LogPut(str As String, file_name As String)
    
    Open file_name For Output As #1
    Print #1, str
    Close #1

End Sub


Public Sub make_html_tts_ko(word As String, file_name As String)

    Dim str As String
    str = Replace(tts_ko_str, "MZM", word)
    Call Module1.toclipboard(word)
    Open file_name For Output As #1
    Print #1, str
    Close #1
    
End Sub


Public Sub make_html_tts_eng(word As String, file_name As String)

    Dim str As String
    str = Replace(tts_eng_str, "MZM", Encode(word))
    Open file_name For Output As #1
    Print #1, str
    Close #1
    
End Sub
'CREATE TABLE `vq01_log` (
'  `idx` int(10) NOT NULL AUTO_INCREMENT,
'  `subj` varchar(255) DEFAULT NULL,
'  `seq` int(11) DEFAULT NULL,
'  `cat` varchar(255) DEFAULT NULL,
'  `quiz` mediumtext,
'  `a` mediumtext,
'  `b` mediumtext,
'  `c` mediumtext,
'  `d` mediumtext,
'  `e` mediumtext,
'  `hint` mediumtext,
'  `ans` varchar(255) DEFAULT NULL,
'  `resid` varchar(255) DEFAULT NULL,
'  `mode` varchar(255) DEFAULT NULL,
'  `updateymd` varchar(255) DEFAULT NULL,
'  `updatechasu` int(11) DEFAULT NULL,
'  UNIQUE KEY `idx` (`idx`),
'  KEY `subj_2` (`subj`,`seq`))

'==================================
'배열의 크기를 구한다 1개
'==================================
Public Function getLength(ByRef arr As Variant) As Long
    On Error GoTo ErrTrap
    getLength = UBound(arr) - LBound(arr) + 1
    Exit Function
ErrTrap:
    getLength = 0
End Function

Public Sub toUTF8(ByRef txtFullPath As String)
    churiFile txtFullPath
End Sub

Public Sub churiFile(ByVal fname As String)
    
    Dim intFileNum As Integer
    Dim instBOM As abyteBOM
    Dim isUTF8 As Boolean
    Dim strK As String
    Dim strTmp As String
    
    Dim abytUTF16() As Byte
    Dim abytUTF16tmp() As Byte
    Dim abytUTF8() As Byte
    Dim bData() As Byte
    Dim isSuccess As Boolean
    Dim lLen As Long
    Dim isUNICODE As Boolean
    Dim retVal As Integer
    Dim i As Long
    
    Dim colList As New Collection
    
    isSuccess = False
    If True Then
     intFileNum = FreeFile()
     ' UTF-8 방식의 파일 이름과 경로를 적절히 ...
     Open fname For Random Access Read As #intFileNum Len = LenB(instBOM)
        instBOM.byte1 = 0
        instBOM.Byte2 = 0
        instBOM.byte3 = 0
        instBOM.byte4 = 0
        
        isUTF8 = False
        lLen = FileLen(fname)
        
        Get #intFileNum, , instBOM
        If lLen >= 2 Then
            isUNICODE = (255 = instBOM.byte1 And 254 = instBOM.Byte2)
            If isUNICODE = False Then
                isUNICODE = (254 = instBOM.byte1 And 255 = instBOM.Byte2) Or (0 = instBOM.byte1 And 0 = instBOM.Byte2 And 254 = instBOM.byte3 And 255 = instBOM.byte4)
            End If
            If isUNICODE = False And lLen >= 3 Then
                isUTF8 = (239 = instBOM.byte1 And 187 = instBOM.Byte2 And 191 = instBOM.byte3) '0xEFBBBF (239,187,191)
            End If
        End If

     Close #intFileNum
     
     Dim indata As String
     
     intFileNum = FreeFile()
     
     'If optMode(1).Value Then
     '   If isUTF8 Then
     '       MsgBox "해당파일은 UTF8 포맷입니다.", vbExclamation, "안내"
     '   ElseIf isUNICODE Then
     '       MsgBox "해당파일은 UNICODE 포맷입니다.", vbExclamation, "안내"
     '   End If
     'End If
     
     If isUTF8 = False And isUNICODE = False Then
         intFileNum = FreeFile()

         Open fname For Input As #intFileNum
         strK = ""
         
         If FileLen(fname) > 0 Then
         
             If Not EOF(intFileNum) Then
                Line Input #intFileNum, strK
                DoEvents
                err.Clear
                On Error GoTo ErrTrap
                
                
                'lstFileContent.Clear
                
                Do Until EOF(intFileNum)
                   'If EOF(intFileNum) = False Then
                    Line Input #intFileNum, strTmp '이렇게 데이터를 읽어오면 파일은 ANSI로보고 문자열변수를 UTF16문자셋으로 변환되어 읽게 된다.
                    colList.Add strTmp
                Loop
                
                GoTo noError
ErrTrap:
                If err.Number = 62 Then
                    '무시하고 다음으로 넘어감
                    Debug.Assert False
                Else
                    Debug.Assert False
                    retVal = MsgBox(fname + "오류가 발생하였습니다.중지하시겠습니까?", vbCritical + vbAbortRetryIgnore, "오류")
                    If retVal = vbRetry Then
                        Resume
                    ElseIf retVal = vbIgnore Then
                        Resume Next
                    Else 'abort
                        Exit Sub
                    End If
                End If
noError:
             End If
         Else
            '파일크기가 0
'            lstErrLog.AddItem "▲경고(파일길이0):" + fname
            
         End If
         Close #intFileNum
         
         abytUTF16tmp = strK '문자열변수를 byte배열로 출력하면 UTF16의 바이트 Array로 리턴된다.
     
        abytUTF8 = UTF8FromUTF16(abytUTF16tmp, isSuccess)
        If isSuccess Or (isSuccess = False And strK = "") Then
            'If chkBackup.Value = vbChecked Then
            '    FileCopy fname, fname + ".ansi.bak"
            'End If
            
            DeleteFile fname
            
            Open fname For Binary As #1
            'If optMode(0).Value Or chkBOMUse.Value = Checked Then
            '    Put #1, , abytUTF8BOM
            'End If
            If True Then
                Put #1, , abytUTF8
                For i = 1 To colList.count
                    strK = vbNewLine & colList.Item(i)
                    abytUTF16tmp = strK
                    abytUTF8 = UTF8FromUTF16(abytUTF16tmp, isSuccess)
                    Put #1, , abytUTF8
                Next
            End If
            Close #1
            'Call lstStdLog.AddItem(fname)
            'txtStdLog.Text = txtStdLog.Text + fname + vbNewLine
        Else
            'txtErrLog.Text = txtErrLog.Text + "●오류:" + fname + vbNewLine
            'Call lstErrLog.AddItem("●오류:" + fname)
            MsgBox fname + vbNewLine + "처리중 오류발생", vbCritical, "알림"
        End If
         
        If isSuccess Then
           'churiCnt = churiCnt + 1
        End If
     End If
   End If
    
   Set colList = Nothing
    
End Sub

' 두번째 방법 : BOM 을 무시한 후, UTF-8 방식으로 변환한 후,
' UTF-8 방식을 나타내는 Signature 를 추가하여 반환
'
Public Function UTF8FromUTF16withMark(ByRef abytUTF16() As Byte, ByRef isSuccess As Boolean) As Byte()
    Dim abytTemp() As Byte
    Dim abytUTF8() As Byte
    Dim lngByteNum As Long
    Dim lngCharCount As Long
    Dim lngUpper As Long
    
    On Error GoTo ConversionErr
    
    abytTemp = abytUTF16
    lngUpper = UBound(abytTemp)
    If lngUpper > 1 Then
    ' UTF-16 LE 의 바이트순서표식이 있을 경우 이를 일단 삭제합니다.
    ' &HFEFF 문자인데, LE에서는 도치되어 저장되므로, &HFF 가 먼저 위치함.
    If abytTemp(0) = &HFF And abytTemp(1) = &HFE Then
    Call CopyMemory(abytTemp(0), abytTemp(2), lngUpper - 1)
    ReDim Preserve abytTemp(lngUpper - 2)
    lngUpper = lngUpper - 2
    End If
    End If
    lngCharCount = (lngUpper + 1) \ 2
    
    ' 이제 변환에 필요한 메모리의 크기를 구합니다.
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, 0, 0, 0, 0)
    
    If lngByteNum > 0 Then
    ReDim abytUTF8(lngByteNum - 1)
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, _
    abytUTF8(0), lngByteNum, 0, 0)
    lngUpper = UBound(abytUTF8)
    ' 변환되어 있는 UTF-8 바이트 배열 선두에 UTF-8 표식을 넣기 위해
    ' 기존의 바이트 배열을 뒤로 밀어내고, 배열 앞부분에 표식을 추가합니다.
    ReDim Preserve abytUTF8(lngUpper + 3)
    Call CopyMemory(abytUTF8(3), abytUTF8(0), lngUpper + 1)
    abytUTF8(0) = &HEF
    abytUTF8(1) = &HBB
    abytUTF8(2) = &HBF
    
    UTF8FromUTF16withMark = abytUTF8
    End If
    isSuccess = True
    Exit Function
    
ConversionErr:
    isSuccess = False
    MsgBox " Conversion failed "
End Function

' 첫번째 방법 : UTF-16을 나타내는 Byte Order Mark(BOM) 가 없을 경우,
'
Public Function UTF8FromUTF16(ByRef abytUTF16() As Byte, ByRef isSuccess As Boolean) As Byte()
    Dim lngByteNum As Long
    Dim abytUTF8() As Byte
    Dim lngCharCount As Long
    
    On Error GoTo ConversionErr
    
    lngCharCount = (UBound(abytUTF16) + 1) \ 2
    ' UTF-16 LE 스트링의 문자의 수를 대입시켜, 변환에 필요한 바이트 수를 구합니다.
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, 0, 0, 0, 0)
    
    If lngByteNum > 0 Then
    ' 변환된 코드를 반환받을 메모리를 확보한 후 함수를 호출합니다.
    ReDim abytUTF8(lngByteNum - 1)
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, _
    abytUTF8(0), lngByteNum, 0, 0)
    UTF8FromUTF16 = abytUTF8
    End If
    isSuccess = True
    Exit Function
    
ConversionErr:
    isSuccess = False
    'MsgBox " Conversion failed "
End Function


Public Function RedFromRGB(ByVal rgb As Long) _
   As Integer
   ' The ampersand after &HFF coerces the number as a
   ' long, preventing Visual Basic from evaluating the
   ' number as a negative value. The logical And is
   ' used to return bit values.
   RedFromRGB = &HFF& And rgb
End Function

Public Function GreenFromRGB(ByVal rgb As Long) _
   As Integer
   ' The result of the And operation is divided by
   ' 256, to return the value of the middle bytes.
   ' Note the use of the Integer divisor.
   GreenFromRGB = (&HFF00& And rgb) \ 256
End Function

Public Function BlueFromRGB(ByVal rgb As Long) _
   As Integer
   ' This function works like the GreenFromRGB above,
   ' except you don't need the ampersand. The
   ' number is already a long. The result divided by
   ' 65536 to obtain the highest bytes.
   BlueFromRGB = (&HFF0000 And rgb) \ 65536
End Function

Function Decode(str As String) As String
Dim C As String
Dim i As Long, j As Long
Dim arr() As Byte
Dim L As Long
If str = "" Then
    Decode = ""
    Exit Function
End If

L = Len(str) - 1
For i = 0 To L
    ReDim Preserve arr(0 To j)
    C = Mid(str, i + 1, 1)
    If (C = "%") Then
        i = i + 1
        C = Mid(str, i + 1, 2)
        i = i + 1
        arr(j) = CByte("&H" & C)
    Else
        On Error Resume Next
        arr(j) = CByte(Asc(C))
        
        If err.Number = 0 Then
        ElseIf err.Number = 6 Then
            
            MsgBox "디코드 할 수 있는 문자가 잘못되었습니다.[" + C + "]문자", vbExclamation
            err.Clear
            Exit Function
        Else
           
            MsgBox err.Description, vbCritical
            err.Clear
        End If
        On Error GoTo 0
    End If
    j = j + 1
Next

str = arr
str = StrConv(str, vbUnicode)
Decode = str

End Function


Function Encode(str As String) As String
Dim C As String
Dim i As Long, j As Long
Dim arr() As Byte
Dim L As Long
If str = "" Then
    Encode = ""
    Exit Function
End If

L = Len(str) - 1
For i = 0 To L
    C = Hex(Asc(Mid(str, i + 1, 1)))
    
    If Len(C) = 1 Then
        Encode = Encode + "%0" + C
    ElseIf Len(C) = 2 Then
        Encode = Encode + "%" + C
    ElseIf Len(C) = 3 Then
        Encode = Encode + "%0" + Mid(C, 1, 1) + "%" + Mid(C, 2, 2)
    ElseIf Len(C) = 4 Then
        Encode = Encode + "%" + Mid(C, 1, 2) + "%" + Mid(C, 3, 2)
    Else
        MsgBox ("test?")
    End If
    
Next

Encode = LCase(Encode)

End Function


'--------------------------------
' AToW
'
' ANSI to UNICODE conversion, via a given codepage.
'--------------------------------
Public Function AToW(ByVal st As String, Optional ByVal cpg As Long = -1, Optional ByVal lFlags As Long = 0) As String
  Dim stBuffer As String
  Dim cwch As Long
  Dim pwz As Long
  Dim pwzBuffer As Long
  If cpg = -1 Then cpg = GetACP()
  pwz = StrPtr(st)
  cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, 0&, 0&)
  stBuffer = String$(cwch + 1, vbNullChar)
  pwzBuffer = StrPtr(stBuffer)
  cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer))
  AToW = Left$(stBuffer, cwch - 1)
End Function
'--------------------------------
' WToA
'
' UNICODE to ANSI conversion, via a given codepage
'--------------------------------
Public Function WToA(ByVal st As String, Optional ByVal cpg As Long = -1, Optional ByVal lFlags As Long = 0) As String
  Dim stBuffer As String
  Dim cwch As Long
  Dim pwz As Long
  Dim pwzBuffer As Long
  Dim lpUsedDefaultChar As Long
  If cpg = -1 Then cpg = GetACP()
  pwz = StrPtr(st)
  cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
  stBuffer = String$(cwch + 1, vbNullChar)
  pwzBuffer = StrPtr(stBuffer)
  cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
  WToA = Left$(stBuffer, cwch - 1)
End Function
 
'Use this to encode before you send...
Function EncodeUTF8(ByVal cnvUni As String)
  If cnvUni = vbNullString Then Exit Function
  EncodeUTF8 = WToA(cnvUni, CP_UTF8, 0)
End Function
 
'Use this to decode received strings...
Function DecodeUTF8(ByVal cnvUni As String)
  If cnvUni = vbNullString Then Exit Function
  cnvUni = WToA(cnvUni, CP_ACP)
  DecodeUTF8 = AToW(cnvUni, CP_UTF8)
End Function

'참고: http://blog.naver.com/PostView.nhn?blogId=knight50&logNo=80061829967&redirect=Dlog&widgetTypeCall=true
'--------------------------------------------------------------------------------------------------------

Public Function URLDecodeUTF8(ByVal pURL)

    Dim i, s1, s2, s3, u1, u2, result
    
On Error GoTo err

    pURL = Replace(pURL, "+", " ")
    For i = 1 To Len(pURL)
    
        If Mid(pURL, i, 1) = "=" Then
            s1 = CLng("&H" & Mid(pURL, i + 1, 2))
    
            '2바이트일 경우
            If ((s1 And &HC0) = &HC0) And ((s1 And &HE0) <> &HE0) Then
                s2 = CLng("&H" & Mid(pURL, i + 4, 2))
                u1 = (s1 And &H1C) / &H4
                u2 = ((s1 And &H3) * &H4 + ((s2 And &H30) / &H10)) * &H10
                u2 = u2 + (s2 And &HF)
                result = result & ChrW((u1 * &H100) + u2)
                i = i + 5
            
            '3바이트일 경우
            ElseIf (s1 And &HE0 = &HE0) Then
                s2 = CLng("&H" & Mid(pURL, i + 4, 2))
                s3 = CLng("&H" & Mid(pURL, i + 7, 2))
    
                u1 = ((s1 And &HF) * &H10)
                u1 = u1 + ((s2 And &H3C) / &H4)
                u2 = ((s2 And &H3) * &H4 + (s3 And &H30) / &H10) * &H10
                u2 = u2 + (s3 And &HF)
                result = result & ChrW((u1 * &H100) + u2)
                i = i + 8
                
            End If
        
        Else
            result = result & Mid(pURL, i, 1)
        End If
        
    Next
    
    URLDecodeUTF8 = result

On Error GoTo 0
    Exit Function

err:

End Function

'---------------------------------------------------
'javascript encodeURI 와 동일
Public Function URLEncodeUTF8(ByVal szSource)

Dim szChar, WideChar, nLength, i, result
nLength = Len(szSource)

'szSource = Replace(szSource," ","+")

For i = 1 To nLength
 szChar = Mid(szSource, i, 1)

 If Asc(szChar) < 0 Then
  WideChar = CLng(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))

  If (WideChar And &HFF80) = 0 Then
   result = result & "%" & Hex(WideChar)
  ElseIf (WideChar And &HF000) = 0 Then
   result = result & _
   "%" & Hex(CInt((WideChar And &HFFC0) / 64) Or &HC0) & _
   "%" & Hex(WideChar And &H3F Or &H80)
  Else
   result = result & _
   "%" & Hex(CInt((WideChar And &HF000) / 4096) Or &HE0) & _
   "%" & Hex(CInt((WideChar And &HFFC0) / 64) And &H3F Or &H80) & _
   "%" & Hex(WideChar And &H3F Or &H80)
  End If
 Else
  result = result + szChar
 End If
Next
URLEncodeUTF8 = result
End Function



Public Function setRndName() As String
mGetRndName = Format(Now(), "DDHHSS") + Rnd()
setRndName = mGetRndName
End Function

Public Function getRndName() As String
getRndName = mGetRndName
End Function


Public Function isEng(str As String) As Boolean

Dim nChkLenB As Long
Dim nChkLenH As Long

nChkLenB = LenB(StrConv(str, vbFromUnicode))
nChkLenH = Len(str)

If (nChkLenB <> nChkLenH) Then
    isEng = False
Else
    
    Dim char_first As String
    
    char_first = Left(str, 1)
    
    Dim asc_first As Integer
    
    asc_first = Asc(char_first)
    
    If asc_first >= Asc("0") And asc_first <= Asc("9") Then
        isEng = False
    Else
        isEng = True
    End If
    
End If

End Function

Public Function CRC16(ByVal S As String) As String
    CRC16 = CRC(S, False)
End Function

Public Function CRC32(ByVal S As String) As String
    CRC32 = CRC(S, True)
End Function

Private Function CRC(ByVal S As String, ByVal IsCRC32 As Boolean) As String
   ' Initializes CRC, updates CRC for each char in S, and finishes
   ' off at the end.
   ' Examining this function, and "CRCUpdate", should make it obvious
   ' how to do a CRC of a large binary file.
   ' 32 bit Returns zero for zero length string (probably why the CRC is
   ' complimented at the end.)
   Dim L As Long
   Dim LCRC As Long
   Dim ICRC As Integer
   
   Select Case IsCRC32
      Case True
         ' Initialise: this is part of the CRC32 protocol.
         LCRC = &HFFFFFFFF
         
         ' Update for each byte in S
         For L = 1 To LenH(S)
            CRCUpdate LCRC, Asc(MidH$(S, L, 1))
         Next
         
         ' Finally flip all bits, again, just part of the prorocol.
         LCRC = Not LCRC
         
         ' Format the long CRC value as a hex string.
         ' Insert leading zeros if required.
         CRC = Hex$(LCRC): Do While Len(CRC) < 8: CRC = "0" & CRC: Loop
         
      Case Else
         ' Initialise. Integer maths used.
         ICRC = &HFFFF
         
         ' Update for each byte in S
         For L = 1 To LenH(S): CRCUpdate ICRC, Asc(MidH$(S, L, 1)): Next
         
         ' Format the long CRC value as a hex string.
         ' Insert leading zeros if required.
         CRC = Hex$(ICRC): Do While Len(CRC) < 4: CRC = "0" & CRC: Loop
         
   End Select
End Function

Private Sub CRCUpdate(ByRef CRC, ByVal B As Byte)
   ' Note no type declaration for CRC, as a long or integer can be passed.
   Const Polynomial16 As Integer = &HA001
   Const Polynomial32 As Long = &HEDB88320
   Dim bits As Byte
   
   CRC = CRC Xor B
   For bits = 0 To 7
      Select Case (CRC And &H1) ' test LSB
            Case 0
               ' LSB zero, just shift.
               CRC = rightShift(CRC)
            Case Else
               ' only xor with polynomial if lsb set.
               Select Case VarType(CRC)
                  Case vbLong
                     CRC = rightShift(CRC) Xor Polynomial32
                  Case Else
                     CRC = rightShift(CRC) Xor Polynomial16
               End Select
      End Select
   Next
End Sub

Private Function rightShift(ByVal v) As Long
   ' Note no type declaration for V, as a long or integer can be passed.
   ' Self-explanatory. The final line is essential (maybe
   ' not obvious) because the number is signed.
   Select Case VarType(v)
      Case vbLong
         rightShift = v And &HFFFFFFFE
         rightShift = rightShift \ &H2
         rightShift = rightShift And &H7FFFFFFF
      Case Else
         rightShift = v And &HFFFE
         rightShift = rightShift \ &H2
         rightShift = rightShift And &H7FFF
   End Select

End Function

' string to byte array
Private Function UnicodeStringToBytes(ByVal str As String) As Byte()
    
    Dim B() As Byte
    
    ReDim B(1 To Len(str))
    Dim i As Long
    
    For i = 1 To Len(str)
        B(i) = Asc(Mid(str, i, 1))
        Debug.Print B(i)
    Next
    UnicodeStringToBytes = B
End Function

Public Function toclipboard(txt)
    Clipboard.Clear
    Clipboard.SetText txt
End Function

''----------file exists--------
'Public Function FileExists(ByVal sFile As String) As Boolean
'    Dim lpFindFileData  As WIN32_FIND_DATA
'    Dim lFileHandle     As Long
'    Dim lRet            As Long
'    Dim sTemp           As String
'    Dim sFileExtension  As String
'    Dim sFileName       As String
'    Dim sFileData()     As String
'    Dim sFileToCompare  As String
'
'    If IsDirectory(sFile) = True Then
'        sFile = AddSlash(sFile) & "*.*"
'    End If
'
'    If InStr(sFile, ".") > 0 Then
'        sFileToCompare = GetFileTitle(sFile)
'        sFileData = Split(sFileToCompare, ".")
'        sFileName = sFileData(0)
'        sFileExtension = sFileData(1)
'    Else
'        Exit Function
'    End If
'
'    ' get a file handle
'    lFileHandle = FindFirstFile(sFile, lpFindFileData)
'    If lFileHandle <> -1 Then
'        If sFileName = "*" Or sFileExtension = "*" Then
'            FileExists = True
'        Else
'            Do Until lRet = ERROR_NO_MORE_FILES
'                ' if it is a file
'                If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
'                    sTemp = StrConv(RemoveNull(lpFindFileData.cFileName), vbProperCase)
'
'                    'remove LCase$ if you want the search to be case sensitive
'                    If LCase$(sTemp) = LCase$(sFileToCompare) Then
'                        FileExists = True ' file found
'                        Exit Do
'                    End If
'                End If
'                'based on the file handle iterate through all files and dirs
'                lRet = FindNextFile(lFileHandle, lpFindFileData)
'                If lRet = 0 Then Exit Do
'            Loop
'        End If
'    End If
'
'    ' close the file handle
'    lRet = FindClose(lFileHandle)
'End Function
'
'Private Function IsDirectory(ByVal sFile As String) As Boolean
'    On Error Resume Next
'    IsDirectory = ((GetAttr(sFile) And vbDirectory) = vbDirectory)
'End Function
'
'Private Function RemoveNull(ByVal strString As String) As String
'    Dim intZeroPos As Integer
'
'    intZeroPos = InStr(strString, Chr$(0))
'    If intZeroPos > 0 Then
'        RemoveNull = Left$(strString, intZeroPos - 1)
'    Else
'        RemoveNull = strString
'    End If
'End Function
'
'Public Function GetFileTitle(ByVal sFileName As String) As String
'    GetFileTitle = Right$(sFileName, Len(sFileName) - InStrRev(sFileName, "\"))
'End Function
'
'Public Function AddSlash(ByVal strDirectory As String) As String
'    If InStrRev(strDirectory, "\") <> Len(strDirectory) Then
'        strDirectory = strDirectory + "\"
'    End If
'    AddSlash = strDirectory
'End Function
''------------file exists end-----------
