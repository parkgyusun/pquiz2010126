VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "로그온"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1045"
   Begin VB.CommandButton cmdDelete 
      Caption         =   "삭제"
      Height          =   375
      Left            =   3780
      TabIndex        =   14
      Top             =   2655
      Width           =   1050
   End
   Begin VB.ComboBox cboUserList 
      Height          =   315
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2655
      Width           =   2355
   End
   Begin VB.FileListBox objListFile 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdIP 
      Caption         =   "내컴퓨터IP"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   150
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtIPport 
      Height          =   270
      Left            =   1290
      TabIndex        =   8
      ToolTipText     =   "IP:[PORT default 3306]"
      Top             =   180
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   360
      Left            =   3735
      TabIndex        =   4
      Tag             =   "1049"
      Top             =   1215
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "로그온"
      Default         =   -1  'True
      Height          =   360
      Left            =   3735
      TabIndex        =   3
      Tag             =   "1048"
      Top             =   810
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1245
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   270
      Left            =   1305
      TabIndex        =   1
      Top             =   855
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "사용자목록"
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Tag             =   "1047"
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "-문의:gyusun93@daum.net 박규선-"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00461C1A&
      Height          =   645
      Index           =   1
      Left            =   315
      TabIndex        =   10
      Top             =   2205
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabels 
      Caption         =   "서버:"
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   7
      Tag             =   "1046"
      Top             =   180
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   180
      X2              =   4890
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "이메일은 아이디로 사용되며 인증번호는 비밀번호로 사용됩니다."
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00461C1A&
      Height          =   645
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1755
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C84B42&
      BorderWidth     =   2
      FillColor       =   &H00F4A48C&
      FillStyle       =   0  'Solid
      Height          =   870
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1665
      Width           =   4695
   End
   Begin VB.Label lblLabels 
      Caption         =   "인증번호:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Tag             =   "1047"
      Top             =   1290
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "이메일:"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Tag             =   "1046"
      ToolTipText     =   "이메일 혹은 아이디"
      Top             =   870
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'새로운 사용자로 등록하는 방법:
'사용자 아이디 이메일에는 공백을 넣고
' 우선 기존 계정으로 로그인 해야 한다.
' 사용자 변경에서
' 아이디에는 공백넣고 이메일 넣고
' 암호에는 한글로 '가입' 이라고 넣는다.

'최초가입시 휴대전화를 통해 인증번호를 부여 받습니다.
Option Explicit
Option Base 0

Dim Pos As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public gLoginFailCnt  As Integer

Public OK As Boolean



Private Sub cboUserList_Change()
txtUserName.Text = cboUserList.Text
    txtPassword.Text = ""
End Sub

Private Sub cboUserList_Click()
   txtUserName.Text = cboUserList.Text
    txtPassword.Text = ""
    
End Sub

Private Sub cmdDelete_Click()
    
    If MsgBox("빠른 로그인 목록에서 삭제하시겠습니까?" + vbNewLine + cboUserList.Text, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Call cFTP.deleteProfile(cboUserList.Text)
        cboUserList.removeItem (cboUserList.ListIndex)
    End If
    
End Sub

'Private WithEvents MyAgent As AgentObjectsCtl.Agent



Private Sub cmdIP_Click()
    txtIPport.Text = "127.0.0.1"
End Sub

Private Sub Form_Activate()

'cTmr.Interval = 2000
'cTmr.Enabled = True
    cmdOK.SetFocus
Screen.MousePointer = vbDefault
Call selall(txtUserName)

If cboUserList.ListCount = 0 Then
    
    cboListFill
    
End If



End Sub


Private Sub cboListFill()

Dim strArr() As String
            
    strArr = cFTP.GetProfiles()
    
    Dim j As Long
    Dim maxIdx As Long
    
    maxIdx = UBound(strArr)
    
    Call cboUserList.Clear
    
    Call QuickSort(strArr, LBound(strArr), UBound(strArr))
    
    For j = 0 To maxIdx
        If 0 < Len(strArr(j)) And strArr(j) <> "MAIN" Then
            cboUserList.AddItem strArr(j)
        End If
    Next
    
    If txtUserName <> "" Then
        On Error Resume Next
        cboUserList.Text = txtUserName
        '애매하게 삭제된 경우 오류가 있을 수 있겠다.
        On Error GoTo 0
        
    End If
    

End Sub

Private Sub Form_Load()

Me.Caption = "로그온 - BUILD NO:" + Replace(Module1.BUILD_NUMBER, "-", "") + ".00"
'On Error GoTo AgentErr

'    Dim lb As Long, pa As Long
'    lb = LoadLibrary("Agentctl.dll") '오류 체크(오류 안남...)
'    pa = GetProcAddress(lb, "showdefaultcharacterproperties")
'    Call CallWindowProc(pa, Me.hWnd, "", ByVal 0&, ByVal 0&)
'    FreeLibrary lb
    
    Dim sBuffer As String
    Dim lSize As Long
    Dim i As Long
    Dim delta As Long

    If STRCON = "" And getParamValue(Command, "h", "") = "127.0.0.1" Then
        '파라메터에 127.0.0.1 이라고 명시되어 호출한 경우
        
        If InStr(LCase(Command), "/mdb") > 0 Then
            If Right(App.Path, 1) = "\" Then
                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
            Else
                '본사 계정 학교마다 2개를 바꿀것. db명과 user명 (중요!) 암호는 넣으면 암됨(왜냐. ODBC드라이버가 지원해주지 못함)
                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & "\" & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
            End If
        Else
            STRCON = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mysql;DESC=;DATABASE=" + getParamValue(Command, "d", "pocket_znlwm_000001") + ";SERVER=" + getParamValue(Command, "h", "127.0.0.1") + ";UID=" + getParamValue(Command, "u", "changeit") + ";PASSWORD=" + getParamValue(Command, "p", "") + ";PORT=" + getParamValue(Command, "P", "3306") + ";SOCKET=;OPTION=" + getParamValue(Command, "o", "2097155") + ";STMT=;"""
        End If
    
    
    
        sSql = "select userid,con_ymd from tu01 order by con_ymd desc, con_time desc"
    
        Con_Open
    
        Set rs = Fn_SQLExec(sSql).rs
        If rs.EOF Then
            '새로운 사용자 등록
            sBuffer = Space$(255)
            lSize = Len(sBuffer)
            Call GetUserName(sBuffer, lSize)
            If lSize > 0 Then
                txtUserName.Text = Left$(sBuffer, lSize)
    
            Else
                txtUserName.Text = vbNullString
                txtUserName.SetFocus
            End If
            lblLabels(0).Caption = "등록할 이름:"
        Else
            txtUserName.Text = cFTP.Profile.LastUser ' rs(0)
            
            Do Until rs.EOF
                delta = Abs(DateDiff("d", str2Date(rs(1)), Now()))
                
                If delta < 365 Then
                    i = i + 1
                    ReDim Preserve gArrUsers(0 To i - 1) '1년 사용 유저이면서 최대 3명
                    gArrUsers(i - 1) = rs(0)
                    lblLabels(0).FontBold = True
                    lblLabels(0).ToolTipText = "사용자를 변경하기 위해 더블클릭 하세요."
                Else
                    Exit Do
                End If
                Call rs.MoveNext
            Loop
            
        End If
            txtIPport.Text = "(내컴퓨터)"
            txtIPport.Enabled = False
    Else
        If gUserid = "" Then
            txtUserName.Text = cFTP.Profile.LastUser
            
            cboListFill
            
        Else
            txtUserName.Text = gUserid
        End If
        If STRCON = "" Then
            
            txtIPport.Text = getParamValue(Command, "h", "127.0.0.1") & ":" & getParamValue(Command, "P", "3306")
            txtIPport.Text = Replace(txtIPport.Text, ":3306", "")
            
            If txtIPport.Text <> "" And Mid(LCase(getParamValue(Command, "D", "no")), 1, 1) = "y" Then
                cmdIP.Visible = True
                txtIPport.Enabled = True
            ElseIf txtIPport.Text <> "" And Mid(LCase(getParamValue(Command, "D", "no")), 1, 1) = "n" Then
            
                cmdIP.Visible = False
                txtIPport.Enabled = False
                
            ElseIf txtIPport.Text = "" And Mid(LCase(getParamValue(Command, "D", "no")), 1, 1) = "y" Then
                cmdIP.Visible = True
                txtIPport.Enabled = True
            Else
                cmdIP.Visible = False
                txtIPport.Enabled = True
            End If
            
        Else
            txtIPport.Enabled = False
                
            If gStrIP = "" Then
                txtIPport.Text = "(내컴퓨터)"
            Else
                txtIPport.Text = gStrIP & ":" & gStrPort
                txtIPport.Text = Replace(txtIPport.Text, ":3306", "")
            End If
        End If
    End If
    
    
    
'    If txtUserName.Text = "" Then
'            Call Shell(App.Path + "\" & "Windows6.1-KB969168-x86.msu", vbNormalFocus)
'    End If
'
'Exit Sub
'AgentErr:
'
'    MsgBox "먼저 MS Agent 핫픽스(Windows6.1-KB969168-x86.msu)를 설치하세요." & vbNewLine & "[" & err.Description & "]", vbCritical, "오류[" & err.Number & "]"
'
'    Call Shell(App.Path + "\" & "Windows6.1-KB969168-x86.msu", vbNormalFocus)
'
'    End

End Sub

Private Sub cmdCancel_Click()
If gChangUser = False Then
    If Not (fMainForm Is Nothing) Then
        Unload fMainForm
    End If
    End
Else
    Unload Me
End If
'    OK = False
'    Me.Hide
End Sub


Private Sub cmdOK_Click()

objListFile.Path = App.Path & "\" & "tmp"
objListFile.FileName = "$*.htm"

Dim i As Long, cnt As Long
cnt = objListFile.ListCount

For i = 0 To cnt - 1
    DeleteFile objListFile.Path & "\" & objListFile.List(i)
Next i

    cmdOK.Enabled = False
    Dim isOk As Boolean
    '작업: 정확한 암호에 대한 정확한 암호 확인을
    '테스트합니다.
    
    Dim arrServerInfo As Variant
    arrServerInfo = Split(txtIPport.Text, ":")


    If STRCON = "" Then
    
        gStrIP = arrServerInfo(LBound(arrServerInfo))
        
        If IsNumeric(Left(gStrIP, 1)) = False Then
            Dim ping_obj As New clsPing
            
            gStrIP = ping_obj.IPAddressFromHostName(gStrIP)
            
            
        End If
        
        gStrPort = "3306"
        
        If LBound(arrServerInfo) <> UBound(arrServerInfo) Then
            gStrPort = arrServerInfo(LBound(arrServerInfo) + 1)
        End If
                
        If InStr(LCase(Command), "/mdb") > 0 Then
            If Right(App.Path, 1) = "\" Then
                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
            Else
                '본사 계정 학교마다 2개를 바꿀것. db명과 user명 (중요!) 암호는 넣으면 암됨(왜냐. ODBC드라이버가 지원해주지 못함)
                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & "\" & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
            End If
        Else
            'STRCON = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mysql;DESC=;DATABASE=" + getParamValue(Command, "d", "pocket_znlwm_000001") + ";SERVER=" + getParamValue(Command, "h", "127.0.0.1") + ";UID=" + getParamValue(Command, "u", "changeit") + ";PASSWORD=" + getParamValue(Command, "p", "") + ";PORT=" + getParamValue(Command, "P", "3306") + ";SOCKET=;OPTION=" + getParamValue(Command, "o", "2097155") + ";STMT=;"""
            STRCON = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mysql;DESC=;DATABASE=" + getParamValue(Command, "d", "pocket_znlwm_000001") + ";SERVER=" + gStrIP + ";UID=" + getParamValue(Command, "u", "changeit") + ";PASSWORD=" + getParamValue(Command, "p", "") + ";PORT=" + gStrPort + ";SOCKET=;OPTION=" + getParamValue(Command, "o", "2097155") + ";STMT=;"""
        End If
        
        txtIPport.Enabled = False
        
        
    End If
    
    Dim conflag As Boolean
    sSql = "select * from ver"
    
    If Con.State = 0 Then
        Con_Open
        conflag = True
    End If
    
    Dim isMutex As Boolean
    Set rs = Fn_SQLExec(sSql, , , , isMutex).rs
    
    If isMutex = True Then
        End '강제종료
    End If
    If Not rs.EOF Then
        gnBangmoonConstraint = rs("visit_welcome")
        If rs("forceminver") > 20121224 Then 'DDD 중요!!! 여기에 버전체크가 들어감
        
            MsgBox "현재 프로그램은 호환성에 문제가 있습니다. " + vbNewLine + "업그레이드후 재실행해주세요.[문의:gyusun93@daum.net 박규선]" + vbNewLine + "실행옵션에서 /noupdate를 제거후 실행해 주세요." + vbNewLine + "그러면 ftp연결로 업그레이드가 진행됩니다.", vbExclamation
            End
        
        End If
    End If
    
    If conflag Then
        Con_Close
    End If
    
    
    
    '초기작업중 창을 모달리스로 띄움
    
    Screen.MousePointer = vbHourglass
    
    
    sSql = "select userid,userpass,con_ymd,con_time,logincnt,islogon,usergrade from tu01 where userid='" & Trim(txtUserName.Text) & "' and userpass=md5('" + txtPassword.Text + "')"
    
    Con_Open
    Set rs = Fn_SQLExec(sSql).rs
    If rs.EOF Then
        If Left(txtUserName.Text, 1) <> " " And InStr(txtUserName.Text, "@") = 0 Then
                Call MsgBox("올바른 이메일주소가 아닙니다.(새로운 사용자 등록 실패)", vbExclamation)
                Call selall(txtUserName)
                isOk = False
        
        '새로운사용자 등록
        ElseIf txtPassword.Text = "rkdlq" And gUserid <> "" Then 'rkdlq가입
            If gLogonCnt >= gnBangmoonConstraint Then
                If MsgBox("새로운 사용자로 등록하시겠습니까?", vbQuestion + vbYesNo, "질문") = vbYes Then
                'MAKE NEW USER
                    If Left(txtUserName.Text, 1) = " " Then
                        '이메일 주소형식이 아닌 사용자 이름의 형태를 원하면 앞에 공백을 넣을것.(이스트에그임)
                        txtUserName.Text = Trim(txtUserName.Text)
                        isOk = True
                    ElseIf (InStr(txtUserName.Text, "@") = 0) Or (InStr(txtUserName.Text, ".") = 0) Then
                        Call MsgBox("올바른 이메일주소가 아닙니다.", vbExclamation)
                        Call selall(txtUserName)
                        isOk = False
                    ElseIf Len(txtUserName.Text) > 50 Then
                        Call MsgBox("메일주소가 너무깁니다.[Max:50]", vbExclamation)
                        Call selall(txtUserName)
                        isOk = False
                    Else
                        isOk = True
                    End If
                    
                    If Len(txtUserName.Text) > 0 And isOk = True Then
                        'use progress bar start
                        Label1(0).Tag = Label1(0).Caption
                        Label1(0).Caption = "Waiting Please..."
                        Call MAKE_NEW_USER(txtUserName.Text, Mid(txtPassword.Text, 6), gUserid)
                        'use progress bar end
                        Label1(0).Caption = Label1(0).Tag
                        OK = True
                        Me.Hide
                        'gUserid = txtUserName
                        cFTP.Class_Initialize
                    End If
                End If
            Else
                MsgBox "[" & gnBangmoonConstraint & "]일 이상 방문해야 새로운 사용자를 등록 할 수 있습니다.", vbExclamation
                Exit Sub
            End If
        Else
            gLoginFailCnt = gLoginFailCnt + 1
            
            If gLoginFailCnt >= 3 Then
                MsgBox "3번이상 로그온 실패하였습니다. 프로그램을 종료합니다.-인증번호 재설정 지원을 받으시려면 gyusun93@daum.net으로 메일을 보내주세요.-", vbExclamation, "로그온"
                If Not (fMainForm Is Nothing) Then
                    Unload fMainForm
                End If
                End
            Else
                MsgBox "아이디(이메일)나 인증번호가 일치하지 않습니다.", vbExclamation, "로그온 실패"
            End If
            
            cmdOK.Enabled = True
            txtUserName.SetFocus
            txtUserName.SelStart = 0
            txtUserName.SelLength = Len(txtUserName.Text)
Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
    Else
        txtUserName.Text = Trim(txtUserName.Text)
        If rs(1) <> "" Then
            '암호일치 맞습니다.
            OK = True
            Me.Hide
            gUserid = txtUserName
            
            cFTP.Class_Initialize
            
            Dim ymd As String, hms As String, ymdhms, ymdnow As String
            ymd = rs(2)
            ymdnow = Format(Now, "YYYYMMDD")
            hms = rs(3)
            
            If rs("islogon") = 1 Then
                If MsgBox("새로운 연결로 접속 하시겠습니까?" & vbNewLine & "최근 로그인 시각: " & ymd & " " & hms, vbQuestion + vbYesNo, "포켓퀴즈") = vbNo Then
                    If Not (fMainForm Is Nothing) Then
                        Unload fMainForm
                    End If
                    End
                End If
            End If
            
            ymdhms = Mid(ymd, 1, 4) & "-" & Mid(ymd, 5, 2) & "-" & Mid(ymd, 7, 2) & " " & Mid(hms, 1, 2) & ":" & Mid(hms, 3, 2)
            
            gLogonCnt = rs("logincnt")
            gbIsSuperAdmin = (rs("usergrade") = 99)
            gPassWord = txtPassword.Text
            If gLogonCnt = 0 And gPassWord = "" Then
                MsgBox "포켓퀴즈 처음사용을 축하합니다.~" & vbNewLine & "먼저 회원정보변경메뉴에서 비밀번호를 설정하세요.", vbExclamation, "포켓퀴즈"
            Else
                'MsgBox "최근 로그온 일시 " + ymdhms, vbExclamation, "포켓퀴즈"
            End If
            
            gSessionId = Format(Now, "HHMMSS")
            
            
            
            If ymd = ymdnow Then
                sSql = "update tu01 set con_time='" & gSessionId & "',islogon=1 where userid='" & gUserid & "'"
            Else
                sSql = "update tu01 set con_ymd='" & ymdnow & "',con_time='" & gSessionId & "',logincnt=logincnt+1,islogon=1 where userid='" & gUserid & "'"
            End If
            
            Con_Open
            Fn_SQLExec (sSql)
'            Con_Close
            
        Else
            If gLoginFailCnt >= 3 Then
                MsgBox "3번이상 인증번호가 틀렸습니다. 프로그램을 종료합니다.-인증번호 재설정 지원을 받으시려면 gyusun93@daum.net으로 메일을 보내주세요.-", vbExclamation, "로그온"
                If Not (fMainForm Is Nothing) Then
                    Unload fMainForm
                End If
                End
            End If
            gLoginFailCnt = gLoginFailCnt + 1
            MsgBox "인증번호가 정확하지 않습니다. 다시 시도하십시오!", vbExclamation, "로그온"
            txtPassword.SetFocus
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword.Text)
            
        End If
    End If
'    Con_Close
    Screen.MousePointer = vbDefault
    cFTP.Profile.LastUser = txtUserName.Text
    cmdOK.Enabled = True
End Sub

Private Sub MAKE_NEW_USER(userid As String, PWD As String, mentoID As String)
'gUserid = userid
    On Error GoTo ErrTrap
    Con_Open

    Dim ymd As String
    
    ymd = Format(Now, "yyyymmdd")
        
    Dim RS2 As New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
 

'----------------------------------------------------------------------------
'총계정원장 등록
'----------------------------------------------------------------------------
'    If Con.State = 1 Then
'        Call insert_totalaccount2
'    Else
'        Con_Open
'        Call insert_totalaccount2
'        Con_Close
'    End If
    
'----------------------------------------------------------------------------
'사용자 등록
'----------------------------------------------------------------------------
    sSql = "INSERT INTO TU01(userid,userpass,con_ymd,con_time,mentoid,signfirst_ymd) VALUES('" & userid & "',md5('" & PWD & "'),'" & ymd & "','" & Format(Now, "hhMmSS") & "','" & mentoID & "','" & ymd & "')"
    Fn_SQLExec (sSql)


'----------------------------------------------------------------------------
'과목별 사용자 등록 ts02
'디폴트 과목 등록
'----------------------------------------------------------------------------

    'sSql = "insert into ts02(subj,userid,startymd,endymd) "
    'sSql = sSql & vbCrLf & "(select a.subj,b.userid,'20000101','21001231' from ts01 a ,tu01 b where "
    'sSql = sSql & vbCrLf & "a.subj in ('영단어1','영속담1','영숙어(중)','영숙어1',"
    'sSql = sSql & vbCrLf & "'중1단어','한자','한자01','토익Voca','운전면허표지안전판',"
    'sSql = sSql & vbCrLf & "'운전01','운전02','운전03','운전04','운전05','운전06','운전07',"
    'sSql = sSql & vbCrLf & "'운전08','운전09','운전10','운전11','운전12','운전A','운전B',"
    'sSql = sSql & vbCrLf & "'운전C','운전D','일본어800','') and b.userid='" & userid & "')"
    
'    defaultinsert_subjes_string
''영단어1','영속담1','영숙어 (중)','영숙어1','중1단어','중2단어','중3단어','중학기초','일본어800','특목고입시용','토익Voca',
''신약성경','구약성경','구약:01','구약:02','구약:03','구약:04','구약:05','구약:06','구약:07','구약:08','구약:09','구약:10','구약:11','구약:12','구약:13','구약:14','구약:15','구약:16','구약:17','구약:18','구약:19','구약:20','구약:21','구약:22','구약:23','구약:24','구약:25','구약:26','구약:27','구약:28','구약:29','구약:30','구약:31','구약:32','구약:33','구약:34','구약:35','구약:36','구약:37','구약:38','구약:39','신약:40','신약:41','신약:42','신약:43','신약:44','신약:45','신약:46','신약:47','신약:48','신약:49','신약:50','신약:51','신약:52','신약:53','신약:54','신약:55','신약:56','신약:57','신약:58','신약:59','신약:60','신약:61','신약:62','신약:63','신약:64','신약:65','신약:66',
''한자','한자01','한자02','한자03','한자04','한자05','한자06','한자07','한자08','한자09','한자10','한자11'
','minus_10deg_','minus_11deg_','minus_12deg_','minus_13deg_','minus_14deg_'
','minus_15deg_','minus_16deg_','minus_17deg_','minus_18deg_','minus_19deg_','minus_1char_','minus_20deg_','plus_Atype_','plus_Btype_','plus_Ctype_','plus_Dtype_','plus_Etype_','plus_Ftype_','plus_Gtype_','plus_Htype_'

'
'
' insert into tsdefault(val,usebit) values('\'영단어1\',\'영숙어 (중)\',\'영숙어1\'
',\'중1단어\',\'중2단어\',\'중3단어\',\'중학기초\'
',\'히라가나\',\'카타카나\'
',\'일본어800\'
',\'특목고입시용\',\'토익Voca\'
',\'한영퀴즈\'
',\'한자\',\'한자01\',\'한자02\',\'한자03\',\'한자04\',\'한자05\',\'한자06\'
',\'한자07\',\'한자08\',\'한자09\',\'한자10\',\'한자11\'',1);

    Dim strDefaultTs01es As String
    
    sSql = "select seq,val from tsdefault order by usebit,seq desc limit 1" '마지막거 하나만 리스팅한다.
    
    Set rs = Fn_SQLExec(sSql).rs
    If Not rs.EOF Then
        strDefaultTs01es = rs("val")
    End If
    
    sSql = "insert into ts02(subj,userid,startymd,endymd) "
    sSql = sSql & vbCrLf & "(select a.subj,b.userid,'" & Format(Now(), "yyyy") & "0101','" & (Format(Now, "yyyy") + 100) & "1231' from ts01 a ,tu01 b where "
    sSql = sSql & vbCrLf & "a.subj in (" & strDefaultTs01es & _
    ") and b.userid='" & userid & "')"
    '사용자 삭제를 하려면 ts02 삭제 하면 되고 또 .
    'tg01,tp02,tu02,ts02,tu01,tp02,tp03,tt01,tt02,tt03 tu03 에서 userid를 해당것 삭제하면 된다.
    '특히 tu02가 많다.tp03도 많다.
    'th01은 지워도 되나. 사용자의 기록(내메모)가 지워진다.
    '
'mysql> delete from tg01 where userid like 'test%' ;
'Query OK, 5 rows affected (0.06 sec)
'
'mysql> delete from tp02 where userid like 'test%' ;
'Query OK, 3142 rows affected (0.89 sec)
'
'mysql> delete from tu02 where userid like 'test%' ;
'Query OK, 9767 rows affected (2.11 sec)
'
'mysql> delete from ts02 where userid like 'test%' ;
'Query OK, 444 rows affected (0.05 sec)
'
'mysql> delete from tu01 where userid like 'test%' ;
'Query OK, 7 rows affected (0.00 sec)
'
'mysql> delete from tp03 where userid like 'test%' ;
'Query OK, 13183 rows affected (5.28 sec)
'
'mysql> delete from tt01 where userid like 'test%' ;
'Query OK, 8 rows affected (0.00 sec)
'
'mysql> delete from tt02 where userid like 'test%' ;
'Query OK, 6589 rows affected (1.14 sec)
'
'mysql> delete from tt03 where userid like 'test%' ;
'Query OK, 6589 rows affected (0.69 sec)
'
'mysql> delete from tu03 where userid like 'test%' ;
'Query OK, 3142 rows affected (0.98 sec)

'+++++++++++++++++++++++++++++++++++++++++++
'다른 지우기 사례
'delete from tg01 where userid in (select userid from tu01 where signfirst_ymd='');
'delete from tp02 where userid in (select userid from tu01 where signfirst_ymd='');
'delete from tu02 where userid  in (select userid from tu01 where signfirst_ymd='');
'delete from ts02 where userid  in (select userid from tu01 where signfirst_ymd='');
'delete from tp03 where userid  in (select userid from tu01 where signfirst_ymd='');
'delete from tt01 where userid  in (select userid from tu01 where signfirst_ymd='');
'delete from tt02 where userid  in (select userid from tu01 where signfirst_ymd='');
'delete from tt03 where userid  in (select userid from tu01 where signfirst_ymd='');
'delete from tu03 where userid  in (select userid from tu01 where signfirst_ymd='');
'delete from tu01 where signfirst_ymd='';

'mysql> delete from tg01 where userid in (select userid from tu01 where signfirst_ymd='');
'Query OK, 1 row affected (0.23 sec)
'
'mysql> delete from tp02 where userid in (select userid from tu01 where signfirst_ymd='');
'Query OK, 0 rows affected (0.50 sec)
'
'mysql> delete from tu02 where userid  in (select userid from tu01 where signfirst_ymd='');
'Query OK, 271609 rows affected (54.94 sec)
'
'mysql> delete from ts02 where userid  in (select userid from tu01 where signfirst_ymd='');
'Query OK, 857 rows affected (2.67 sec)
'
'mysql> delete from tp03 where userid  in (select userid from tu01 where signfirst_ymd='');
'Query OK, 0 rows affected (5.44 sec)
'
'mysql> delete from tt01 where userid  in (select userid from tu01 where signfirst_ymd='');
'Query OK, 0 rows affected (0.02 sec)
'
'mysql> delete from tt02 where userid  in (select userid from tu01 where signfirst_ymd='');
'Query OK, 0 rows affected (0.33 sec)
'
'mysql> delete from tt03 where userid  in (select userid from tu01 where signfirst_ymd='');
'Query OK, 0 rows affected (0.47 sec)
'
'mysql> delete from tu03 where userid  in (select userid from tu01 where signfirst_ymd='');
'Query OK, 0 rows affected (0.38 sec)
'
'mysql> delete from tu01 where signfirst_ymd='';
'Query OK, 9 rows affected (0.00 sec)

'지운회원데이터
'장범익
'서동은
'aaa
'altnrwjs@ hanmail.net
'pquiz@ pquiz.codns.com
'pquiz2@ pquiz.codns.com
'pquiz3@ pquiz.codns.com
'새테스트1
'ps8629@ naver.com

    Fn_SQLExec (sSql)

'----------------------------------------------------------------------------
'포켓퀴즈정보 등록
'----------------------------------------------------------------------------
    
    sSql = "SELECT pocketnm FROM TP01"
    Set rs = Fn_SQLExec(sSql).rs
'    Do Until rs.EOF()
'
'        If con.State = 1 Then
'
'            Call insert_pocketinfo(rs(0))
'        Else
'            Con_Open
'                Call insert_pocketinfo(rs(0))
'            Con_Close
'        End If
'
'        rs.MoveNext
'    Loop
    
    
    Con_Close
    Screen.MousePointer = vbDefault
Exit Sub
ErrTrap:
    MsgBox err.Description, vbCritical
    
    
End Sub



Private Sub lblLabels_DblClick(Index As Integer)
    Dim siz As Long
    siz = getLength(gArrUsers)
    If siz < 2 Then Exit Sub '0 이나 1이면 빠져나간다.
    
    Do
        Pos = Pos Mod siz
        txtUserName.Tag = txtUserName
        txtUserName = gArrUsers(Pos)
        Pos = Pos + 1
    Loop Until txtUserName.Tag <> txtUserName
    If Pos = 1 Then
        lblLabels(0).FontBold = Not lblLabels(0).FontBold
    End If
End Sub

'Private Sub MyAgent_Command(ByVal UserInput As Object)
''samplet코드
'Select Case LCase(UserInput.Name)
'
'Case "bye bye"
'   MyAgent.Speak "Bye bye"
'   MyAgent.Play "Wave"
'   Set MyRequest = MyAgent.Hide
'
'Case "down"
'   MyAgent.GestureAt MyAgent.Left, Screen.Height / Screen.TwipsPerPixelY
'   MyAgent.Speak "All fall down."
'
'Case "happy"
'   MyAgent.Play "Pleased"
'   MyAgent.Speak "I'm happy."
'Case "hello"
'   MyAgent.Speak "Hello Merlin."
'   MyAgent.Play "DoMagic1"
'   MyAgent.Speak "Abracadabra, hocus pocus."
'   MyAgent.Play "DoMagic2"
'   MyAgent.Speak "Visual Basic now can"
'
'Case "left"
'   MyAgent.GestureAt 0, MyAgent.Height
'   MyAgent.Speak "Left is that away."
'
'Case "move"
'   MyAgent.MoveTo Screen.Width / Screen.TwipsPerPixelX * Rnd, _
'                  Screen.Height / Screen.TwipsPerPixelY * Rnd
'   MyAgent.Speak "I like flying."
'
'Case "right"
'   MyAgent.GestureAt Screen.Width / Screen.TwipsPerPixelX, MyAgent.Height
'   MyAgent.Speak "Keep to the right."
'Case "sad"
'   MyAgent.Play "Sad"
'   MyAgent.Speak "I'm sad."
'Case "up"
'   MyAgent.GestureAt MyAgent.Left, 0
'   MyAgent.Speak "The sky is falling."
'Case Else
'   MyAgent.Speak "I didn't understand you."
'   MyAgent.Play "Confused"
'End Select
'
'End Sub

'Public Sub insert_pocketinfo(pocketnm As String)
'Dim RS2 As New ADODB.Recordset
'
'Dim SSQL1 As String
'Dim SSQL2 As String
'Dim SSQL3 As String
'
'Dim ymd As String
'
'ymd = GETYMD()
'
'
'Dim CHASU As Integer
'
'SSQL1 = "SELECT MAX(CHASU) FROM TP02 WHERE USERID='" & gUserid & "' AND POCKETNM='" & pocketnm & "' "
'RS2.Open SSQL1, con, 1, 3
'If Not RS2.EOF Then
'    If IsNull(RS2(0)) Then
'        CHASU = 0
'    Else
'        CHASU = RS2(0) + 1
'    End If
'
'End If
'RS2.Close
'
'SSQL1 = "INSERT INTO TP02 VALUES('" & pocketnm & "'," & CHASU & ",'" & gUserid & "',0,'" & ymd & "')"
'con.Execute SSQL1
'
'SSQL1 = "SELECT CONd  FROM TP01 WHERE POCKETNM='" & pocketnm & "'"
'RS2.Open SSQL1, con, 1, 3
'
'
'COND = RS2(0)
'OBJTABLE = "TU02" 'RS2(1)
'RS2.Close
'
'SSQL1 = "SELECT * FROM " & OBJTABLE & " WHERE " & COND
'
'RS2.Open SSQL1, con, 1, 3
'
'i = 0
'Do Until RS2.EOF
'    i = i + 1
'    SSQL1 = "INSERT INTO TP03 VALUES('" & pocketnm & "'," & CHASU & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
'    Call con.Execute(SSQL1, affected)
'    Debug.Assert affected > 0
'    RS2.MoveNext
'
'Loop
'
'SSQL1 = "insert into tu03 values('" & pocketnm & "'," & CHASU & ",'" & gUserid & "',1,0,10)"
'Call con.Execute(SSQL1, affected)
'Debug.Assert affected > 0
'RS2.Close
'
'End Sub





'' sort the array
'    QuickSort MyStrArray, LBound(MyStrArray), UBound(MyStrArray)
'
'    ' print the sorted string to the immediate window
'    Debug.Print vbNewLine & "Sorted strings:"
'    For K = LBound(MyStrArray) To UBound(MyStrArray)
'        Debug.Print MyStrArray(K)
'    Next K

 
Private Sub QuickSort(C() As String, ByVal First As Long, ByVal Last As Long)
    '
    '  Made by Michael Ciurescu (CVMichael from vbforums.com)
    '  Original thread: [url]http://www.vbforums.com/showthread.php?t=231925[/url]
    '
    Dim Low As Long, High As Long
    Dim MidValue As String
    
    Low = First
    High = Last
    MidValue = C((First + Last) \ 2)
    
    Do
        While C(Low) < MidValue
            Low = Low + 1
        Wend
        
        While C(High) > MidValue
            High = High - 1
        Wend
        
        If Low <= High Then
            swap C(Low), C(High)
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then QuickSort C, First, High
    If Low < Last Then QuickSort C, Low, Last
End Sub
 
Private Sub swap(ByRef a As String, ByRef B As String)
    Dim T As String
    
    T = a
    a = B
    B = T
End Sub
