VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�α׿�"
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
      Caption         =   "����"
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
      Caption         =   "����ǻ��IP"
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
      Caption         =   "���"
      Height          =   360
      Left            =   3735
      TabIndex        =   4
      Tag             =   "1049"
      Top             =   1215
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�α׿�"
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
      Caption         =   "����ڸ��"
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
      Caption         =   "-����:gyusun93@daum.net �ڱԼ�-"
      BeginProperty Font 
         Name            =   "����ü"
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
      Caption         =   "����:"
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
      Caption         =   "�̸����� ���̵�� ���Ǹ� ������ȣ�� ��й�ȣ�� ���˴ϴ�."
      BeginProperty Font 
         Name            =   "����ü"
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
      Caption         =   "������ȣ:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Tag             =   "1047"
      Top             =   1290
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "�̸���:"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Tag             =   "1046"
      ToolTipText     =   "�̸��� Ȥ�� ���̵�"
      Top             =   870
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ο� ����ڷ� ����ϴ� ���:
'����� ���̵� �̸��Ͽ��� ������ �ְ�
' �켱 ���� �������� �α��� �ؾ� �Ѵ�.
' ����� ���濡��
' ���̵𿡴� ����ְ� �̸��� �ְ�
' ��ȣ���� �ѱ۷� '����' �̶�� �ִ´�.

'���ʰ��Խ� �޴���ȭ�� ���� ������ȣ�� �ο� �޽��ϴ�.
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
    
    If MsgBox("���� �α��� ��Ͽ��� �����Ͻðڽ��ϱ�?" + vbNewLine + cboUserList.Text, vbYesNo + vbQuestion, App.Title) = vbYes Then
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
        '�ָ��ϰ� ������ ��� ������ ���� �� �ְڴ�.
        On Error GoTo 0
        
    End If
    

End Sub

Private Sub Form_Load()

Me.Caption = "�α׿� - BUILD NO:" + Replace(Module1.BUILD_NUMBER, "-", "") + ".00"
'On Error GoTo AgentErr

'    Dim lb As Long, pa As Long
'    lb = LoadLibrary("Agentctl.dll") '���� üũ(���� �ȳ�...)
'    pa = GetProcAddress(lb, "showdefaultcharacterproperties")
'    Call CallWindowProc(pa, Me.hWnd, "", ByVal 0&, ByVal 0&)
'    FreeLibrary lb
    
    Dim sBuffer As String
    Dim lSize As Long
    Dim i As Long
    Dim delta As Long

    If STRCON = "" And getParamValue(Command, "h", "") = "127.0.0.1" Then
        '�Ķ���Ϳ� 127.0.0.1 �̶�� ��õǾ� ȣ���� ���
        
        If InStr(LCase(Command), "/mdb") > 0 Then
            If Right(App.Path, 1) = "\" Then
                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
            Else
                '���� ���� �б����� 2���� �ٲܰ�. db��� user�� (�߿�!) ��ȣ�� ������ �ϵ�(�ֳ�. ODBC����̹��� ���������� ����)
                STRCON = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & "\" & MDBFILENM & """;Jet OLEDB:Database Passwor" & GPWD2 & GPWD3 & ";User Id=Admin;password;"
            End If
        Else
            STRCON = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mysql;DESC=;DATABASE=" + getParamValue(Command, "d", "pocket_znlwm_000001") + ";SERVER=" + getParamValue(Command, "h", "127.0.0.1") + ";UID=" + getParamValue(Command, "u", "changeit") + ";PASSWORD=" + getParamValue(Command, "p", "") + ";PORT=" + getParamValue(Command, "P", "3306") + ";SOCKET=;OPTION=" + getParamValue(Command, "o", "2097155") + ";STMT=;"""
        End If
    
    
    
        sSql = "select userid,con_ymd from tu01 order by con_ymd desc, con_time desc"
    
        Con_Open
    
        Set rs = Fn_SQLExec(sSql).rs
        If rs.EOF Then
            '���ο� ����� ���
            sBuffer = Space$(255)
            lSize = Len(sBuffer)
            Call GetUserName(sBuffer, lSize)
            If lSize > 0 Then
                txtUserName.Text = Left$(sBuffer, lSize)
    
            Else
                txtUserName.Text = vbNullString
                txtUserName.SetFocus
            End If
            lblLabels(0).Caption = "����� �̸�:"
        Else
            txtUserName.Text = cFTP.Profile.LastUser ' rs(0)
            
            Do Until rs.EOF
                delta = Abs(DateDiff("d", str2Date(rs(1)), Now()))
                
                If delta < 365 Then
                    i = i + 1
                    ReDim Preserve gArrUsers(0 To i - 1) '1�� ��� �����̸鼭 �ִ� 3��
                    gArrUsers(i - 1) = rs(0)
                    lblLabels(0).FontBold = True
                    lblLabels(0).ToolTipText = "����ڸ� �����ϱ� ���� ����Ŭ�� �ϼ���."
                Else
                    Exit Do
                End If
                Call rs.MoveNext
            Loop
            
        End If
            txtIPport.Text = "(����ǻ��)"
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
                txtIPport.Text = "(����ǻ��)"
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
'    MsgBox "���� MS Agent ���Ƚ�(Windows6.1-KB969168-x86.msu)�� ��ġ�ϼ���." & vbNewLine & "[" & err.Description & "]", vbCritical, "����[" & err.Number & "]"
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
    '�۾�: ��Ȯ�� ��ȣ�� ���� ��Ȯ�� ��ȣ Ȯ����
    '�׽�Ʈ�մϴ�.
    
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
                '���� ���� �б����� 2���� �ٲܰ�. db��� user�� (�߿�!) ��ȣ�� ������ �ϵ�(�ֳ�. ODBC����̹��� ���������� ����)
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
        End '��������
    End If
    If Not rs.EOF Then
        gnBangmoonConstraint = rs("visit_welcome")
        If rs("forceminver") > 20121224 Then 'DDD �߿�!!! ���⿡ ����üũ�� ��
        
            MsgBox "���� ���α׷��� ȣȯ���� ������ �ֽ��ϴ�. " + vbNewLine + "���׷��̵��� ��������ּ���.[����:gyusun93@daum.net �ڱԼ�]" + vbNewLine + "����ɼǿ��� /noupdate�� ������ ������ �ּ���." + vbNewLine + "�׷��� ftp����� ���׷��̵尡 ����˴ϴ�.", vbExclamation
            End
        
        End If
    End If
    
    If conflag Then
        Con_Close
    End If
    
    
    
    '�ʱ��۾��� â�� ��޸����� ���
    
    Screen.MousePointer = vbHourglass
    
    
    sSql = "select userid,userpass,con_ymd,con_time,logincnt,islogon,usergrade from tu01 where userid='" & Trim(txtUserName.Text) & "' and userpass=md5('" + txtPassword.Text + "')"
    
    Con_Open
    Set rs = Fn_SQLExec(sSql).rs
    If rs.EOF Then
        If Left(txtUserName.Text, 1) <> " " And InStr(txtUserName.Text, "@") = 0 Then
                Call MsgBox("�ùٸ� �̸����ּҰ� �ƴմϴ�.(���ο� ����� ��� ����)", vbExclamation)
                Call selall(txtUserName)
                isOk = False
        
        '���ο����� ���
        ElseIf txtPassword.Text = "rkdlq" And gUserid <> "" Then 'rkdlq����
            If gLogonCnt >= gnBangmoonConstraint Then
                If MsgBox("���ο� ����ڷ� ����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbYes Then
                'MAKE NEW USER
                    If Left(txtUserName.Text, 1) = " " Then
                        '�̸��� �ּ������� �ƴ� ����� �̸��� ���¸� ���ϸ� �տ� ������ ������.(�̽�Ʈ������)
                        txtUserName.Text = Trim(txtUserName.Text)
                        isOk = True
                    ElseIf (InStr(txtUserName.Text, "@") = 0) Or (InStr(txtUserName.Text, ".") = 0) Then
                        Call MsgBox("�ùٸ� �̸����ּҰ� �ƴմϴ�.", vbExclamation)
                        Call selall(txtUserName)
                        isOk = False
                    ElseIf Len(txtUserName.Text) > 50 Then
                        Call MsgBox("�����ּҰ� �ʹ���ϴ�.[Max:50]", vbExclamation)
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
                MsgBox "[" & gnBangmoonConstraint & "]�� �̻� �湮�ؾ� ���ο� ����ڸ� ��� �� �� �ֽ��ϴ�.", vbExclamation
                Exit Sub
            End If
        Else
            gLoginFailCnt = gLoginFailCnt + 1
            
            If gLoginFailCnt >= 3 Then
                MsgBox "3���̻� �α׿� �����Ͽ����ϴ�. ���α׷��� �����մϴ�.-������ȣ �缳�� ������ �����÷��� gyusun93@daum.net���� ������ �����ּ���.-", vbExclamation, "�α׿�"
                If Not (fMainForm Is Nothing) Then
                    Unload fMainForm
                End If
                End
            Else
                MsgBox "���̵�(�̸���)�� ������ȣ�� ��ġ���� �ʽ��ϴ�.", vbExclamation, "�α׿� ����"
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
            '��ȣ��ġ �½��ϴ�.
            OK = True
            Me.Hide
            gUserid = txtUserName
            
            cFTP.Class_Initialize
            
            Dim ymd As String, hms As String, ymdhms, ymdnow As String
            ymd = rs(2)
            ymdnow = Format(Now, "YYYYMMDD")
            hms = rs(3)
            
            If rs("islogon") = 1 Then
                If MsgBox("���ο� ����� ���� �Ͻðڽ��ϱ�?" & vbNewLine & "�ֱ� �α��� �ð�: " & ymd & " " & hms, vbQuestion + vbYesNo, "��������") = vbNo Then
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
                MsgBox "�������� ó������� �����մϴ�.~" & vbNewLine & "���� ȸ����������޴����� ��й�ȣ�� �����ϼ���.", vbExclamation, "��������"
            Else
                'MsgBox "�ֱ� �α׿� �Ͻ� " + ymdhms, vbExclamation, "��������"
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
                MsgBox "3���̻� ������ȣ�� Ʋ�Ƚ��ϴ�. ���α׷��� �����մϴ�.-������ȣ �缳�� ������ �����÷��� gyusun93@daum.net���� ������ �����ּ���.-", vbExclamation, "�α׿�"
                If Not (fMainForm Is Nothing) Then
                    Unload fMainForm
                End If
                End
            End If
            gLoginFailCnt = gLoginFailCnt + 1
            MsgBox "������ȣ�� ��Ȯ���� �ʽ��ϴ�. �ٽ� �õ��Ͻʽÿ�!", vbExclamation, "�α׿�"
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
'�Ѱ������� ���
'----------------------------------------------------------------------------
'    If Con.State = 1 Then
'        Call insert_totalaccount2
'    Else
'        Con_Open
'        Call insert_totalaccount2
'        Con_Close
'    End If
    
'----------------------------------------------------------------------------
'����� ���
'----------------------------------------------------------------------------
    sSql = "INSERT INTO TU01(userid,userpass,con_ymd,con_time,mentoid,signfirst_ymd) VALUES('" & userid & "',md5('" & PWD & "'),'" & ymd & "','" & Format(Now, "hhMmSS") & "','" & mentoID & "','" & ymd & "')"
    Fn_SQLExec (sSql)


'----------------------------------------------------------------------------
'���� ����� ��� ts02
'����Ʈ ���� ���
'----------------------------------------------------------------------------

    'sSql = "insert into ts02(subj,userid,startymd,endymd) "
    'sSql = sSql & vbCrLf & "(select a.subj,b.userid,'20000101','21001231' from ts01 a ,tu01 b where "
    'sSql = sSql & vbCrLf & "a.subj in ('���ܾ�1','���Ӵ�1','������(��)','������1',"
    'sSql = sSql & vbCrLf & "'��1�ܾ�','����','����01','����Voca','��������ǥ��������',"
    'sSql = sSql & vbCrLf & "'����01','����02','����03','����04','����05','����06','����07',"
    'sSql = sSql & vbCrLf & "'����08','����09','����10','����11','����12','����A','����B',"
    'sSql = sSql & vbCrLf & "'����C','����D','�Ϻ���800','') and b.userid='" & userid & "')"
    
'    defaultinsert_subjes_string
''���ܾ�1','���Ӵ�1','������ (��)','������1','��1�ܾ�','��2�ܾ�','��3�ܾ�','���б���','�Ϻ���800','Ư����Խÿ�','����Voca',
''�ž༺��','���༺��','����:01','����:02','����:03','����:04','����:05','����:06','����:07','����:08','����:09','����:10','����:11','����:12','����:13','����:14','����:15','����:16','����:17','����:18','����:19','����:20','����:21','����:22','����:23','����:24','����:25','����:26','����:27','����:28','����:29','����:30','����:31','����:32','����:33','����:34','����:35','����:36','����:37','����:38','����:39','�ž�:40','�ž�:41','�ž�:42','�ž�:43','�ž�:44','�ž�:45','�ž�:46','�ž�:47','�ž�:48','�ž�:49','�ž�:50','�ž�:51','�ž�:52','�ž�:53','�ž�:54','�ž�:55','�ž�:56','�ž�:57','�ž�:58','�ž�:59','�ž�:60','�ž�:61','�ž�:62','�ž�:63','�ž�:64','�ž�:65','�ž�:66',
''����','����01','����02','����03','����04','����05','����06','����07','����08','����09','����10','����11'
','minus_10deg_','minus_11deg_','minus_12deg_','minus_13deg_','minus_14deg_'
','minus_15deg_','minus_16deg_','minus_17deg_','minus_18deg_','minus_19deg_','minus_1char_','minus_20deg_','plus_Atype_','plus_Btype_','plus_Ctype_','plus_Dtype_','plus_Etype_','plus_Ftype_','plus_Gtype_','plus_Htype_'

'
'
' insert into tsdefault(val,usebit) values('\'���ܾ�1\',\'������ (��)\',\'������1\'
',\'��1�ܾ�\',\'��2�ܾ�\',\'��3�ܾ�\',\'���б���\'
',\'���󰡳�\',\'īŸī��\'
',\'�Ϻ���800\'
',\'Ư����Խÿ�\',\'����Voca\'
',\'�ѿ�����\'
',\'����\',\'����01\',\'����02\',\'����03\',\'����04\',\'����05\',\'����06\'
',\'����07\',\'����08\',\'����09\',\'����10\',\'����11\'',1);

    Dim strDefaultTs01es As String
    
    sSql = "select seq,val from tsdefault order by usebit,seq desc limit 1" '�������� �ϳ��� �������Ѵ�.
    
    Set rs = Fn_SQLExec(sSql).rs
    If Not rs.EOF Then
        strDefaultTs01es = rs("val")
    End If
    
    sSql = "insert into ts02(subj,userid,startymd,endymd) "
    sSql = sSql & vbCrLf & "(select a.subj,b.userid,'" & Format(Now(), "yyyy") & "0101','" & (Format(Now, "yyyy") + 100) & "1231' from ts01 a ,tu01 b where "
    sSql = sSql & vbCrLf & "a.subj in (" & strDefaultTs01es & _
    ") and b.userid='" & userid & "')"
    '����� ������ �Ϸ��� ts02 ���� �ϸ� �ǰ� �� .
    'tg01,tp02,tu02,ts02,tu01,tp02,tp03,tt01,tt02,tt03 tu03 ���� userid�� �ش�� �����ϸ� �ȴ�.
    'Ư�� tu02�� ����.tp03�� ����.
    'th01�� ������ �ǳ�. ������� ���(���޸�)�� ��������.
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
'�ٸ� ����� ���
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

'����ȸ��������
'�����
'������
'aaa
'altnrwjs@ hanmail.net
'pquiz@ pquiz.codns.com
'pquiz2@ pquiz.codns.com
'pquiz3@ pquiz.codns.com
'���׽�Ʈ1
'ps8629@ naver.com

    Fn_SQLExec (sSql)

'----------------------------------------------------------------------------
'������������ ���
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
    If siz < 2 Then Exit Sub '0 �̳� 1�̸� ����������.
    
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
''samplet�ڵ�
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
