VERSION 5.00
Begin VB.Form frmUserUpdate 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "FORM"
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   855
      Left            =   4680
      TabIndex        =   20
      Top             =   6075
      Width           =   2505
   End
   Begin VB.TextBox txtPhone 
      Height          =   345
      Left            =   1770
      TabIndex        =   17
      Top             =   4905
      Width           =   3075
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   1755
      TabIndex        =   13
      Top             =   1050
      Width           =   3075
   End
   Begin VB.TextBox txtMobile 
      Height          =   345
      Left            =   1770
      TabIndex        =   11
      Top             =   3750
      Width           =   3075
   End
   Begin VB.TextBox txtPwd3 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3210
      Width           =   3075
   End
   Begin VB.TextBox txtPwd2 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2655
      Width           =   3075
   End
   Begin VB.TextBox txtPwd1 
      BackColor       =   &H00C0C0FF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2085
      Width           =   3075
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   345
      Left            =   1770
      TabIndex        =   6
      Top             =   1530
      Width           =   3075
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      Height          =   855
      Left            =   1845
      TabIndex        =   5
      Top             =   6075
      Width           =   2505
   End
   Begin VB.Label Label12 
      Caption         =   "※ 암호는 복호화가 불가능한 상태로 데이터베이스에 저장되므로 관리자도 알 수 없습니다."
      Height          =   555
      Left            =   900
      TabIndex        =   19
      Top             =   7290
      Width           =   7350
   End
   Begin VB.Label Label11 
      Caption         =   "연락가능한 전화번호를 입력합니다."
      Height          =   555
      Left            =   1845
      TabIndex        =   18
      Top             =   5355
      Width           =   5460
   End
   Begin VB.Label Label10 
      Caption         =   "일반전화번호"
      Height          =   465
      Left            =   630
      TabIndex        =   16
      Top             =   4935
      Width           =   1125
   End
   Begin VB.Label Label9 
      Caption         =   "인증번호 분실로 인한 인증번호 재발급시 이메일과 해당 핸드폰번호로도 신규 인증번호를 발송하여 안내해드립니다."
      Height          =   600
      Left            =   1800
      TabIndex        =   15
      Top             =   4185
      Width           =   6270
   End
   Begin VB.Label Label8 
      Caption         =   "암호변경시만입력"
      Height          =   870
      Left            =   5895
      TabIndex        =   14
      Top             =   2745
      Width           =   2265
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   5805
      Y1              =   2295
      Y2              =   2970
   End
   Begin VB.Line Line1 
      X1              =   4995
      X2              =   5805
      Y1              =   3510
      Y2              =   2970
   End
   Begin VB.Label Label7 
      Caption         =   "이름(실명)"
      Height          =   285
      Left            =   615
      TabIndex        =   12
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label6 
      Caption         =   "핸드폰번호"
      Height          =   465
      Left            =   630
      TabIndex        =   10
      Top             =   3780
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "새암호확인"
      Height          =   435
      Left            =   630
      TabIndex        =   4
      Top             =   3210
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "새암호"
      Height          =   465
      Left            =   630
      TabIndex        =   3
      Top             =   2595
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "기존암호"
      Height          =   345
      Left            =   630
      TabIndex        =   2
      Top             =   2115
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "이메일"
      Height          =   375
      Left            =   630
      TabIndex        =   1
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "●회원정보 변경"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   390
      TabIndex        =   0
      Top             =   360
      Width           =   5850
   End
End
Attribute VB_Name = "frmUserUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-문항만들기 절차
'insert into ts01 values('한영퀴즈','한영퀴즈-종합4751');/*과목추가*/
'create table tq16 as select * from tq15 limit 0;
'load data infile 'tq16.txt' into table tq16 fields terminated by '\t';
'
'alter table tq16
'add(userid varchar(50) not null default '' comment '변경자ID');
'
'insert into vq01 select * from tq16;
'insert into ts02 values('한영퀴즈','영어','20121228','21121228');/*사용자에게 권한부여*/

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Con_Open
If txtPwd1.Text <> "" Or txtPwd2.Text <> "" Then
    If txtPwd2.Text <> txtPwd3.Text Then
        Con_Close
        MsgBox "신규비밀번호가 서로 일치하지 않습니다.", vbExclamation, "포켓퀴즈"
            
        Exit Sub
    ElseIf txtPwd2.Text = "" And InStr(gUserid, "@") > 0 Then
        Con_Close
        MsgBox "신규비밀번호가 잘못되었습니다.(공백만으로는 불가능합니다.)", vbExclamation, "포켓퀴즈"
        Exit Sub
    End If
    
    sSql = "update tu01 set username='" & txtName.Text & "',mobile='" & txtMobile.Text & "',phone='" & txtPhone.Text & "',userpass=md5('" & txtPwd2 & "') where userid='" & gUserid & "'"
    
    affected = Fn_SQLExec(sSql).nrow
    
    Con_Close
    
    If affected = 1 Then
        MsgBox "저장되었습니다.", vbExclamation, "포켓퀴즈-회원정보변경"
    End If
    
    
Else
    sSql = "update tu01 set username='" & txtName.Text & "',mobile='" & txtMobile.Text & "',phone='" & txtPhone.Text & "' where userid='" & gUserid & "'"
    
    affected = Fn_SQLExec(sSql).nrow
    
    Con_Close
    
    If affected = 1 Then
        MsgBox "저장되었습니다.", vbExclamation, "포켓퀴즈-회원정보변경"
    End If
    
End If


End Sub

Private Sub Form_Load()
'도킹프로그램 제작 절차 및 방법: 폼의 Tag 속성의 값이 "FORM" 이어야 도킹 폼을 제작할 수 있습니다.
Con_Open
If gPassWord = "" Then
    Label3.Visible = False
    txtPwd1.Visible = False
    Line1.Visible = False
    Line2.Visible = False
    Label8.Visible = False
Else
    Label3.Visible = True
    txtPwd1.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    Label8.Visible = True
End If
sSql = "select * from tu01 where userid='" + gUserid + "'"
    Set rs = Fn_SQLExec(sSql).rs
    If Not rs.EOF Then
        txtID = gUserid
        txtName = rs("username")
        txtMobile = rs("mobile")
        txtPhone = rs("phone")
    End If

Con_Close

End Sub

