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
      Caption         =   "�ݱ�"
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
      Caption         =   "����"
      Height          =   855
      Left            =   1845
      TabIndex        =   5
      Top             =   6075
      Width           =   2505
   End
   Begin VB.Label Label12 
      Caption         =   "�� ��ȣ�� ��ȣȭ�� �Ұ����� ���·� �����ͺ��̽��� ����ǹǷ� �����ڵ� �� �� �����ϴ�."
      Height          =   555
      Left            =   900
      TabIndex        =   19
      Top             =   7290
      Width           =   7350
   End
   Begin VB.Label Label11 
      Caption         =   "���������� ��ȭ��ȣ�� �Է��մϴ�."
      Height          =   555
      Left            =   1845
      TabIndex        =   18
      Top             =   5355
      Width           =   5460
   End
   Begin VB.Label Label10 
      Caption         =   "�Ϲ���ȭ��ȣ"
      Height          =   465
      Left            =   630
      TabIndex        =   16
      Top             =   4935
      Width           =   1125
   End
   Begin VB.Label Label9 
      Caption         =   "������ȣ �нǷ� ���� ������ȣ ��߱޽� �̸��ϰ� �ش� �ڵ�����ȣ�ε� �ű� ������ȣ�� �߼��Ͽ� �ȳ��ص帳�ϴ�."
      Height          =   600
      Left            =   1800
      TabIndex        =   15
      Top             =   4185
      Width           =   6270
   End
   Begin VB.Label Label8 
      Caption         =   "��ȣ����ø��Է�"
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
      Caption         =   "�̸�(�Ǹ�)"
      Height          =   285
      Left            =   615
      TabIndex        =   12
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label6 
      Caption         =   "�ڵ�����ȣ"
      Height          =   465
      Left            =   630
      TabIndex        =   10
      Top             =   3780
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "����ȣȮ��"
      Height          =   435
      Left            =   630
      TabIndex        =   4
      Top             =   3210
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "����ȣ"
      Height          =   465
      Left            =   630
      TabIndex        =   3
      Top             =   2595
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "������ȣ"
      Height          =   345
      Left            =   630
      TabIndex        =   2
      Top             =   2115
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "�̸���"
      Height          =   375
      Left            =   630
      TabIndex        =   1
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "��ȸ������ ����"
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
'-���׸���� ����
'insert into ts01 values('�ѿ�����','�ѿ�����-����4751');/*�����߰�*/
'create table tq16 as select * from tq15 limit 0;
'load data infile 'tq16.txt' into table tq16 fields terminated by '\t';
'
'alter table tq16
'add(userid varchar(50) not null default '' comment '������ID');
'
'insert into vq01 select * from tq16;
'insert into ts02 values('�ѿ�����','����','20121228','21121228');/*����ڿ��� ���Ѻο�*/

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Con_Open
If txtPwd1.Text <> "" Or txtPwd2.Text <> "" Then
    If txtPwd2.Text <> txtPwd3.Text Then
        Con_Close
        MsgBox "�űԺ�й�ȣ�� ���� ��ġ���� �ʽ��ϴ�.", vbExclamation, "��������"
            
        Exit Sub
    ElseIf txtPwd2.Text = "" And InStr(gUserid, "@") > 0 Then
        Con_Close
        MsgBox "�űԺ�й�ȣ�� �߸��Ǿ����ϴ�.(���鸸���δ� �Ұ����մϴ�.)", vbExclamation, "��������"
        Exit Sub
    End If
    
    sSql = "update tu01 set username='" & txtName.Text & "',mobile='" & txtMobile.Text & "',phone='" & txtPhone.Text & "',userpass=md5('" & txtPwd2 & "') where userid='" & gUserid & "'"
    
    affected = Fn_SQLExec(sSql).nrow
    
    Con_Close
    
    If affected = 1 Then
        MsgBox "����Ǿ����ϴ�.", vbExclamation, "��������-ȸ����������"
    End If
    
    
Else
    sSql = "update tu01 set username='" & txtName.Text & "',mobile='" & txtMobile.Text & "',phone='" & txtPhone.Text & "' where userid='" & gUserid & "'"
    
    affected = Fn_SQLExec(sSql).nrow
    
    Con_Close
    
    If affected = 1 Then
        MsgBox "����Ǿ����ϴ�.", vbExclamation, "��������-ȸ����������"
    End If
    
End If


End Sub

Private Sub Form_Load()
'��ŷ���α׷� ���� ���� �� ���: ���� Tag �Ӽ��� ���� "FORM" �̾�� ��ŷ ���� ������ �� �ֽ��ϴ�.
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

