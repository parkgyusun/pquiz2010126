VERSION 5.00
Begin VB.Form frmVq01 
   Caption         =   "문항수정"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   Icon            =   "frmVq01.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5100
      TabIndex        =   30
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "<"
      Height          =   375
      Left            =   4740
      TabIndex        =   29
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtUpdatechasu 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   270
      Left            =   3780
      TabIndex        =   28
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton cmdTH01 
      Caption         =   "신고내용보기(&V)"
      Height          =   435
      Left            =   5940
      TabIndex        =   26
      Top             =   540
      Width           =   2475
   End
   Begin VB.TextBox txtHint 
      Height          =   1170
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   25
      Top             =   8280
      Width           =   7275
   End
   Begin VB.TextBox txtE 
      Height          =   1170
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   24
      Top             =   7080
      Width           =   7275
   End
   Begin VB.TextBox txtD 
      Height          =   1170
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   23
      Top             =   5880
      Width           =   7275
   End
   Begin VB.TextBox txtC 
      Height          =   1170
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   4680
      Width           =   7275
   End
   Begin VB.TextBox txtB 
      Height          =   1170
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      Top             =   3480
      Width           =   7275
   End
   Begin VB.TextBox txtA 
      Height          =   1170
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Top             =   2280
      Width           =   7275
   End
   Begin VB.TextBox txtQuiz 
      Height          =   1170
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Top             =   1080
      Width           =   7275
   End
   Begin VB.TextBox txtAns 
      Height          =   270
      Left            =   3780
      TabIndex        =   18
      Top             =   780
      Width           =   1695
   End
   Begin VB.TextBox txtCat 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      TabIndex        =   17
      Top             =   780
      Width           =   1515
   End
   Begin VB.TextBox txtMode 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   270
      Left            =   3780
      TabIndex        =   16
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtSeq 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      TabIndex        =   15
      Top             =   480
      Width           =   1515
   End
   Begin VB.TextBox txtSubj 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      TabIndex        =   14
      Top             =   180
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "닫기(&C)"
      Height          =   435
      Left            =   7200
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장(&S)"
      Height          =   435
      Left            =   5940
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "차수"
      Height          =   255
      Index           =   12
      Left            =   2940
      TabIndex        =   27
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "모드"
      Height          =   255
      Index           =   11
      Left            =   2940
      TabIndex        =   13
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "정답"
      Height          =   255
      Index           =   10
      Left            =   2940
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "힌트"
      Height          =   255
      Index           =   9
      Left            =   180
      TabIndex        =   11
      Top             =   8340
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "보기5"
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   10
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "보기4"
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   9
      Top             =   5940
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "보기3"
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   8
      Top             =   4740
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "보기2"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "보기1"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   2340
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "퀴즈"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "분류"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "일련번호"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "과목명"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmVq01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public parent As frmQuiz
Public saved As Boolean
Private maxChasu As Long
Private savedChasu As Long

Public Sub fillAll(Optional chasu As Long)
Con_Open
'
'CREATE TABLE `vq01_log` (
'   txt varchar(255) default NULL ,??
'  `subj` varchar(255) default NULL,
'  `seq` int(11) default NULL,
'  `cat` varchar(255) default NULL,
'  `quiz` mediumtext,
'  `a` mediumtext,
'  `b` mediumtext,
'  `c` mediumtext,
'  `d` mediumtext,
'  `e` mediumtext,
'  `hint` mediumtext,
'  `ans` varchar(255) default NULL,
'  `resid` varchar(255) default NULL,
'  `mode` varchar(255) default NULL,
'  `updateymd` varchar(255) default NULL,
'  `updatechasu` int(11) default NULL,
'  UNIQUE KEY `subjuniq` (`subj`,`seq`),
'  KEY `subj_2` (`subj`)
') ENGINE=MyISAM DEFAULT CHARSET=latin1;

If IsMissing(chasu) Or chasu = 0 Or chasu = maxChasu Then
   sSql = "select * from vq01 where subj='" + txtSubj + "' and seq=" + txtSeq + ""
Else
   sSql = "select * from vq01_log where subj='" + txtSubj + "' and seq=" + txtSeq + " and updatechasu=" + CStr(chasu)
End If

Set rs = Fn_SQLExec(sSql, , , True).rs
Dim sHint As Variant
Dim vA As Variant, vB, vC, vD, vE

If Not rs.EOF Then
    txtCat.Text = rs("cat")
    txtQuiz.Text = rs("quiz")
    txtAns.Text = rs("ans")
    txtMode.Text = rs("mode")
    
'    txtA.Text = rs("a")
    
    vA = rs("a")
    If (IsNull(vA)) Then
        vA = ""
    End If
    
    
    vB = rs("b")
    If (IsNull(vB)) Then
        vB = ""
    End If
    
    vC = rs("c")
    If (IsNull(vC)) Then
        vC = ""
    End If
    
    vD = rs("d")
    If (IsNull(vD)) Then
        vD = ""
    End If
    
    vE = rs("e")
    If (IsNull(vE)) Then
        vE = ""
    End If
    
    txtA.Text = vA
    txtB.Text = vB
    txtC.Text = vC
    txtD.Text = vD
    txtE.Text = vE
    
    sHint = rs("hint")
    If IsNull(sHint) Then
        txtHint.Text = ""
    Else
        txtHint.Text = sHint
    End If
    
    txtUpdatechasu = rs("updatechasu")
    If txtUpdatechasu.Text = "1" Then
        cmdPre.Enabled = False
    End If
        
    If IsMissing(chasu) Or chasu = 0 Then
        maxChasu = CLng(txtUpdatechasu.Text)
        chasu = maxChasu
    End If
    
    If chasu = maxChasu Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    
    If chasu = 1 Then
        cmdPre.Enabled = False
    Else
        cmdPre.Enabled = True
    End If
    
    savedChasu = chasu
    
    txtUpdatechasu.Text = savedChasu & "/" & maxChasu
    
'    txtHint.Text = IIf(IsNull(rs("hint")), "", rs("hint"))
End If


Con_Close

End Sub

Private Sub cmdNext_Click()
    Call fillAll(savedChasu + 1)
End Sub

Private Sub cmdPre_Click()
    Call fillAll(savedChasu - 1)
End Sub

Private Sub cmdSave_Click()
Dim affected As Integer
Con_Open

sSql = "select 1 from tu01 where userid='" & gUserid & "' and usergrade=99"
If Fn_SQLExec(sSql).nrow > 0 Then
    gbIsSuperAdmin = True
Else
    gbIsSuperAdmin = False
End If

If gbIsSuperAdmin Then
    sSql = "insert into vq01_log(txt,subj,seq,cat,quiz,a,B,C,d,e,hint,ans,resid,mode,updateymd,updatechasu,userid) select null,subj,seq,cat,quiz,a,b,c,d,e,hint,ans,resid,mode,updateymd,updatechasu,userid "
    sSql = sSql + " from vq01 where subj='" + txtSubj.Text + " ' and seq=" + txtSeq.Text + ""
    
    affected = Fn_SQLExec(sSql).nrow
    Debug.Assert affected = 1
    
    sSql = "update vq01 set "
    sSql = sSql & " cat= '" & txtCat.Text & "' ,"
    sSql = sSql & " quiz= '" & MakeDbTxt(txtQuiz.Text) & "' ,"
    sSql = sSql & " a= '" & MakeDbTxt(txtA.Text) & "' ,"
    sSql = sSql & " b= '" & MakeDbTxt(txtB.Text) & "' ,"
    sSql = sSql & " c= '" & MakeDbTxt(txtC.Text) & "' ,"
    sSql = sSql & " d= '" & MakeDbTxt(txtD.Text) & "' ,"
    sSql = sSql & " e= '" & MakeDbTxt(txtE.Text) & "' ,"
    sSql = sSql & " hint= '" & MakeDbTxt(txtHint.Text) & "' ,"
    sSql = sSql & " ans= '" & txtAns.Text & "' ,"
    sSql = sSql & " mode= '" & txtMode.Text & "' ,"
    sSql = sSql & " updateymd=date_format(now(), '%Y%m%d' ),"
    sSql = sSql & " updatechasu=updatechasu+1,"
    sSql = sSql & " userid='" & gUserid & "'"
    sSql = sSql & " where subj='" & MakeDbTxt(txtSubj.Text) & "'"
    sSql = sSql & " and seq=" & txtSeq.Text & ""
    
    affected = Fn_SQLExec(sSql).nrow

End If

Debug.Assert affected = 1

Con_Close
saved = True

Me.Hide

End Sub

Private Sub cmdTH01_Click()
Dim formTH01 As New frmTH01

Load formTH01

formTH01.sSubj = txtSubj.Text
formTH01.sSeq = txtSeq.Text

formTH01.fillAll
TmrAfterTTS_exit = True
formTH01.Show vbModal
TmrAfterTTS_exit = False

Unload formTH01
cmdSave.Enabled = True
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()

'계속같은게 나오면 아래것에서 uerid='@' 인 것을 지우면 된다.
'mysql> select subj,seq,count(*)  from th01 where userid like '@%' group by subj,
'seq having count(*)>1;

End Sub
