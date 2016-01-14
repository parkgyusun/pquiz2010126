VERSION 5.00
Begin VB.Form frmTH01EE 
   Caption         =   "이의신청"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frmTH01EE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture1 
      Align           =   2  '아래 맞춤
      Appearance      =   0  '평면
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6360
      TabIndex        =   1
      Top             =   3990
      Width           =   6360
      Begin VB.CommandButton cmdSave 
         Caption         =   "이의신청(&S)"
         Height          =   435
         Left            =   1740
         TabIndex        =   3
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "닫기(&C)"
         Height          =   435
         Left            =   3180
         TabIndex        =   2
         Top             =   60
         Width           =   1335
      End
   End
   Begin VB.TextBox rtxt1 
      Height          =   1770
      IMEMode         =   10  '한글 
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   60
      Width           =   1545
   End
End
Attribute VB_Name = "frmTH01EE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sSubj As String, sSeq As String

Private Sub cmdSave_Click()
 Dim affected As Long
    
    'If rtxt1.Text = mvarMemo Then Exit Sub
    
    sSql = "update th01 set userid='@', hint='" & MakeDbTxt(rtxt1.Text) & "' where userid like '@%' and subj = '" & sSubj & "' and seq=" & sSeq & ""
    
    affected = Fn_SQLExec(sSql, , True).nrow
    
    If affected = 0 Then
        
        'refCquiz.memo = rtxt1.Text
        
        sSql = "insert into th01(userid,subj,seq,hint,bshare) values('@','" & sSubj & "'," & sSeq & ",'" & MakeDbTxt(rtxt1.Text) & "','1')"
        
        affected = Fn_SQLExec(sSql).nrow
        
        Debug.Assert affected > 0
        
    End If
 Me.Hide
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Resize()

Call rtxt1.Move(0, 0, Me.Width - 110, Me.Height - 850)

End Sub

Public Sub fillAll()
Con_Open

sSql = "select * from th01 where subj='" & sSubj & "' and seq=" & sSeq & " and userid like '@%' order by userid desc"

Set rs = Fn_SQLExec(sSql, , , True).rs

Dim sHint As Variant

'이부분 while 문으로 하면 안되네~ㅇㅇㅇ
If Not rs.EOF Then
    
    sHint = rs("hint")
    If IsNull(sHint) Then
        rtxt1.Text = "[신청자:" & gUserid & "]" + vbCrLf
    Else
        rtxt1.Text = sHint
    End If
    
    'sUserid = rs("userid")
    
'    txtUserID.Text = sUserid
Else
    rtxt1.Text = "[신청자:" & gUserid & "]" + vbCrLf
End If

Con_Close

End Sub
