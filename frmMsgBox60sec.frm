VERSION 5.00
Begin VB.Form frmMsgBox60sec 
   AutoRedraw      =   -1  'True
   Caption         =   "메시지박스"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4275
   HasDC           =   0   'False
   Icon            =   "frmMsgBox60sec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4275
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "취소(&N)"
      Height          =   345
      Index           =   1
      Left            =   2190
      TabIndex        =   2
      Top             =   1110
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   900
      TabIndex        =   1
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   270
      Top             =   990
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   330
      Picture         =   "frmMsgBox60sec.frx":000C
      Top             =   330
      Width           =   465
   End
   Begin VB.Label Label1 
      Height          =   555
      Left            =   960
      TabIndex        =   0
      Top             =   300
      Width           =   3195
   End
End
Attribute VB_Name = "frmMsgBox60sec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_msg As String
Public cnt As Integer
Public btn_flag As Integer
Public btn_type As VbMsgBoxStyle

Private Sub Command1_Click(Index As Integer)
    btn_flag = Index
    Command1(Index).Default = True
    cnt = 0
End Sub

Private Sub Command1_GotFocus(Index As Integer)
    Command1(Index).Default = True
End Sub

Private Sub Form_Activate()
    Select Case btn_type
    Case vbOKOnly
        Command1(1).Visible = False
        Command1(0).Left = 1950
    Case vbYesNo
        Command1(1).Visible = True
        Command1(0).Left = 900
    Case Else
        MsgBox "미구현 버튼 상태", vbCritical
    End Select
End Sub

Private Sub Form_Load()
    btn_type = vbYesNo
End Sub

Private Sub Timer1_Timer()
    cnt = cnt - 1
    If cnt <= 0 Then
        If Command1(0).Default Then
            btn_flag = 0
        Else
            btn_flag = 1
        End If
        Unload Me
    End If
    Call localSetMsg
End Sub

Public Sub setMsg(msg As String)
    cnt = 60
    
    m_msg = msg
    Timer1.Enabled = True
    localSetMsg
End Sub

Private Sub localSetMsg()
    Label1.Caption = m_msg
    Me.Caption = CStr(cnt)
End Sub
