VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnecting 
   BorderStyle     =   1  '단일 고정
   Caption         =   "포켓퀴즈"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4830
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   4740
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   240
         Left            =   225
         TabIndex        =   2
         Top             =   495
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
         Max             =   6000
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   4185
         Top             =   0
      End
      Begin VB.Label Label1 
         Caption         =   "연결을 시도합니다.."
         Height          =   240
         Left            =   1575
         TabIndex        =   1
         Top             =   225
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmConnecting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0


Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Timer1_Timer()

If ProgressBar1.Value = ProgressBar1.Max Then
    ProgressBar1.Value = 0
End If

ProgressBar1.Value = ProgressBar1.Value + 100

End Sub
