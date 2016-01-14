VERSION 5.00
Begin VB.Form frmLogin2 
   Caption         =   "암호 확인"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1335
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   360
      Left            =   525
      TabIndex        =   1
      Tag             =   "1048"
      Top             =   975
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   360
      Left            =   2130
      TabIndex        =   0
      Tag             =   "1049"
      Top             =   975
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "새로운 사용자로 등록 합니다. "
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   3570
   End
   Begin VB.Label lblLabels 
      Caption         =   "암호 재입력:"
      Height          =   255
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Tag             =   "1047"
      Top             =   495
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
