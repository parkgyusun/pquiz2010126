VERSION 5.00
Begin VB.Form frmRef 
   AutoRedraw      =   -1  'True
   Caption         =   "참고도"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmRef.frx":0000
   LinkTopic       =   "frmRef"
   ScaleHeight     =   4095
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      Height          =   870
      Left            =   1980
      ScaleHeight     =   810
      ScaleWidth      =   900
      TabIndex        =   0
      Top             =   810
      Width           =   960
   End
   Begin VB.Image img 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   3615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5460
   End
End
Attribute VB_Name = "frmRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub Form_Resize()
img.Width = Me.Width - Screen.TwipsPerPixelX * 8 '- Screen.TwipsPerPixelX * 40
img.Height = Me.Height - Screen.TwipsPerPixelY * 28 '- Screen.TwipsPerPixelY * 40

End Sub
