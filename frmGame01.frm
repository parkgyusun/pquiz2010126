VERSION 5.00
Begin VB.Form frmGame01 
   BorderStyle     =   0  '없음
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11445
   ScaleWidth      =   17955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Tag             =   "FORM"
   Begin VB.PictureBox picMain 
      Appearance      =   0  '평면
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   10185
      Left            =   450
      Picture         =   "frmGame01.frx":0000
      ScaleHeight     =   10155
      ScaleWidth      =   17355
      TabIndex        =   0
      Top             =   270
      Width           =   17385
      Begin VB.PictureBox picQuizWait 
         Appearance      =   0  '평면
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   3195
         TabIndex        =   6
         Top             =   -60
         Width           =   3225
      End
      Begin VB.PictureBox picQuiz 
         Appearance      =   0  '평면
         BackColor       =   &H00000040&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   690
         ScaleHeight     =   465
         ScaleWidth      =   3165
         TabIndex        =   5
         Top             =   420
         Width           =   3195
      End
      Begin VB.PictureBox picD1 
         Appearance      =   0  '평면
         BackColor       =   &H00112024&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   4380
         ScaleHeight     =   495
         ScaleWidth      =   12945
         TabIndex        =   4
         Top             =   5010
         Width           =   12975
      End
      Begin VB.PictureBox picC1 
         Appearance      =   0  '평면
         BackColor       =   &H00111D24&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   3690
         ScaleHeight     =   495
         ScaleWidth      =   13635
         TabIndex        =   3
         Top             =   5520
         Width           =   13665
      End
      Begin VB.PictureBox picB1 
         Appearance      =   0  '평면
         BackColor       =   &H00111B24&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   3000
         ScaleHeight     =   495
         ScaleWidth      =   14325
         TabIndex        =   2
         Top             =   6030
         Width           =   14355
      End
      Begin VB.PictureBox picA1 
         Appearance      =   0  '평면
         BackColor       =   &H00111824&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   2310
         ScaleHeight     =   495
         ScaleWidth      =   15015
         TabIndex        =   1
         Top             =   6540
         Width           =   15045
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "보유지식: 11591311313"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Top             =   8370
         Width           =   3555
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "습득지식: 3"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   8130
         Width           =   3465
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgGalaxy 
         Height          =   435
         Left            =   4440
         Picture         =   "frmGame01.frx":1FFE42
         Stretch         =   -1  'True
         Top             =   4530
         Width           =   555
      End
      Begin VB.Image imgKTX4 
         Height          =   420
         Left            =   3690
         Picture         =   "frmGame01.frx":203743
         Stretch         =   -1  'True
         Top             =   5070
         Width           =   705
      End
      Begin VB.Image imgKTX3 
         Height          =   420
         Left            =   3000
         Picture         =   "frmGame01.frx":20457C
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   705
      End
      Begin VB.Image imgKTX2 
         Height          =   420
         Left            =   2310
         Picture         =   "frmGame01.frx":2053B5
         Stretch         =   -1  'True
         Top             =   6060
         Width           =   705
      End
      Begin VB.Image imgKTX1 
         Height          =   420
         Left            =   1620
         Picture         =   "frmGame01.frx":2061EE
         Stretch         =   -1  'True
         Top             =   6570
         Width           =   705
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "남은시간: 3초"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   7890
         Width           =   3045
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경과시간: 3초"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   7650
         Width           =   3315
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "기회: 3번"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   7410
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   9
         X1              =   900
         X2              =   1020
         Y1              =   6960
         Y2              =   7380
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   7
         X1              =   1590
         X2              =   1080
         Y1              =   6990
         Y2              =   7380
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   6
         X1              =   210
         X2              =   900
         Y1              =   6960
         Y2              =   7380
      End
      Begin VB.Image imgCan 
         Height          =   720
         Left            =   1740
         Picture         =   "frmGame01.frx":207027
         Stretch         =   -1  'True
         Top             =   2250
         Width           =   480
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   5
         X1              =   210
         X2              =   210
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   4
         X1              =   900
         X2              =   900
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   3
         X1              =   4350
         X2              =   4350
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   2
         X1              =   3660
         X2              =   3660
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   1
         X1              =   2970
         X2              =   2970
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   0
         X1              =   2280
         X2              =   2280
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line lnGrdH 
         BorderColor     =   &H0000FF00&
         Index           =   13
         X1              =   1590
         X2              =   1590
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   12
         X1              =   0
         X2              =   5040
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   11
         X1              =   0
         X2              =   5040
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   10
         X1              =   0
         X2              =   5040
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   9
         X1              =   0
         X2              =   5040
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   8
         X1              =   0
         X2              =   5040
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   7
         X1              =   0
         X2              =   5040
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   6
         X1              =   0
         X2              =   5040
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   5
         X1              =   0
         X2              =   5040
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Image imgWall 
         Height          =   5040
         Left            =   5070
         Picture         =   "frmGame01.frx":207BDB
         Top             =   -30
         Width           =   420
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   4
         X1              =   0
         X2              =   5040
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   3
         X1              =   0
         X2              =   4350
         Y1              =   5010
         Y2              =   5010
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   2
         X1              =   0
         X2              =   4350
         Y1              =   5490
         Y2              =   5490
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   1
         X1              =   0
         X2              =   4350
         Y1              =   6510
         Y2              =   6510
      End
      Begin VB.Line lnGrdV 
         BorderColor     =   &H0000FF00&
         Index           =   0
         X1              =   0
         X2              =   4350
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Image imgEarth 
         Height          =   1425
         Left            =   10290
         Picture         =   "frmGame01.frx":20EA5D
         Stretch         =   -1  'True
         Top             =   -1290
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Image imgSpacecraft 
         Height          =   5910
         Left            =   5370
         Picture         =   "frmGame01.frx":2115E2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8835
      End
      Begin VB.Image imgBlackhole 
         Height          =   2040
         Left            =   -150
         Picture         =   "frmGame01.frx":2BAC9C
         Stretch         =   -1  'True
         Top             =   6510
         Width           =   2280
      End
      Begin VB.Image imgWhiteHole 
         Height          =   405
         Left            =   1650
         Picture         =   "frmGame01.frx":35C49E
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   600
      End
      Begin VB.Image imgWhiteHoleWait 
         Height          =   405
         Left            =   960
         Picture         =   "frmGame01.frx":36BA65
         Stretch         =   -1  'True
         Top             =   2010
         Width           =   600
      End
      Begin VB.Image imgAdballoon 
         Height          =   2925
         Left            =   750
         Picture         =   "frmGame01.frx":37B02C
         Stretch         =   -1  'True
         Top             =   180
         Width           =   2265
      End
      Begin VB.Image imgAdballoonWait 
         Height          =   2925
         Left            =   120
         Picture         =   "frmGame01.frx":37D41B
         Stretch         =   -1  'True
         Top             =   -150
         Width           =   2265
      End
      Begin VB.Image imgCanWait 
         Height          =   720
         Left            =   1050
         Picture         =   "frmGame01.frx":37F80A
         Stretch         =   -1  'True
         Top             =   1710
         Width           =   480
      End
      Begin VB.Label lbNews 
         BackColor       =   &H00000000&
         BackStyle       =   0  '투명
         Caption         =   "instead = ~대신에"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1395
         Left            =   2160
         TabIndex        =   12
         Top             =   7050
         Width           =   3255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   6510
   End
End
Attribute VB_Name = "frmGame01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public nowVisible As Boolean '보여지고있는동안에 이벤트가 걸리면 빠져나가기 위함이다.

Public frmMain As frmMain

Dim SSQL1 As String

Public parent As Form

Public WithEvents objGame01 As clsGame01
Attribute objGame01.VB_VarHelpID = -1


Private Sub Form_Activate()
Dim hIMC As Long
hIMC = ImmGetContext(Me.hwnd)
'Call ImmSetConversionStatus(hIMC, IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE)
Call ImmSetConversionStatus(hIMC, IME_CMODE_FIXED, IME_SMODE_NONE)
End Sub

Private Sub Form_Load()
Set objGame01 = New clsGame01

End Sub

Private Sub Form_Resize()

Debug.Assert Me.Tag = "FORM" '윈도우 만들기 위해 태그를 FORM으로 해야 메인과 같이 적용된 크기로 변경된다.

    Dim grdHeight As Single
    Dim mchartWidth As Single
    
    Dim fMargin As Single
    
    fMargin = 10
    grdHeight = Me.Height - picMain.Top * 2 - fMargin
    If grdHeight < 10 Then
        picMain.Height = 10
    Else
        picMain.Height = grdHeight
    End If
    
    mchartWidth = Me.Width - fMargin - picMain.Left * 2
    
    If fMargin < 10 Then
        picMain.Width = 10
    Else
        picMain.Width = mchartWidth
    End If

End Sub

