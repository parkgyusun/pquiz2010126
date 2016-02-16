VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmQuiz 
   BorderStyle     =   0  'None
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12690
   ControlBox      =   0   'False
   Icon            =   "frmQuiz.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   846
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtJindo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      HideSelection   =   0   'False
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   5175
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pgBarJindo 
      Height          =   195
      Left            =   -45
      TabIndex        =   34
      Top             =   5760
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdImgHint 
      Height          =   510
      Left            =   10755
      Picture         =   "frmQuiz.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "그림힌트"
      Top             =   45
      Width           =   510
   End
   Begin VB.Timer TmrAfterTTS_focus 
      Enabled         =   0   'False
      Left            =   90
      Top             =   3330
   End
   Begin VB.CommandButton cmdPen 
      BackColor       =   &H0000FF00&
      Height          =   405
      Left            =   11790
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6540
      Top             =   6330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   37
      ImageHeight     =   37
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuiz.frx":08CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuiz.frx":194E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuiz.frx":29D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuiz.frx":3A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuiz.frx":4AD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   435
      Left            =   11340
      TabIndex        =   22
      Top             =   90
      Width           =   405
   End
   Begin VB.CommandButton cmdEE 
      Caption         =   "이의신청"
      Height          =   240
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   300
      Width           =   975
   End
   Begin VB.OptionButton optAutoClickOff 
      Caption         =   "모션클릭Off"
      Height          =   420
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   90
      Width           =   885
   End
   Begin VB.OptionButton optAutoClickOn 
      Caption         =   "모션클릭On"
      Height          =   420
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   90
      Width           =   1020
   End
   Begin VB.CommandButton cmdMemo 
      Caption         =   "내메모(&M)"
      Height          =   240
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   90
      Width           =   1005
   End
   Begin VB.Timer tmrTG 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   315
      Top             =   5580
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "삭제(&D)"
      Height          =   420
      Left            =   4245
      TabIndex        =   7
      Top             =   90
      Width           =   1005
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "이전(&B)"
      Height          =   420
      Left            =   2085
      TabIndex        =   6
      Top             =   90
      Width           =   1005
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   45
      Top             =   5670
   End
   Begin VB.PictureBox PicMain 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   180
      ScaleHeight     =   5535
      ScaleWidth      =   12240
      TabIndex        =   5
      Top             =   720
      Width           =   12300
      Begin VB.CommandButton optTTSE 
         Height          =   285
         Left            =   45
         Picture         =   "frmQuiz.frx":5B56
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   330
      End
      Begin VB.CommandButton optTTSD 
         Height          =   285
         Left            =   45
         Picture         =   "frmQuiz.frx":607B
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1890
         Width           =   330
      End
      Begin VB.CommandButton optTTSC 
         Height          =   285
         Left            =   45
         Picture         =   "frmQuiz.frx":65A0
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1575
         Width           =   330
      End
      Begin VB.CommandButton optTTSB 
         Height          =   285
         Left            =   90
         Picture         =   "frmQuiz.frx":6AC5
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1305
         Width           =   330
      End
      Begin VB.CommandButton optTTSA 
         Height          =   285
         Left            =   135
         Picture         =   "frmQuiz.frx":6FEA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1080
         Width           =   330
      End
      Begin VB.CommandButton optTTSQuiz 
         Height          =   285
         Left            =   135
         Picture         =   "frmQuiz.frx":750F
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   810
         Width           =   330
      End
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         Height          =   4875
         Left            =   3480
         ScaleHeight     =   4815
         ScaleWidth      =   5235
         TabIndex        =   20
         Top             =   300
         Width           =   5295
         Begin VB.Image imgE0 
            Height          =   420
            Left            =   0
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Image imgD0 
            Height          =   420
            Left            =   60
            Top             =   1920
            Width           =   2130
         End
         Begin VB.Image imgC0 
            Height          =   420
            Left            =   300
            Top             =   1380
            Width           =   2040
         End
         Begin VB.Image imgB0 
            Height          =   375
            Left            =   180
            Top             =   660
            Width           =   1950
         End
         Begin VB.Image imgA0 
            Height          =   330
            Left            =   1140
            Top             =   0
            Width           =   1995
         End
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "참고도(&R)"
         Height          =   420
         Left            =   480
         TabIndex        =   19
         Top             =   4860
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdHint 
         Caption         =   "힌트(&H)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   300
         TabIndex        =   18
         Top             =   4440
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.OptionButton optE 
         Alignment       =   1  'Right Justify
         Caption         =   "E)"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3840
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.OptionButton optD 
         Alignment       =   1  'Right Justify
         Caption         =   "D)"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton optC 
         Alignment       =   1  'Right Justify
         Caption         =   "C)"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3060
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton optB 
         Alignment       =   1  'Right Justify
         Caption         =   "B)"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2700
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton optA 
         Alignment       =   1  'Right Justify
         Caption         =   "A)"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   555
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   465
         Left            =   225
         TabIndex        =   31
         Top             =   4995
         Width           =   5775
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   -1  'True
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   10186
         _cy             =   820
      End
      Begin VB.Image imgE1 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   720
         Top             =   4080
         Width           =   7260
      End
      Begin VB.Image imgD1 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   660
         Top             =   3420
         Width           =   7260
      End
      Begin VB.Image imgC1 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   600
         Top             =   2100
         Width           =   7260
      End
      Begin VB.Shape shpSel 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Height          =   690
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.Image imgB1 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   540
         Top             =   660
         Width           =   7260
      End
      Begin VB.Image imgA1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         Top             =   120
         Width           =   7260
      End
   End
   Begin VB.TextBox txtTimer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   8  'SBCS ALPHABET
      Left            =   8205
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   90
      Width           =   465
   End
   Begin VB.CommandButton cmdTimer 
      Caption         =   "▶||"
      Height          =   420
      Left            =   7485
      TabIndex        =   3
      Top             =   90
      Width           =   645
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "다음(&N)"
      Default         =   -1  'True
      Height          =   420
      Left            =   3165
      TabIndex        =   2
      Top             =   90
      Width           =   1005
   End
   Begin MSComctlLib.ProgressBar pgBar 
      Height          =   150
      Left            =   225
      TabIndex        =   1
      Top             =   540
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
      Max             =   60
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "닫기(&X)"
      Height          =   420
      Left            =   5325
      TabIndex        =   0
      ToolTipText     =   "닫기"
      Top             =   90
      Width           =   1005
   End
   Begin VB.CheckBox chk_tts_dic 
      Caption         =   "영어학습"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   90
      Width           =   1095
   End
   Begin VB.CheckBox chk 
      Caption         =   "체크"
      Height          =   420
      Left            =   225
      TabIndex        =   9
      Top             =   90
      Width           =   870
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   645
      Left            =   0
      TabIndex        =   23
      Top             =   5940
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   1138
      ButtonWidth     =   1164
      ButtonHeight    =   1138
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Begin MSComDlg.CommonDialog dlgCMDialog 
         Left            =   10470
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin SHDocVwCtl.WebBrowser wbQuizMain 
      Height          =   5460
      Left            =   12510
      TabIndex        =   33
      Top             =   720
      Width           =   2535
      ExtentX         =   4471
      ExtentY         =   9631
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   12210
      Tag             =   "shape1의색상을정의해야함"
      Top             =   390
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   12210
      Tag             =   "&H00FF00FF&"
      Top             =   60
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   423
      X2              =   423
      Y1              =   2
      Y2              =   36
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   0
      X1              =   424
      X2              =   424
      Y1              =   2
      Y2              =   36
   End
   Begin VB.Menu mnuPop1 
      Caption         =   "mnuPop1"
      Begin VB.Menu mnuCopy 
         Caption         =   "복사(&C)"
      End
      Begin VB.Menu mnu_bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebPages 
         Caption         =   "웹페이지열기"
         Begin VB.Menu mnuNaverDic 
            Caption         =   "네이버사전(&N)"
         End
         Begin VB.Menu mnu_bar9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGugleTransPage 
            Caption         =   "구글번역페이지"
         End
         Begin VB.Menu mnuInternetEx 
            Caption         =   "인터넷익스플로러(&Z)"
         End
         Begin VB.Menu mnuChrome 
            Caption         =   "구글크롬브라우저"
         End
      End
      Begin VB.Menu mnu_bar10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoogleSearch 
         Caption         =   "구글검색"
      End
      Begin VB.Menu mnuNaverSearch 
         Caption         =   "네이버검색"
      End
      Begin VB.Menu mnuDaumSearch 
         Caption         =   "다음검색"
      End
      Begin VB.Menu mnu_naverModum_search 
         Caption         =   "네이버사전검색"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDoosan100 
         Caption         =   "두산백과사전"
      End
      Begin VB.Menu mnuWiki100 
         Caption         =   "위키백과사전"
      End
      Begin VB.Menu mnuZum100 
         Caption         =   "줌[zum]검색"
      End
      Begin VB.Menu mnuBritannica100 
         Caption         =   "브리태니커 사전"
      End
      Begin VB.Menu mnu_bar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageSearch 
         Caption         =   "이미지검색(구글)"
      End
      Begin VB.Menu mnuVideo 
         Caption         =   "동영상검색(구글)"
      End
      Begin VB.Menu mnuYouTubeSearch 
         Caption         =   "동영상검색(YouTube)"
      End
      Begin VB.Menu mnu_bar6 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnu_all_in_one 
         Caption         =   "영어사전1+TTS(&A)"
      End
      Begin VB.Menu mnuendic 
         Caption         =   "영어사전(&1)-naver"
      End
      Begin VB.Menu mnuEndic2 
         Caption         =   "영어사전(&2)-ybm"
      End
      Begin VB.Menu mnuEngStudyPekr 
         Caption         =   "영어학습사전(&3)-dic.impact.pe.kr"
      End
      Begin VB.Menu mnuEndic4 
         Caption         =   "영영사전(&4)[영국]-collins"
      End
      Begin VB.Menu mnu_ee_dic 
         Caption         =   "영영사전(&5)[미국]-collins"
      End
      Begin VB.Menu mnuWordListenENUS 
         Caption         =   "영단어발음[미국]듣기(&Q)"
      End
      Begin VB.Menu mnu_sepa7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_kr_dic 
         Caption         =   "국어사전(&K)"
      End
      Begin VB.Menu mnuHanjaDic 
         Caption         =   "한자사전(&5)"
      End
      Begin VB.Menu mnu_jp_dic 
         Caption         =   "일어사전(&J)"
      End
      Begin VB.Menu mnu_cn_dic 
         Caption         =   "중국어사전(&C)"
      End
      Begin VB.Menu mnu_bar7 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuBibleSearch 
         Caption         =   "성경검색(&b)"
      End
      Begin VB.Menu mnu_bar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit 
         Caption         =   "수정(&E)"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "새로고침(&R)"
      End
      Begin VB.Menu mnu_bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTTS0 
         Caption         =   "TTS실행(&X)"
      End
      Begin VB.Menu mnu_bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_auto_tts 
         Caption         =   "자동 TTS 설정"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_auto_dic 
         Caption         =   "학습사전설정"
      End
      Begin VB.Menu mnu_sepa8 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ref_see 
         Caption         =   "관련문제보기"
      End
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'기본시험지 변경 절차
'update tsdefault set usebit=2;
'insert into tsdefault values(null,'''영단어1'',''영속담1'',''영숙어 (중)''
',''영숙어1'',''중1단어'',''중2단어'',''중3단어'',''중학기초'',''일본어800''
',''특목고입시용'',''토익Voca'',''한영퀴즈'',
'''신약성경'',''구약성경'',''구약:01'',''구약:02'',''구약:03'',''구약:04''
',''구약:05'',''구약:06'',''구약:07'',''구약:08'',''구약:09'',''구약:10''
',''구약:11'',''구약:12'',''구약:13'',''구약:14'',''구약:15'',''구약:16''
',''구약:17'',''구약:18'',''구약:19'',''구약:20'',''구약:21'',''구약:22''
',''구약:23'',''구약:24'',''구약:25'',''구약:26
''',''구약:27'',''구약:28'',''구약:29'',''구약:30'',''구약:31'',''구약:32''
',''구약:33'',''구약:34'',''구약:35'',''구약:36'',''구약:37'',''구약:38''
',''구약:39'',''신약:40'',''신약:41'',''신약:42'',''신약:43'',''신약:44''
',''신약:45'',''신약:46'',''신약:47'',''신약:48'',''신약:49'',''신약:50''
',''신약:51'',''신약:52'',''신약:53'',''신약:54
''',''신약:55'',''신약:56'',''신약:57'',''신약:58'',''신약:59'',''신약:60''
',''신약:61'',''신약:62'',''신약:63'',''신약:64'',''신약:65'',''신약:66'',
'''한자'',''한자01'',''한자02'',''한자03'',''한자04'',''한자05'',''한자06''
',''한자07'',''한자08'',''한자09'',''한자10'',''한자11''
',''minus_10deg_'',''minus_11deg_'',''minus_12deg_'',''minus_13deg_''
',''minus_14deg_''
',''minus_15deg_'',''minus_16deg_'',''minus_17deg_'',''minus_18deg_''
',''minus_19deg_'',''minus_1char_'',''minus_20deg_'',''plus_Atype_''
',''plus_Btype_'',''plus_Ctype_'',''plus_Dtype_'',''plus_Etype_''
',''plus_Ftype_'',''plus_Gtype_'',''plus_Htype_''',1);
'

Option Explicit
Option Base 0

Private wbQuizFocused As Boolean

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Public nowVisible As Boolean '보여지고있는동안에 이벤트가 걸리면 빠져나가기 위함이다.

Public frmMain As frmMain
Public lfrmMemo As frmMemo

Public mGetRndName As String
Public gSetLang As String '한:ko-KR, 중: zh-CN, 영: en-US 일:ja-JP, 자동:(공백)

Public TotalQuizCount As Long

Dim lastUpButtonLeft As Boolean

Dim cMNIQ As New clsMNIQ
Public cQuiz As New clsQuiz
Dim strToolTip As String
Public clipText As String
Dim bonDraw As Boolean
Dim lastSelExample As String
Dim bAutoClick As Boolean
Dim iCatch As Integer '마우스 아웃으로 먹힐놈의 보기번호로 1~4까지이며 0인경우 아무것도 아닌경우다.

Dim sClaimeSubj As String
Dim sClaimeSeq As String

'============================
'Agent 설정
'----------------------------
'Const BalloonOn = 1
'Const SizeToText = 2
'Const AutoHide = 4
'Const AutoPace = 8
'============================

Dim X1 As Single
Dim Y1 As Single


Public oTooltip As clsToolTip

Dim time_tick1 As Long
Dim SSQL1 As String
Dim trust_firstHWND_wndPlayer As Long
Dim TmrAfterTTS_pushCount As Long

'The Main API call that will be used for the playback.
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal _
    lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As _
        Long, ByVal hwndCallback As Long) As Long
        
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long



Private Sub chk_Click()
'Debug.Print "chk_click()"

If chk.Tag = "m" Then '진정한 마우스 다운
    If chk.Value = vbChecked Then
        cQuiz.chk = 1
    Else
        cQuiz.chk = 0
    End If
Else
'Debug.Assert False
End If
End Sub

Private Sub chk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "chk_mousedown"
chk.Tag = "m"
End Sub

Private Sub chk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk.Tag = ""
End Sub

Private Sub chk_tts_dic_Click()
    If chk_tts_dic.Value = vbChecked Then
        mnu_auto_dic.Checked = True
        mnu_auto_tts.Checked = True
        mnuRefresh_Click
    Else
        mnu_auto_dic.Checked = False
        mnu_auto_tts.Checked = False
    End If
End Sub

Public Sub cmdClose_Click()
Dim ans As Integer
If cFTP.Profile.setPoP3 Then
    ans = MsgBox("시험문제를 종료하시겠습니까?     ", vbQuestion + vbYesNo)
Else
    ans = vbYes
End If
If ans = vbYes Then
    Me.Hide
    Unload Me
End If

#If MYAGENTUSE_ON Then
If Not frmMain.Character Is Nothing Then
    frmMain.Character.Balloon.Visible = False
    frmMain.Character.Stop
    Call frmMain.characterPlay("Congratulate")
End If
#End If

End Sub

Private Sub cmdDel_Click()
Dim ans As Integer
If cFTP.Profile.setPoP1 Then
    ans = MsgBox("삭제하시겠습니까?     ", vbYesNo + vbDefaultButton1 + vbQuestion)
Else
    ans = vbYes
End If

If ans = vbYes Then
    delprocess (cQuiz.num)
    
    quizDisp True, True
End If
End Sub

Private Sub cmdEE_Click()
Dim formTH01EE As New frmTH01EE

formTH01EE.sSubj = cQuiz.subj
formTH01EE.sSeq = cQuiz.seq

Call formTH01EE.fillAll

TmrAfterTTS_focus.Enabled = False
TmrAfterTTS_exit = True
formTH01EE.Show vbModal
TmrAfterTTS_exit = False

End Sub

Private Sub cmdHelp_Click()
TmrAfterTTS_focus.Enabled = False
frmHelp.Show
End Sub

Public Sub cmdHintRefresh()
If 5 < Len(frmMain.cQuiz.hint) Then
    If (Left(frmMain.cQuiz.hint, 4) = "@htt") Then
         Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" " & Mid(frmMain.cQuiz.hint, 2) & "")
        Exit Sub
    End If
End If

Call frmMain.characterPlay("Suggest")

If cQuiz.NoVisibleHintCmdCase() <> cmdHint.Value Then
    
'    Debug.Assert False
    frmHint.Visible = False
    If Not cQuiz.NoExistABCDinHint() Then
        quizDisp True
    End If
    
End If

frmHint.picMain.cls
frmHint.ScaleMode = vbTwips
frmHint.picMain.ScaleMode = vbTwips
frmHint.picMain.FontBold = True
frmHint.picMain.FontName = picMain.FontName
frmHint.picMain.FontSize = picMain.FontSize

Set frmHint.frmMain = frmMain

frmHint.picMain.FontSize = pic1.FontSize
'Debug.Assert frmHint.picMain.FontSize >= 11


frmHint.WindowState = vbNormal  '


If frmHint.WindowState <> vbNormal Then '위에 설정한 값이 안먹히는 이유는 안보여서 이다.
    frmHint.Visible = True
    
    frmHint.WindowState = vbNormal
End If

'frmHint.picMain.Width = pic1.Width

frmHint.Width = frmHint.picMain.TextWidth(cQuiz.hint2) + frmHint.picMain.Left * 2
frmHint.Height = frmHint.picMain.TextHeight(cQuiz.hint2) + frmHint.picMain.Top * 2.5 + Screen.TwipsPerPixelX + frmHint.TextHeight("A") * 1.5

frmHint.picMain.DrawMode = 13

frmHint.picMain.Print cQuiz.hint2


'========= 2004 12 25 s
'frmHINT 폼에 그림 생성
frmHint.cls
frmHint.PaintPicture frmHint.picMain.Image, 360, 405
frmHint.picMain.Visible = False
'========= 2004 12 25 e
'frmHint.Move frmQuiz.Left + picMain.Left, frmQuiz.Top + picMain.Top + cmdHint.Top + pic1.Top + cmdHint.Height, pic1.Width
frmHint.Visible = True
frmHint.ZOrder 0
cQuiz.chk = 1 '힌트보면 체크함
chk.Value = vbChecked
End Sub

Private Sub cmdHint_Click()
cmdHintRefresh
frmHint.Move frmQuiz.Left + picMain.Left, frmQuiz.Top + picMain.Top + cmdHint.Top + pic1.Top + cmdHint.Height, pic1.Width
End Sub

Private Sub cmdImgHint_Click()

mnuImageSearch_Click

End Sub

Private Sub cmdMemo_Click()

    TmrAfterTTS_focus.Enabled = False
Call frmMain.characterPlay("Writing")


If Not (lfrmMemo Is Nothing) Then
    If lfrmMemo.Visible = False Then
'        Debug.Assert False
        lfrmMemo.rtxt1.Text = cQuiz.memo
    End If
    Call Unload(lfrmMemo)
End If

Set lfrmMemo = Nothing

'If lfrmMemo Is Nothing Then
    Set lfrmMemo = New frmMemo
    Set lfrmMemo.parent = Me
    
    'With cQuiz
        cQuiz.memo = getTableVal("hint", "th01", "where userid='" & gUserid & "' and subj = '" & cQuiz.subj & "' and seq = " & cQuiz.seq & " ")
    'End With
    
    Call lfrmMemo.setVal(gUserid, cQuiz.subj, cQuiz.seq, cQuiz.memo, cFTP, cQuiz)
    Set lfrmMemo.frmQuiz_obj = Me
    
    TmrAfterTTS_focus.Enabled = False
    TmrAfterTTS_exit = True
    lfrmMemo.Show vbModal
    TmrAfterTTS_exit = False
'Else
    
'    Debug.Assert False
    
'End If

End Sub

Private Sub cmdNext_Click()
Dim depthAll As Long
Dim depthnow As Long
Dim ans As Integer
Dim Jindo As String

Dim delta As Double
If cQuiz.bNext Then '두번 연달아 클릭인 경우 무효화
    Exit Sub
End If
cQuiz.bNext = True
Static SOLVEPOSSILBE As Boolean
DoEvents
delta = Timer()


If cmdNext.Enabled = False Then
    cQuiz.bNext = False
    Exit Sub
End If
Me.cls

If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False And optE.Value = False Then
'    Debug.Assert False
    If pgBar.Value > (pgBar.Max - cFTP.Profile.DelayTime2) And Not SOLVEPOSSILBE Then
        MsgBox "문제를 이제 푸세요.     ", vbExclamation
        SOLVEPOSSILBE = True
        cQuiz.bNext = False
        Exit Sub '2초동안은 공백이 된 상태에서 다음 진행을 막는다.
    End If
Else

End If

cQuiz.Correct_chk = False
cQuiz.Correct = False
If Not cQuiz.isNew Then
    Select Case cQuiz.ans
    Case "A", "1"
        If optA.Value Then cQuiz.Correct = True
    Case "B", "2"
        If optB.Value Then cQuiz.Correct = True
    Case "C", "3"
        If optC.Value Then cQuiz.Correct = True
    Case "D", "4"
        If optD.Value Then cQuiz.Correct = True
    Case "E", "5"
        If optE.Value Then cQuiz.Correct = True
    End Select
    If cQuiz.Correct = False Then
        If optA.Value Then
            Call FocusRect1(1)
        ElseIf optB.Value Then
            Call FocusRect1(2)
        ElseIf optC.Value Then
            Call FocusRect1(3)
        ElseIf optD.Value Then
            Call FocusRect1(4)
        ElseIf optE.Value Then
            Call FocusRect1(5)
        Else
            Call FocusRect1(0)
        End If
    End If
End If
cQuiz.Correct_chk = True

Call cTg01.add_next(gUserid, cQuiz.subj)

If cQuiz.isNew Then
'전에 새로운 문제로 여기에 추가하였으나 아래로 옮겨서 update_ymd와 비교하였음
ElseIf cQuiz.forReview = False And cQuiz.isNew = False Then

'    If cQuiz.o + cQuiz.x = 0 Then
'        Call cTg01.add_new(gUserid, cQuiz.subj, cQuiz.isBefore)
'    ElseIf cQuiz.update_ymd <> GETYMD() Then
'        Call cTg01.add_review(gUserid, cQuiz.subj, cQuiz.isBefore)
'    End If
    
End If

SOLVEPOSSILBE = False

cmdNext.Enabled = False

Debug.Assert cQuiz.num <> 0

If Not (lfrmMemo Is Nothing) Then
   If lfrmMemo.Visible Then
      lfrmMemo.save
   Else
'      Debug.Assert False '메모의 내용이 변조되는지 테스트 해볼 것.(테스트 정상이면 지울것)2010.01.31
      lfrmMemo.rtxt1.Text = cQuiz.memo
   End If
End If

If Not cQuiz.forReview Then
    If cQuiz.isNew = False And cQuiz.Correct = False Then
        '틀린경우
'        If Incress2(0) = False Then Exit Sub 'db 에 업데이트 날짜별 학습관리
        If Incress(0) = False Then
            cQuiz.bNext = False
            Exit Sub 'db 에 업데이트 틀림체트
        End If
        cQuiz.forReview = True
        pgBar.Value = pgBar.Min
        pgBar.Max = cFTP.Profile.TimeOutSecStudy
        pgBar.Value = pgBar.Max
        quizDisp
        
            '5초 대기
        picMain.Refresh
        pic1.Refresh
        Call SleepEx(cFTP.Profile.DelayTime1, True)
        
        
        DoEvents
        cmdNext.Enabled = True
        cQuiz.bNext = False
        Exit Sub
    ElseIf cQuiz.isNew = False And cQuiz.Correct = True Then
        '맞은경우
'        If Incress2(1) = False Then Exit Sub 'db 에 업데이트 날짜별 학습관리
        
        delta = Timer() - delta
        
        If delta > 0 Then
            cQuiz.dLastSec = cQuiz.dLastSec - delta
        End If
        
'        If (Timer() - cQuiz.dLastSec) < 2 Then
'            cQuiz.dLastSec = 0
'        Else
            cQuiz.dLastSec = Timer() - cQuiz.dLastSec
'        End If

'      If IDEMODE Then
picMain.CurrentX = optA.Left
picMain.CurrentY = pic1.Top
picMain.ForeColor = vbWhite
bonDraw = False '하얀색 선이 안그려지도록 한다.
         cQuiz.dLastSec = Round15(cQuiz.dLastSec, 2)
         If cQuiz.dLastSec >= 60# Then
            cQuiz.dLastSec = Round15(cQuiz.dLastSec, 0)
         End If
         
         picMain.Print Round15(cQuiz.dLastSec, 2)
         
'      End If

      If cQuiz.chk = 0 Then
            picMain.DrawMode = 12
            picMain.DrawWidth = 10
'            picMain.Circle (3800, 1800), 1550, rgb(255 - 190, 255 - 50, 255 - 50), , , 1

            If cQuiz.dLastSec < Round15((Len(cQuiz.quiz) + Len(cQuiz.a)) / 10, 2) + 1 Then
                picMain.DrawWidth = 30
                cQuiz.dLastSec = 0
            Else
                picMain.DrawWidth = 10
            End If
            
            picMain.Circle (3800, 1800), 1550, vbYellow, , , 1
            picMain.DrawMode = 13
        Else '세모
            picMain.DrawMode = 12
            
            If cQuiz.dLastSec = 0 Then
                picMain.DrawWidth = 30
            Else
                picMain.DrawWidth = 10
            End If

            
'            picMain.Line (3800, (1800 - 1550))-(3800 + 1550 * 3 / 4, 1550 + 1550 / 2 + (1800 - 1550)), rgb(255 - 190, 255 - 50, 255 - 50)
'            picMain.Line -(3800 - 1550 * 3 / 4, 1550 + 1550 / 2 + (1800 - 1550)), rgb(255 - 190, 255 - 50, 255 - 50)
'            picMain.Line -(3800, (1800 - 1550)), rgb(255 - 190, 255 - 50, 255 - 50)
            picMain.Line (3800, (1800 - 1550))-(3800 + 1550 * 3 / 4, 1550 + 1550 / 2 + (1800 - 1550)), vbYellow
            picMain.Line -(3800 - 1550 * 3 / 4, 1550 + 1550 / 2 + (1800 - 1550)), vbYellow
            picMain.Line -(3800, (1800 - 1550)), vbYellow
            
            picMain.DrawMode = 13
        End If



        If Incress(1) = False Then
            cQuiz.bNext = False
            Exit Sub 'db 에 업데이트 맞음체크
        End If
    End If
End If

optA.Value = False
optB.Value = False
optC.Value = False
optD.Value = False
optE.Value = False


pgBar.Value = pgBar.Max
cQuiz.forReview = False

gHangSu = gHangSu + 1
gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow)


'성경인경우 난수발생으로 섞지 않는다.
If (cQuiz.subj = "구약성경" Or cQuiz.subj = "신약성경") Then
'성경인경우 난수발생으로 섞지 않는다.
Else
    Call nansu(gNowNum, gIsNew, cFTP.Profile.nansu)
End If

'Debug.Assert gLastNew <> 100
Write_TU03

'If cQuiz.overFlow And cQuiz.bPass = False  Then
If cQuiz.overFlow And cQuiz.bPass = False And cQuiz.isNew = False Then

    Jindo = GETjindo(gPocket, gChasu, Con.State)

    If Fix(Jindo) >= 100 Then
        ans = MsgBox("해당 시험지 문제를 모두 푸셨습니다. 중단하시겠습니까?", vbAbortRetryIgnore + vbQuestion)
        If ans = vbYes Then
            deletePaper gPocket, gChasu
            
            frmMain.mnuRefresh_Click ' mnuRefresh_Click
            Unload Me
        ElseIf ans = vbRetry Or ans = vbIgnore Then
            cQuiz.overFlow = False
            cQuiz.bPass = True
            quizDisp
        ElseIf ans = vbAbort Then 'vbno
            cmdClose_Click
            
            'quizDisp = False
            
            ans = MsgBox("틀린문제를 확인하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1)
            
            If ans = vbYes Then
                '틀린문제 시험지를 만든다
                
                Dim bSuccess As Boolean
                bSuccess = frmMain.TvNodeSelect(gPocket & " " & gChasu)
                
                If bSuccess Then
                    
                    Module1.vDataGlobal2 = ""
                    frmMain.mnuMakeSub_Click (0) '틀린문제시험지를 만든다음
                    
                    If Module1.vDataGlobal2 <> "" Then
                        bSuccess = frmMain.TvNodeSelect(Module1.vDataGlobal2)
                        
                        If bSuccess Then
                            '학습생략 후 바로 시작
                            frmMain.mnuQuickStart_Click
                        End If
                        
                    End If
                    
                End If
            End If
            
        Else
            Debug.Assert False
        End If
        
    Else
        quizDisp
            '5초 대기
        picMain.Refresh
        pic1.Refresh
        Call SleepEx(cFTP.Profile.DelayTime1, True)
        

    End If
    
Else
    cQuiz.overFlow = False
    
    If quizDisp = False Then
        cQuiz.bNext = False
        Exit Sub
    End If
    picMain.Refresh
    pic1.Refresh
    '5초 대기
    
    Call SleepEx(cFTP.Profile.DelayTime1, True)
    
End If


DoEvents
'종료되었는데 이부분을 진입하는 경우가 있음. 2010.11.03
If ans <> vbAbort Then '2010.11.03 if 문을 달은 이유는 unload 된경우에 의해 오류방지하기 위함.
    cmdNext.Enabled = True
End If
cQuiz.bNext = False '전역으로 사용하여 next버튼이 클릭된 상태에서 퀴즈가 디스플레이되는지의 여부를 판단한다. quizDisp 에서 사용함

End Sub

Private Sub cmdNext_KeyDown(keycode As Integer, Shift As Integer)
Call Form_KeyDown(keycode, Shift)
End Sub

Private Sub PenColorSettingProcess(ByVal lPenColor As Long)

    Dim red As Long, green As Long, blue As Long, Sum As Long
    
    red = RedFromRGB(lPenColor)
    green = GreenFromRGB(lPenColor)
    blue = BlueFromRGB(lPenColor)
    
'    sum = red + green + blue
    
'    If sum < 255 And red < 128 Then
'        red = red * 2
'    End If
'
'    If sum < 255 And green < 128 Then
'        green = green * 2
'    End If
'
'    If sum < 255 And blue < 128 Then
'        blue = blue * 2
'    End If
    
    cmdPen.BackColor = rgb(red, blue, green)
'--------------------------
'반전색 구하는 알고리즘
'    red = 255 - red
'    green = 255 - green
'    blue = 255 - blue
'--------------------------

'--------------------------
'보색 ComplementaryColor
'    sum = getMin(red, green, blue) + getMax(red, green, blue)
'
'    red = sum - red
'    green = sum - green
'    blue = sum - blue
'--------------------------

'-----------------------------
'Xor
'    red = red Xor &HFF
'    green = green Xor &HFF
'    blue = blue Xor &HFF
'-----------------------------

    Shape1.FillColor = rgb(red, green, blue)
    Shape2.FillColor = Shape1.FillColor Xor &HFFFFFF
    cmdPen.BackColor = Shape1.FillColor

End Sub

Private Sub cmdPen_Click()
    'color picker
    Dim lPenColor As Long
    dlgCMDialog.ShowColor
    
    lPenColor = dlgCMDialog.Color
    
    PenColorSettingProcess (lPenColor)
    
    cFTP.Profile.PenColor = cmdPen.BackColor
    
    'cmdPen.BackColor = lPenColor
End Sub

Public Function getMin(c1 As Long, C2 As Long, c3 As Long) As Long
    If c1 < C2 Then
        getMin = c1
        If getMin > c3 Then
            getMin = c3
        End If
    Else
        getMin = C2
        If getMin > c3 Then
            getMin = c3
        End If
    End If
End Function

Public Function getMax(c1 As Long, C2 As Long, c3 As Long) As Long
    If c1 < C2 Then
        getMax = C2
        If getMax < c3 Then
            getMax = c3
        End If
    Else
        getMax = c1
        If getMax < c3 Then
            getMax = c3
        End If
    End If
End Function


Private Sub cmdPre_Click()
Dim depth1 As Long
Dim depthnow As Long



cQuiz.isBefore = True
cQuiz.forReview = False '복습을 위한 설정은 아님  isnew 의 설정과는 틀림 forReview는 틀린경우만   true값이됨.
cQuiz.overFlow = False '오버플로 된 후 전으로 하면 오버플로 flag는 false 상태이다.


If gHangSu <> 1 Then
    gHangSu = gHangSu - 1

    gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depth1, depthnow)
    Write_TU03
    quizDisp
    
    Call cTg01.add_back(gUserid, cQuiz.subj, cQuiz.isNew)
Else
    gHangSu = 1
    cQuiz.isBefore = True
    gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depth1, depthnow)
    Write_TU03
    quizDisp
End If

cQuiz.isBefore = False
End Sub

Private Sub cmdRef_Click()
Load frmRef
Dim gubunja As String
Dim cd As Long
Dim cd2 As String
Dim fileNm As String
gubunja = Left(cQuiz.resid, 1)
cd2 = Mid(cQuiz.resid, 2)

'filenm=get

Select Case gubunja
Case "ⓘ" 'image (gif)
    Set frmRef.img.Picture = LoadPictureResource(cd2, "GIF")
    Set frmRef.pic.Picture = frmRef.img.Picture
    
    frmRef.pic.Visible = False
    
    frmRef.Width = frmRef.pic.Width + Screen.TwipsPerPixelX * 8
    frmRef.Height = frmRef.pic.Height + Screen.TwipsPerPixelY * 28
    
'    frmRef.img.Stretch = True
    
    'SaveResItemToDisk(cd,"GIF",t1)
Case "ⓗ" 'html(html)

Case "ⓜ" 'movie mpeg

Case "ⓢ" 'sound (mp3)

Case "@" 'text

Case Else 'res bitmap
    Set frmRef.img.Picture = LoadResPicture(cQuiz.resid, 0)
    Set frmRef.pic.Picture = frmRef.img.Picture
    
    
    frmRef.pic.Visible = False
    
    frmRef.Width = frmRef.pic.Width + Screen.TwipsPerPixelX * 8
    frmRef.Height = frmRef.pic.Height + Screen.TwipsPerPixelY * 28
    
    
End Select
frmRef.Visible = True
End Sub

Private Sub cmdTimer_Click()
Timer1.Enabled = Not Timer1.Enabled

cFTP.Profile.CntOfStreatOutNow = 0
'If Timer1.Enabled Then
'    cmdTimer.Caption = "||"
'Else
'    cmdTimer.Caption = "▶"
'End If
End Sub


Private Sub Command2_Click()
Read_TU03
Write_TU03
End Sub

'=============================
'간격 오류보정
'=============================
Private Sub Command1_Click()
Dim rs As ADODB.Recordset

Dim delta As Long

sSql = "select * from tu02 where reserve_ymd<>'99999999'"
Set rs = Fn_SQLExec(sSql).rs
Do Until rs.EOF
    cQuiz.update_ymd = rs("update_ymd")
    cQuiz.reserve_ymd = rs("reserve_ymd")
    cQuiz.gangyek = rs("gangyek")
    
    delta = Abs(DateDiff("d", str2Date(cQuiz.reserve_ymd), str2Date(cQuiz.update_ymd)))
    
    If cQuiz.gangyek > delta * 4 Then
        cQuiz.gangyek = delta * 4
        sSql = "update tu02 set gangyek=" & cQuiz.gangyek & " where subj='" & rs("subj") & "' and seq=" & rs("seq") & " and userid='" & rs("userid") & "'"
        delta = Fn_SQLExec(sSql).nrow
        Debug.Assert delta = 1
    ElseIf cQuiz.gangyek > 365 * 2 Then
        cQuiz.gangyek = delta * 2
        sSql = "update tu02 set gangyek=" & cQuiz.gangyek & " where subj='" & rs("subj") & "' and seq=" & rs("seq") & " and userid='" & rs("userid") & "'"
        delta = Fn_SQLExec(sSql).nrow
        Debug.Assert delta = 1
'        Debug.Assert False
    End If
    rs.MoveNext
Loop
End Sub

Private Sub Form_Activate()
Dim hIMC As Long





hIMC = ImmGetContext(Me.hwnd)

If bAutoClick Then
    optAutoClickOn.Value = True
Else
    optAutoClickOff.Value = True
End If

'Call ImmSetConversionStatus(hIMC, IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE)

Call ImmSetConversionStatus(hIMC, IME_CMODE_FIXED, IME_SMODE_NONE)

Shape2.FillColor = Shape1.FillColor Xor &HFFFFFF

End Sub

Private Sub Form_Deactivate()
 If frmQuiz.Visible Then
    '''TmrAfterTTS_focus.Enabled = True
 End If
End Sub

'Private Sub Form_Activate()
''Form_Resize
'End Sub

Private Sub Form_GotFocus()
'On Error Resume Next 'frmEndSetting 이 모달로 떳을때 이부분이 수행되면서 오류가 발생되어 스킵하게 한다.

If Not frmEndSetting.Visible Then
    frmMain.SetFocus
End If
End Sub

Private Sub Form_Initialize()
Set oTooltip = New clsToolTip
End Sub

Public Sub Form_KeyDown(keycode As Integer, Shift As Integer)

If wbQuizFocused = True Then Exit Sub

If Not optA.Visible Then Exit Sub '주관식문제인경우 진입하지 않음

Static oldkey As Integer



If optA.Visible And (keycode = vbKey1 Or keycode = vbKeyA Or keycode = vbKeyNumpad1) Then
    optA.Value = True
    picMain.SetFocus
    If oldkey = keycode Then
        oldkey = 0
        keycode = 0
        SendKeys "{enter}"
    ElseIf keycode = vbKeyNumpad1 Or keycode = vbKey1 Then
        oldkey = 0
        keycode = 0
        cmdNext_Click
    End If
    
ElseIf optB.Visible And (keycode = vbKey2 Or keycode = vbKeyS Or keycode = vbKeyNumpad2) Then
    optB.Value = True
    picMain.SetFocus
    If oldkey = keycode Then
        oldkey = 0
        keycode = 0
        SendKeys "{enter}"
    ElseIf keycode = vbKeyNumpad2 Or keycode = vbKey2 Then
        oldkey = 0
        keycode = 0
        cmdNext_Click
    End If

ElseIf optC.Visible And (keycode = vbKey3 Or keycode = vbKeyD Or keycode = vbKeyNumpad3) Then
    optC.Value = True
    picMain.SetFocus
    If oldkey = keycode Then
        oldkey = 0
        keycode = 0
        SendKeys "{enter}"
    ElseIf keycode = vbKeyNumpad3 Or keycode = vbKey3 Then
        oldkey = 0
        keycode = 0
        cmdNext_Click
    End If

ElseIf optD.Visible And (keycode = vbKey4 Or keycode = vbKeyF Or keycode = vbKeyNumpad4) Then
    optD.Value = True
    picMain.SetFocus
    If oldkey = keycode Then
        oldkey = 0
        keycode = 0
        SendKeys "{enter}"
    ElseIf keycode = vbKeyNumpad4 Or keycode = vbKey4 Then
        oldkey = 0
        keycode = 0
        cmdNext_Click
    End If
ElseIf optE.Visible And (keycode = vbKey5 Or keycode = vbKeyG Or keycode = vbKeyNumpad5) Then
    optE.Value = True
    picMain.SetFocus
    If oldkey = keycode Then
        oldkey = 0
        keycode = 0
        SendKeys "{enter}"
    ElseIf keycode = vbKeyNumpad5 Or keycode = vbKey5 Then
        oldkey = 0
        keycode = 0
        cmdNext_Click
    End If
ElseIf keycode = vbKeySpace Or keycode = vbKeyRight Or keycode = vbKeyN Then
    oldkey = 0
    keycode = 0
    cmdNext_Click
ElseIf keycode = vbKeyLeft Or keycode = vbKeyB Then
    oldkey = 0
    keycode = 0
    cmdPre_Click
ElseIf keycode = vbKeyDelete And cQuiz.isNew Then
    oldkey = 0
    keycode = 0
    cmdDel_Click
Else
    If keycode = vbKey5 Or keycode = vbKey6 Then
        keycode = 0
        cmdHint_Click
    End If
    oldkey = 0
End If

If keycode = vbKeyJ Or keycode = vbKeyDown Then
    If optA.Value And optB.Visible Then
        optB.Value = True
    ElseIf optB.Value And optC.Visible Then
        optC.Value = True
    ElseIf optC.Value And optD.Visible Then
        optD.Value = True
    ElseIf optD.Value And optE.Visible Then
        optE.Value = True
    Else
        'Debug.Assert Fase
        If True Then '맨위 아래 다시 나타나기
            If optA.Visible And optB.Value = False And optC.Value = False And optD.Value = False And optE.Value = False Then
                optA.Value = True
            ElseIf optA.Visible Then
                optA.Value = True
            End If
        Else
            If optA.Visible And optB.Value = False And optC.Value = False And optD.Value = False And optE.Value = False Then
                optA.Value = True
            End If
        End If
    End If
    keycode = 0
ElseIf keycode = vbKeyK Or keycode = vbKeyUp Then

    If True Then '맨위 아래로 다시 나타나기
        If optA.Value And optA.Visible Then
            If optE.Visible Then
                optE.Value = True
            ElseIf optD.Visible Then
                optD.Value = True
            ElseIf optC.Visible Then
                optC.Value = True
            ElseIf optB.Visible Then
                optB.Value = True
            ElseIf optA.Visible Then
                optA.Value = True
            End If

        ElseIf optB.Value And optA.Visible Then
            optA.Value = True
        ElseIf optC.Value And optB.Visible Then
            optB.Value = True
        ElseIf optD.Value And optC.Visible Then
            optC.Value = True
        ElseIf optE.Value And optD.Visible Then
            optD.Value = True
        Else
    
            If optE.Visible Then
                optE.Value = True
            ElseIf optD.Visible Then
                optD.Value = True
            ElseIf optC.Visible Then
                optC.Value = True
            ElseIf optB.Visible Then
                optB.Value = True
            ElseIf optA.Visible Then
                optA.Value = True
            End If
        End If
    
    Else
        If optA.Value And optA.Visible Then
            optA.Value = True
        ElseIf optB.Value And optA.Visible Then
            optA.Value = True
        ElseIf optC.Value And optB.Visible Then
            optB.Value = True
        ElseIf optD.Value And optC.Visible Then
            optC.Value = True
        ElseIf optE.Value And optD.Visible Then
            optD.Value = True
        Else
    
            If optE.Visible Then
                optE.Value = True
            ElseIf optD.Visible Then
                optD.Value = True
            ElseIf optC.Visible Then
                optC.Value = True
            ElseIf optB.Visible Then
                optB.Value = True
            ElseIf optA.Visible Then
                optA.Value = True
            End If
        End If
    
    End If
    keycode = 0
ElseIf keycode = vbKeyH Then
    cmdHint_Click
    keycode = 0
ElseIf keycode = vbKeyI Or keycode = vbKeyR Or keycode = vbKeyE Then
    cmdTimer_Click
    keycode = 0
End If

If keycode = vbKeyA And optA.Visible Or _
   keycode = vbKeyS And optB.Visible Or _
   keycode = vbKeyD And optC.Visible Or _
   keycode = vbKeyF And optD.Visible Or _
   keycode = vbKeyG And optE.Visible Then
    oldkey = keycode
'    KeyCode = 0
End If

If keycode = vbKeyZ Then
    optA.Value = False: optB.Value = False: optC.Value = False: optD.Value = False
    optE.Value = False
    oldkey = 0
End If

'If KeyCode = Asc("n") Or KeyCode = Asc("N") Or KeyCode = vbKeyRight Then
'    cmdNext_Click
'End If

'If KeyCode = Asc("b") Or KeyCode = Asc("B") Or KeyCode = vbKeyLeft Then
'    cmdPre_Click
'End If

'If KeyCode = vbKeyDelete And cQuiz.isNew Then
'    cmdDel_Click
'End If

If keycode > 200 Then  '객관식이면
'    MsgBox "한글입력 모드입니다.     ", vbInformation
    Form_Activate
End If
    oldkey = keycode
End Sub


Public Sub Form_KeyPress(KeyAscii As Integer)
Debug.Print "KEYPRESSED", KeyAscii
End Sub


'Public Sub Form_KeyPress(KeyAscii As Integer)
'Debug.Print "KEYPRESSED", KeyAscii
'End Sub

Private Sub Form_Load()

PenColorSettingProcess (cFTP.Profile.PenColor)

Dim ctrl As Control

WindowsMediaPlayer1.Visible = False

Set refCquiz = cQuiz
Toolbar1.Align = vbAlignBottom
pic1.AutoRedraw = True
picMain.AutoRedraw = True
    
imgA1.BorderStyle = 0
imgB1.BorderStyle = 0
imgC1.BorderStyle = 0
imgD1.BorderStyle = 0
imgE1.BorderStyle = 0


imgA1.Width = Screen.Width
imgB1.Width = Screen.Width
imgC1.Width = Screen.Width
imgD1.Width = Screen.Width
imgE1.Width = Screen.Width

imgE1.ZOrder 0
imgD1.ZOrder 0
imgC1.ZOrder 0
imgB1.ZOrder 0
imgA1.ZOrder 0


mnuPop1.Visible = False

Call oTooltip.Create(Me)
'With oTooltip

    oTooltip.MaxTipWidth = 240
    oTooltip.DelayTime(ttDelayShow) = 20000
'
'    .Transparency = 150
    oTooltip.Style = ttBaloon
    
    For Each ctrl In Me.Controls
        If ctrl.Name Like "cmd*" Then
            Call oTooltip.AddTool(ctrl)
        End If
    Next
'
'    Call .AddTool(cmdMemo)
    
'End With

pgBar.Value = pgBar.Max

pic1.BorderStyle = 0 'None
'gNowNum = Read_TU03

Load frmHint

frmHint.Visible = False
tmrTG.Enabled = True '학습시간 기록


read_tg01

'wb 위치 조정
'frmMain.wb.Left = PicMain.Left
'frmMain.wb.Top = PicMain.Top

'frmMain.wb2.Left = PicMain.Left
'frmMain.wb2.Top = PicMain.Top
mnu_auto_tts.Checked = cFTP.Profile.bChkTTSuse

If gbIsSuperAdmin Then
    mnu_Edit.Visible = True
Else
    mnu_Edit.Visible = False
End If

Set fMainForm.cQuiz = cQuiz
'Set frmMain.cQuiz = cQuiz

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cTg01.addsec(gUserid, cQuiz.subj, True) '저장한다.


cQuiz.cls
Unload frmHint
Timer1.Enabled = False
tmrTG.Enabled = False
save_status



End Sub

Function save_status()
If Len(cmdMemo.Tag) > 0 Then
    Unload frmMemo
End If
End Function

Private Sub Form_Resize()
gQuizOnResize = True '이 전역변수 값은 프로그래스 값에서 쓰인다.
gPgBarSaveValue = pgBar.Value '값 저장
Me.ScaleMode = vbTwips
picMain.ScaleMode = vbTwips
pic1.ScaleMode = vbTwips



Dim wd As Single

wd = Me.Width - pgBar.Left - picMain.Left
On Error Resume Next
pgBar.Width = wd ' Me.Width - pgBar.Left - picMain.Left
On Error GoTo ErrTrap

picMain.Width = (frmQuiz.Width - picMain.Left * 2) / 2
picMain.Height = frmQuiz.Height - picMain.Top - Toolbar1.Height
'pic1.Width = Screen.Width ' PicMain.Width
'pic1.Height = Screen.Height ' PicMain.Height
pic1.Width = picMain.Width '200802
pic1.Height = picMain.Height

wbQuizMain.Left = picMain.Width + 10
wbQuizMain.Top = picMain.Top
wbQuizMain.Width = picMain.Width
wbQuizMain.Height = picMain.Height

cQuiz.hint2 = autoCRLF(cQuiz.hint, picMain.Width - pic1.Left * 2, pic1, True)

'quizDisp True

If cQuiz.ans <> "" Then
    MagicHintView
End If
If Me.Visible Then
    quizDisp False 'refresh
End If
gQuizOnResize = False

pgBarJindo.Width = Toolbar1.Width / 2 - 20
pgBarJindo.Left = Toolbar1.Left
pgBarJindo.Top = Toolbar1.Top - pgBarJindo.Height

txtJindo.Left = (pgBarJindo.Width - txtJindo.Width) / 2
txtJindo.Top = pgBarJindo.Top + pgBarJindo.Height + 50

Exit Sub
ErrTrap:
    MsgBox err.Description, vbExclamation
gQuizOnResize = False
End Sub

Private Sub Form_Terminate()
Set oTooltip = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (frmMain Is Nothing) Then
    frmMain.mnuEdit2.Visible = False
End If
If Not (lfrmMemo Is Nothing) Then
    If lfrmMemo.Visible = False Then
        lfrmMemo.rtxt1.Text = cQuiz.memo
    End If
    Unload lfrmMemo
    Set lfrmMemo = Nothing
End If
If Not (frmMain Is Nothing) Then
    frmMain.SetFocus
End If
End Sub



Private Sub imgA1_Click()
If lastUpButtonLeft Then
    optA.Value = True
    Call FocusRect1(1)
End If
End Sub

Sub FocusRect1(idx As Integer)

Dim W As Single
Dim opt As OptionButton
Dim img As Image
Dim h As Single
Dim T As String

W = 80

If idx = 0 Then
    shpSel.Visible = False
    Exit Sub
ElseIf idx = 1 Then
    Set opt = optA
    Set img = imgA1
    T = cQuiz.a2
ElseIf idx = 2 Then
    Set opt = optB
    Set img = imgB1
    T = cQuiz.b2
ElseIf idx = 3 Then
    Set opt = optC
    Set img = imgC1
    T = cQuiz.C2
ElseIf idx = 4 Then
    Set opt = optD
    Set img = imgD1
    T = cQuiz.d2
ElseIf idx = 5 Then
    Set opt = optE
    Set img = imgE1
    T = cQuiz.e2
End If

If T = "" Then
    T = " "
End If

h = picMain.TextHeight(T)

'Call shpSel.Move(opt.Left - w, opt.Top - w, picMain.Width - opt.Left - w, H + 2 * w)

Call shpSel.Move(opt.Left - W, opt.Top - W, picMain.Width - opt.Left - W, h + 2 * W)

shpSel.Visible = True

End Sub

Private Sub imgA1_DblClick()
cmdNext_Click
End Sub


Private Sub imgA1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bonDraw = True
    picMain.DrawWidth = cFTP.Profile.FontSize / 5
    Call PicMainDrawModeForeColor
    picMain.Line (X + imgA1.Left, Y + imgA1.Top)-(X + imgA1.Left, Y + imgA1.Top)

End Sub

Private Sub imgA1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static preX As Single
Static preY As Single

Dim Dist As Single

Call opt_MouseOut

If bonDraw Then
    
    If preX = 0 And preY = 0 Then
        preX = X
        preY = Y
        Exit Sub
    End If
    Dist = ((X - preX) ^ 2 + (Y - preY) ^ 2) ^ 0.5 '[twip]
    
    If Dist > picMain.DrawWidth * Screen.TwipsPerPixelX Then
    
        picMain.Line -(X + imgA1.Left, Y + imgA1.Top)
        preX = X
        preY = Y
    End If
End If

End Sub

Private Sub imgA1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button And vbLeftButton) = vbLeftButton Then
'    optA.Value = True
 lastUpButtonLeft = True
Else
    lastUpButtonLeft = False
    clipText = cQuiz.a
    Me.PopupMenu mnuPop1
End If
Call PicMain_MouseUp(Button, Shift, X, Y)
End Sub




Private Sub imgB1_Click()
If lastUpButtonLeft Then
    optB.Value = True
    Call FocusRect1(2)
End If
End Sub

Private Sub imgB1_DblClick()
cmdNext_Click
End Sub


Private Sub imgB1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bonDraw = True
    picMain.DrawWidth = cFTP.Profile.FontSize / 5
    Call PicMainDrawModeForeColor
    picMain.Line (X + imgB1.Left, Y + imgB1.Top)-(X + imgB1.Left, Y + imgB1.Top)

End Sub

Private Sub imgB1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static preX As Single
Static preY As Single

Dim Dist As Single

Call opt_MouseOut

If bonDraw Then
    
    If preX = 0 And preY = 0 Then
        preX = X
        preY = Y
        Exit Sub
    End If
    Dist = ((X - preX) ^ 2 + (Y - preY) ^ 2) ^ 0.5 '[twip]
    
    If Dist > picMain.DrawWidth * Screen.TwipsPerPixelX Then
    
        picMain.Line -(X + imgB1.Left, Y + imgB1.Top)
        preX = X
        preY = Y
    End If
End If


End Sub

Private Sub imgB1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button And vbLeftButton) = vbLeftButton Then
    lastUpButtonLeft = True
Else
    lastUpButtonLeft = False
    clipText = cQuiz.B
    Me.PopupMenu mnuPop1
End If
Call PicMain_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub imgC1_Click()
If lastUpButtonLeft Then
    optC.Value = True
    Call FocusRect1(3)
End If
End Sub

Private Sub imgC1_DblClick()
cmdNext_Click
End Sub


Private Sub imgC1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bonDraw = True
    picMain.DrawWidth = cFTP.Profile.FontSize / 5
    Call PicMainDrawModeForeColor
    picMain.Line (X + imgC1.Left, Y + imgC1.Top)-(X + imgC1.Left, Y + imgC1.Top)

End Sub

Private Sub imgC1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static preX As Single
Static preY As Single

Dim Dist As Single

Call opt_MouseOut

If bonDraw Then
    
    If preX = 0 And preY = 0 Then
        preX = X
        preY = Y
        Exit Sub
    End If
    Dist = ((X - preX) ^ 2 + (Y - preY) ^ 2) ^ 0.5 '[twip]
    
    If Dist > picMain.DrawWidth * Screen.TwipsPerPixelX Then
    
        picMain.Line -(X + imgC1.Left, Y + imgC1.Top)
        preX = X
        preY = Y
    End If
End If


End Sub

Private Sub imgC1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button And vbLeftButton) = vbLeftButton Then
    lastUpButtonLeft = True
Else
    lastUpButtonLeft = False
    clipText = cQuiz.C
    Me.PopupMenu mnuPop1
End If
Call PicMain_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub imgD1_Click()
If lastUpButtonLeft Then
    optD.Value = True
    Call FocusRect1(4)
End If
End Sub

Private Sub imgD1_DblClick()
cmdNext_Click
End Sub

Private Sub imgD1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bonDraw = True
    picMain.DrawWidth = cFTP.Profile.FontSize / 5
    Call PicMainDrawModeForeColor
    picMain.Line (X + imgD1.Left, Y + imgD1.Top)-(X + imgD1.Left, Y + imgD1.Top)

End Sub

Private Sub imgD1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static preX As Single
Static preY As Single

Dim Dist As Single

Call opt_MouseOut

If bonDraw Then
    
    If preX = 0 And preY = 0 Then
        preX = X
        preY = Y
        Exit Sub
    End If
    Dist = ((X - preX) ^ 2 + (Y - preY) ^ 2) ^ 0.5 '[twip]
    
    If Dist > picMain.DrawWidth * Screen.TwipsPerPixelX Then
    
        picMain.Line -(X + imgD1.Left, Y + imgD1.Top)
        preX = X
        preY = Y
    End If
End If


End Sub

Private Sub imgD1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button And vbLeftButton) = vbLeftButton Then
    lastUpButtonLeft = True
Else
    lastUpButtonLeft = False
    clipText = cQuiz.d
    Me.PopupMenu mnuPop1
End If
Call PicMain_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub imgE1_Click()
If lastUpButtonLeft Then
    optE.Value = True
    Call FocusRect1(5)
End If
End Sub

Private Sub imgE1_DblClick()
cmdNext_Click
End Sub


Private Sub imgE1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bonDraw = True
    picMain.DrawWidth = cFTP.Profile.FontSize / 5
    Call PicMainDrawModeForeColor
    picMain.Line (X + imgE1.Left, Y + imgE1.Top)-(X + imgE1.Left, Y + imgE1.Top)

End Sub

Private Sub imgE1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static preX As Single
Static preY As Single

Dim Dist As Single

Call opt_MouseOut

If bonDraw Then
    
    If preX = 0 And preY = 0 Then
        preX = X
        preY = Y
        Exit Sub
    End If
    Dist = ((X - preX) ^ 2 + (Y - preY) ^ 2) ^ 0.5 '[twip]
    
    If Dist > picMain.DrawWidth * Screen.TwipsPerPixelX Then
    
        picMain.Line -(X + imgE1.Left, Y + imgE1.Top)
        preX = X
        preY = Y
    End If
End If


End Sub

Private Sub imgE1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button And vbLeftButton) = vbLeftButton Then
    lastUpButtonLeft = True
Else
    lastUpButtonLeft = False
    clipText = cQuiz.e
    Me.PopupMenu mnuPop1
End If
Call PicMain_MouseUp(Button, Shift, X, Y)
End Sub



Private Sub mnu_auto_dic_Click()
mnu_auto_dic.Checked = Not mnu_auto_dic.Checked
End Sub

Private Sub mnu_auto_tts_Click() '새로운 문항에서 TTS 설정을 하면 TTS를 실행한다.
'TTS실행 종류는 2가지가 있는데 새로운 문항을 학습 할경우만 수행하는 경우와 모두가있다.
mnu_auto_tts.Checked = Not mnu_auto_tts.Checked

cFTP.Profile.bChkTTSuse = mnu_auto_tts.Checked

cFTP.Profile.save

If mnu_auto_tts.Checked Then
    If gIsNew Then
        '새로고침 수행
        quizDisp False
    End If
End If



End Sub

'중국어사전
Private Sub mnu_cn_dic_Click()
    Call make_html_pop_cn(clipText, "tmp\$cn_dic" + setRndName() + ".htm")
    frmMain.wb2.Navigate2 App.Path + "\tmp\$cn_dic" + getRndName() + ".htm"
End Sub

Private Function isExistsClaime() As Boolean
Con_Open

sSql = "select * from th01 where userid='@'"

Set rs = Fn_SQLExec(sSql, , , True).rs

Dim sHint As Variant

If Not rs.EOF Then
    
    sClaimeSubj = rs("subj")
    sClaimeSeq = rs("seq")
    isExistsClaime = True
End If

Con_Close

End Function

Private Sub mnu_Edit_Click()

Dim formVq01 As New frmVq01
Dim retVal As VbMsgBoxResult

Load formVq01



formVq01.txtSubj = cQuiz.subj
formVq01.txtSeq = cQuiz.seq

If isExistsClaime() Then
    retVal = MsgBox("이의신청(오류신고)된 내용이 있습니다. 신고된 내역을 보시겠습니까?", vbExclamation + vbYesNo, "질문")
    If retVal = vbYes Then
        formVq01.txtSubj = sClaimeSubj ' "1998AA"
        formVq01.txtSeq = sClaimeSeq '"75"
        formVq01.cmdSave.Enabled = False
    Else
        formVq01.cmdTH01.Enabled = False
    End If
Else
    formVq01.cmdTH01.Enabled = False
End If
'
'formVq01.txtSubj = "1998AA"
'formVq01.txtSeq = "75"
'
'formVq01.txtSubj = "영속담1"
'formVq01.txtSeq = "34"
'
'formVq01.txtSubj = "특목고입시용"
'formVq01.txtSeq = "666"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "38"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "50"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "85"
'
'formVq01.txtSubj = "모의AA"
'formVq01.txtSeq = "184"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "20"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "119"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "219"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "164"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "209"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "248"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "242"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "264"
'
'formVq01.txtSubj = "2003AA"
''formVq01.txtSeq = "269"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "270"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "272"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "273"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "298"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "276"
'
'formVq01.txtSubj = "모의AA"
'formVq01.txtSeq = "177"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "303"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "260"
'
'formVq01.txtSubj = "토익Voca"
'formVq01.txtSeq = "205"
'
'formVq01.txtSubj = "토익Voca"
'formVq01.txtSeq = "637"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "318"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "334"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "335"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "328"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "149"
'
'formVq01.txtSubj = "영속담1"
'formVq01.txtSeq = "227"
'
'formVq01.txtSubj = "영단어1"
'formVq01.txtSeq = "658"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "343"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "344"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "345"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "352"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "342"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "294"
'
'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "137"
'
'formVq01.txtSubj = "토익Voca"
'formVq01.txtSeq = "135"
'
'formVq01.txtSubj = "토익Voca"
'formVq01.txtSeq = "230"
'
'formVq01.txtSubj = "토익Voca"
'formVq01.txtSeq = "437"
''
'formVq01.txtSubj = "영속담1"
'formVq01.txtSeq = "99"

'formVq01.txtSubj = "2003AA"
'formVq01.txtSeq = "3"
'
'formVq01.txtSubj = "토익990"
'formVq01.txtSeq = "21"

'formVq01.txtSubj = "토익Voca"
'formVq01.txtSeq = "978"

'formVq01.txtSubj = "영속담1"
'formVq01.txtSeq = "14"


Call formVq01.fillAll

Set formVq01.parent = Me

TmrAfterTTS_focus.Enabled = False
formVq01.Show vbModal
If formVq01.saved Then
    Call quizDisp
End If

Unload formVq01
Set formVq01 = Nothing

End Sub

'일어사전
Private Sub mnu_jp_dic_Click()
    Call make_html_pop_jp(clipText, "tmp\$jp_dic" + setRndName() + ".htm")
    frmMain.wb2.Navigate2 App.Path + "\tmp\$jp_dic" + getRndName() + ".htm"
End Sub
'국어사전
Private Sub mnu_kr_dic_Click()
    Call make_html_pop_kr(clipText, "tmp\$kr_dic" + setRndName() + ".htm")
    frmMain.wb2.Navigate2 App.Path + "\tmp\$kr_dic" + getRndName() + ".htm"
End Sub

Private Sub mnu_naverModum_search_Click()
    If Len(clipText) = LenH(clipText) Then
        Call wbQuizMain.Navigate("http://dic.naver.com/search.nhn?target=dic&ie=utf8&query_utf=&isOnlyViewEE=&query=" & clipText)
    Else
        Call make_html_pop_modum(clipText, "tmp\$modum_dic" + setRndName() + ".htm")
        frmMain.wb2.Navigate2 App.Path + "\tmp\$modum_dic" + getRndName() + ".htm"
    'Call wbQuizMain.Navigate(App.Path + "\tmp\$modum_dic" + getRndName() + ".htm")
    End If
End Sub

Private Sub mnu_ref_see_Click()
Dim str As String

Con_Open
sSql = "select quiz from vq01 where subj='" & cQuiz.subj & "' and a='" & clipText & "' limit 1"
Set rs = Fn_SQLExec(sSql, , , True).rs
If Not rs.EOF Then
    str = rs(0)
    Con_Close
    
    Dim saved_status As Boolean
    saved_status = TmrAfterTTS_focus.Enabled
    TmrAfterTTS_focus.Enabled = False
    
    Dim retVal
    retVal = MsgBox(clipText + vbNewLine + vbNewLine + str & vbNewLine & "━━━━━━━━━━━━━━━━━━━━" & vbNewLine & "예: 클립보드로 모두 복사" & vbNewLine & "아니요: 클립모드로 관련문제 복사" & vbNewLine & "취소: 복사하지 않고 닫기", vbInformation + vbYesNoCancel + vbDefaultButton2, "클립보드로 복사하려면 확인버튼을 클릭하세요.")
    If retVal = vbYes Then
        Clipboard.Clear
        Clipboard.SetText clipText & "=>" & str & vbNewLine
    ElseIf retVal = vbNo Then
        Clipboard.Clear
        Clipboard.SetText str
    End If
    
    TmrAfterTTS_focus.Enabled = saved_status
Else
    Con_Close
    MsgBox "조회 결과가 없습니다", vbExclamation + vbOKOnly, "클립보드로 복사하려면 확인버튼을 클릭하세요."
End If
End Sub

Private Sub mnuBibleSearch_Click()
Dim strQ As String
strQ = cQuiz.cat

Dim pos1 As Integer
pos1 = InStr(strQ, ":")

If pos1 > 0 Then
    strQ = Mid(strQ, pos1 + 3)
Else
    strQ = ""
End If

If strQ = "" Then 'http://www.bskorea.or.kr/infobank/korSearch/korbibReadpage.aspx?version=SAE&book=est&chap=3&sec=12&cVersion=&fontString=12px&fontSize=1#focus#focus
    Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.bskorea.or.kr") '홀리바이블에서 대한성서공회로 바꿈, 이유 사이트가 더이상 안열려서...2015.02.24
Else

    Dim chap1 As Long, sep1 As Long
    chap1 = InStr(strQ, " ")
    sep1 = InStr(strQ, ":")
    
    Dim str1 As String, str2 As String
    str1 = Mid(Left(strQ, sep1 - 1), chap1 + 1)
    str2 = Mid(strQ, sep1 + 1)
    
    'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.holybible.or.kr/B_HDB/cgi/biblesrch.php?VR=HDB&QR=" & strQ & "&OD=")
    Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.bskorea.or.kr/infobank/korSearch/korbibReadpage.aspx?version=SAE&book=" & Kor2EngBibleChapter(Left(strQ, chap1 - 1)) & "&chap=" & str1 & "&sec=" & str2 & "&cVersion=&fontString=12px&fontSize=1#focus#focus")
End If

End Sub

Private Function Kor2EngBibleChapter(str As String) As String

Dim retVal As String

If str = "창" Then
retVal = "gen"
ElseIf str = "출" Then
retVal = "exo"
ElseIf str = "레" Then
retVal = "lev"
ElseIf str = "민" Then
retVal = "num"
ElseIf str = "신" Then
retVal = "deu"
ElseIf str = "수" Then
retVal = "jos"
ElseIf str = "삿" Then
retVal = "jdg"
ElseIf str = "룻" Then
retVal = "rut"
ElseIf str = "삼상" Then
retVal = "1sa"
ElseIf str = "삼하" Then
retVal = "2sa"
ElseIf str = "왕상" Then
retVal = "1ki"
ElseIf str = "왕하" Then
retVal = "2ki"
ElseIf str = "대상" Then
retVal = "1ch"
ElseIf str = "대하" Then
retVal = "2ch"
ElseIf str = "스" Then
retVal = "ezr"
ElseIf str = "느" Then
retVal = "neh"
ElseIf str = "에" Then
retVal = "est"
ElseIf str = "욥" Then
retVal = "job"
ElseIf str = "시" Then
retVal = "psa"
ElseIf str = "잠" Then
retVal = "pro"
ElseIf str = "전" Then
retVal = "ecc"
ElseIf str = "아" Then
retVal = "sng"
ElseIf str = "사" Then
retVal = "isa"
ElseIf str = "렘" Then
retVal = "jer"
ElseIf str = "애" Then
retVal = "lam"
ElseIf str = "겔" Then
retVal = "ezk"
ElseIf str = "단" Then
retVal = "dan"
ElseIf str = "호" Then
retVal = "hos"
ElseIf str = "욜" Then
retVal = "jol"
ElseIf str = "암" Then
retVal = "amo"
ElseIf str = "옵" Then
retVal = "oba"
ElseIf str = "욘" Then
retVal = "jnh"
ElseIf str = "미" Then
retVal = "mic"
ElseIf str = "나" Then
retVal = "nam"
ElseIf str = "합" Then
retVal = "hab"
ElseIf str = "습" Then
retVal = "zep"
ElseIf str = "학" Then
retVal = "hag"
ElseIf str = "슥" Then
retVal = "zec"
ElseIf str = "말" Then
retVal = "mal"
ElseIf str = "마" Then
retVal = "mat"
ElseIf str = "막" Then
retVal = "mrk"
ElseIf str = "눅" Then
retVal = "luk"
ElseIf str = "요" Then
retVal = "jhn"
ElseIf str = "행" Then
retVal = "act"
ElseIf str = "롬" Then
retVal = "rom"
ElseIf str = "고전" Then
retVal = "1co"
ElseIf str = "고후" Then
retVal = "2co"
ElseIf str = "갈" Then
retVal = "gal"
ElseIf str = "엡" Then
retVal = "eph"
ElseIf str = "빌" Then
retVal = "php"
ElseIf str = "골" Then
retVal = "col"
ElseIf str = "살전" Then
retVal = "1th"
ElseIf str = "살후" Then
retVal = "2th"
ElseIf str = "딤전" Then
retVal = "1ti"
ElseIf str = "딤후" Then
retVal = "2ti"
ElseIf str = "딛" Then
retVal = "tit"
ElseIf str = "몬" Then
retVal = "phm"
ElseIf str = "히" Then
retVal = "heb"
ElseIf str = "약" Then
retVal = "jas"
ElseIf str = "벧전" Then
retVal = "1pe"
ElseIf str = "벧후" Then
retVal = "2pe"
ElseIf str = "요일" Then
retVal = "1jn"
ElseIf str = "요이" Then
retVal = "2jn"
ElseIf str = "요삼" Then
retVal = "3jn"
ElseIf str = "유" Then
retVal = "jud"
ElseIf str = "계" Then
retVal = "rev"
End If

Kor2EngBibleChapter = retVal
End Function

Private Sub mnuBritannica100_Click()
'브리태니커 사전
Dim strQ As String
'strQ = URLEncodeUTF8(clipText)
strQ = Replace(Replace(Replace(clipText, "(", " "), ")", " "), "+", " ") '무한루프검색 방지
Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://preview.britannica.co.kr/search/s97_utf8.exe?QueryText=" & strQ & "&DBase=Article_Up")

End Sub

Private Sub mnuChrome_Click()

If isFile("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
    Call Shell("""C:\Program Files\Google\Chrome\Application\chrome.exe""")
Else
    MsgBox "크롬이 C:\Program Files\Google\Chrome\Application\chrome.exe 경로에 설치되지 않았습니다.", vbExclamation, "포켓퀴즈"
    mnuChrome.Enabled = False
End If

End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText clipText

End Sub
'
'Private Sub optA_Click()
'cmdNext_Click
'End Sub
'
'Private Sub optB_Click()
'cmdNext_Click
'End Sub
'Private Sub optC_Click()
'cmdNext_Click
'End Sub
'
'Private Sub optD_Click()
'cmdNext_Click
'End Sub
'
'Private Sub optE_Click()
'cmdNext_Click
'End Sub

Private Sub mnu_all_in_one_Click()
Call mnuTTS0_Click
Call mnuEndic_Click

End Sub



Private Sub mnuDoosan100_Click()
'http://www.doopedia.co.kr/search/encyber/totalSearch.jsp?category=total&autoType=web&searchTerm=이나무
Dim strQ As String
strQ = URLEncodeUTF8(clipText)
Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.doopedia.co.kr/search/encyber/totalSearch.jsp?category=total&autoType=web&searchTerm=" & strQ & "")

End Sub

Private Sub mnuEndic_Click()

'Call Shell("explorer ""http://endic.naver.com/search.naver?where=endic&mode=srch_ke&query=" + clipText + """", vbNormalFocus)

'wb.Navigate2 "http://endic.naver.com/search.naver?where=endic&mode=srch_ke&query=" + clipText

'Call make_html("http://endic.naver.com/search.naver?where=endic&mode=srch_ke&query=" + clipText, "endic.htm")
'wb.Navigate2 App.Path + "\endic.htm"


'Call make_html_pop(clipText, "tmp\$endic" + setRndName() + ".htm")
'frmMain.wb2.Navigate2 App.Path + "\tmp\$endic" + getRndName() + ".htm"

frmQuiz.wbQuizMain.Navigate2 ("http://endic.naver.com/popManager.nhn?sLn=kr&m=search&searchOption=&query=" + clipText)

'Call frmMain.characterPlay("Searching", "Search")

Call frmMain.characterPlay("")


End Sub
'영영사전
Private Sub mnu_ee_dic_Click()

Dim strResultURL As String
strResultURL = "http://www.collinsdictionary.com/dictionary/american-cobuild-learners/" + Replace(LCase$(clipText), " ", "_") + "?showCookiePolicy=true"
'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" " & strResultURL)
Call frmMain.characterPlay("")

Call wbQuizMain.Navigate(strResultURL)

End Sub


Private Sub mnuEngStudyPekr_Click()
'       [ ] { } ( ) * & ~ \
clipText = Replace(clipText, ";", "")
clipText = Replace(clipText, "^", "")
clipText = Replace(clipText, ",", "")
clipText = Replace(clipText, """", "")
clipText = Replace(clipText, "|", "")
clipText = Replace(clipText, "<", "")
clipText = Replace(clipText, ">", "")
clipText = Replace(clipText, "{", "")
clipText = Replace(clipText, "}", "")
clipText = Replace(clipText, "(", "")
clipText = Replace(clipText, ")", "")
clipText = Replace(clipText, "*", "")
clipText = Replace(clipText, "&", "")
clipText = Replace(clipText, "~", "")
clipText = Replace(clipText, "\", "")
clipText = Replace(clipText, "'", "")

    If Len(clipText) = LenH(clipText) Then
        
        Call wbQuizMain.Navigate("http://dic.impact.pe.kr/ecmaster-cgi/search.cgi?kwd=" & clipText)
    Else
        Call wbQuizMain.Navigate("http://dic.impact.pe.kr/ecmaster-cgi/search.cgi?kwd=" & clipText)
    End If
End Sub

Private Sub mnuGugleTransPage_Click()
Dim strQ As String
strQ = URLEncodeUTF8(clipText)
'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://translate.google.co.kr/?hl=ko&tab=wT")
Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" https://translate.google.com/translate_tts?ie=UTF-8&q=" & strQ & "&tl=ko-KR&client=tw-ob")
'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" " & Module1.TTS_GLOBAL_URL & "")
'Call Module1.toclipboard(clipText)

End Sub

'한자사전
Private Sub mnuHanjaDic_Click()
    Call make_html_pop_hanja(clipText, "tmp\$hanja_dic" + setRndName() + ".htm")
    frmMain.wb2.Navigate2 App.Path + "\tmp\$hanja_dic" + getRndName() + ".htm"
    'Call frmMain.characterPlay("Searching", "Search")
    
    
Call frmMain.characterPlay("")

End Sub
Private Sub mnuEndic2_Click()

'Call Shell("explorer ""http://www.ybmallinall.com/dic/dic_view.asp?range=SR&kwd=" + clipText + "&dict=dic_all&printdict=&dd_select=a&search_kwd=" + clipText + """", vbNormalFocus)

'wb.Navigate2 "http://www.ybmallinall.com/dic/dic_view.asp?range=SR&kwd=" + clipText + "&dict=dic_all&printdict=&dd_select=a&search_kwd=" + clipText

Call make_html("http://www.ybmallinall.com/dic/dic_view.asp?range=SR&kwd=" + clipText + "&dict=dic_all&printdict=&dd_select=a&search_kwd=" + clipText, "tmp\$endic1" + setRndName() + ".htm")
frmMain.wb2.Navigate2 App.Path + "\tmp\$endic1" + getRndName() + ".htm"

    'Call frmMain.characterPlay("Searching", "Search")
    
Call frmMain.characterPlay("")


End Sub

Private Sub mnuEndic3_Click()

'Call Shell("explorer ""http://www.ybmallinall.com/dic/mpplay.asp?dictnum=1&dictword=" + clipText + """", vbNormalFocus)

'wb.Navigate2 "http://www.ybmallinall.com/dic/mpplay.asp?dictnum=1&dictword=" + clipText

Call make_html("http://www.ybmallinall.com/dic/mpplay.asp?dictnum=1&dictword=" + clipText, "tmp\$endic3" + setRndName() + ".htm")
frmMain.wb2.Navigate2 App.Path + "\tmp\$endic3" + getRndName() + ".htm"

    'Call frmMain.characterPlay("Searching", "Search")
    
Call frmMain.characterPlay("")

End Sub

'영어사전[영국]
Private Sub mnuEndic4_Click()
'http://www.ybmallinall.com/Data/KESoundA/c/cite.mp3
'Call Shell("explorer ""http://www.ybmallinall.com/Data/KESoundA/" + Left(clipText, 1) + "/" + clipText + ".mp3""", vbNormalFocus)

'wb.Navigate2 "http://www.ybmallinall.com/Data/KESoundA/" + Left(clipText, 1) + "/" + clipText + ".mp3"

'Call make_html("http://www.ybmallinall.com/Data/KESoundA/" + Left(clipText, 1) + "/" + clipText + ".mp3", "tmp\$endic4" + setRndName() + ".htm")
'frmMain.wb2.Navigate2 App.Path + "\tmp\$endic4" + getRndName() + ".htm"

    'Call frmMain.characterPlay("Searching", "Search")
    

'    Call make_html_pop_ee(clipText, "tmp\$ee_dic" + setRndName() + ".htm")
'    frmMain.wb2.Navigate2 App.Path + "\tmp\$ee_dic" + getRndName() + ".htm"

'    frmMain.wb2.Navigate2 "http://www.collinsdictionary.com/dictionary/american-cobuild-learners/" + LCase$(clipText) + "?showCookiePolicy=true"


Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.collinsdictionary.com/dictionary/english-cobuild-learners/" + Replace(LCase$(clipText), " ", "_") + "?showCookiePolicy=true")
Call frmMain.characterPlay("")

End Sub

Private Sub mnuImageSearch_Click()
'이미지검색
Dim strQ As String
strQ = cQuiz.quiz ' clipText

If strQ = "cock" Then
strQ = "수탉" '야한 이미지가 많아 검색어 변경함
End If

If chk.Enabled Then
chk.Value = vbChecked
End If

Dim strResult As String
strResult = "http://images.google.co.kr/search?num=10&hl=ko&newwindow=1&site=&tbm=isch&source=hp&q=" + strQ + "&oq=" + strQ + ""
'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" " & strResult, vbMaximizedFocus)
'wbQuizMain.AddressBar = True
Call wbQuizMain.Navigate(strResult)

End Sub

Private Sub mnuInternetEx_Click()
Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe""")
End Sub

Private Sub mnuNaverDic_Click()
Call Shell("explorer http://dic.naver.com", vbMaximizedFocus)
Clipboard.Clear
Clipboard.SetText clipText
End Sub


Private Sub mnuRefresh_Click()
quizDisp False
End Sub

Public Sub mnuTTS0_Click()

Dim sLang As String

sLang = "ko-KR"

Dim clipText2 As String

clipText2 = Trim(clipText)
If clipText2 = "" Then Exit Sub

clipText2 = Replace(clipText2, "(A)", "")
clipText2 = Replace(clipText2, "(B)", "")
clipText2 = Replace(clipText2, "(C)", "")
clipText2 = Replace(clipText2, "(D)", "")
clipText2 = Replace(clipText2, vbCrLf, " ")
clipText2 = Replace(clipText2, "   ", " ")
clipText2 = Replace(clipText2, "  ", " ")

clipText2 = Replace(clipText2, "______", "_")
clipText2 = Replace(clipText2, "_____", "_")
clipText2 = Replace(clipText2, "____", "_")
clipText2 = Replace(clipText2, "___", "_")
clipText2 = Replace(clipText2, "__", "_")
'clipText2 = LeftH(clipText2, 200)

Dim nChkLenB As Long
Dim nChkLenH As Long


If gSetLang = "" Then
    nChkLenB = LenB(StrConv(clipText2, vbFromUnicode))
    nChkLenH = Len(clipText2)
    
    If (nChkLenB <> nChkLenH) Then
        'isHan = True
        sLang = "ko-KR" '?
    Else
        Dim char_first As String
        
        char_first = Left(clipText2, 1)
        
        Dim asc_first As Integer
        
        asc_first = Asc(char_first)
        
        If asc_first >= Asc("0") And asc_first <= Asc("9") Then
            sLang = "ko-KR" 'isHan = True
        Else
            'isHan = False
            sLang = "en-US"
        End If
        
    End If
Else
    sLang = gSetLang
End If

Dim tailSec As String
tailSec = Format(Now, "ss")

Static idx   As Long

idx = idx + 1
idx = idx Mod 3

Dim URL As String, url_hash As String



If sLang = "ko-KR" Then
    
    'Call make_html_tts_ko(clipText2, "tmp\$" + setRndName() & "tts_ko.htm")
    
    '일본어URL인지 확인한다.
    If IsJAJP(clipText2) Then
        'URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=ja&q=" + URLEncodeUTF8(clipText2)
        URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=ja-JP&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
        'URL = Module1.TTS_GLOBAL_URL
        'Call Module1.toclipboard(clipText2)
    Else
        'URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=ko&q=" + URLEncodeUTF8(clipText2)
        URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=ko-KR&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
        'URL = Module1.TTS_GLOBAL_URL
        'Call Module1.toclipboard(clipText2)
    End If
    Call TTS_PLAY_SMART3(URL)
    
'    Dim tmpPath As String
'    tmpPath = App.Path + "\tmp\$" & getRndName() & "tts_ko.htm"
        'Call toUTF8(tmpPath)
        ''frmMain.wb(idx).Navigate2 tmpPath
        'Call TTS_PLAY_SMART(tmpPath)
ElseIf sLang = "en-US" Then
    Dim bSuccessFromDaum As Boolean
    
    If 0 = InStr(clipText2, " ") Then
        If TTS_PLAY_SMART4(clipText2) Then
            bSuccessFromDaum = True
        End If
    End If
    
    If bSuccessFromDaum = False Then
        'URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=en&q=" + URLEncodeUTF8(clipText2)
        URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
        'URL = Module1.TTS_GLOBAL_URL
        'Call Module1.toclipboard(clipText2)
        Call TTS_PLAY_SMART3(URL)
    End If
    
ElseIf sLang = "zh-CN" Then
    'frmMain.wb(idx).Navigate2 "http://translate.google.com/translate_tts?ie=UTF-8&tl=zh&q=" + clipText2
    'Call TTS_PLAY_SMART3("http://translate.google.com/translate_tts?ie=UTF-8&tl=zh&q=" + URLEncodeUTF8(clipText2))
    Call TTS_PLAY_SMART3("https://translate.google.com/translate_tts?ie=UTF-8&tl=zh-CN&client=tw-ob&q=" + URLEncodeUTF8(clipText2))
'    Call TTS_PLAY_SMART3(Module1.TTS_GLOBAL_URL)
'    Call Module1.toclipboard(clipText2)
ElseIf sLang = "ja-JP" Then
    'frmMain.wb(idx).Navigate2 "http://translate.google.com/translate_tts?ie=UTF-8&tl=ja&q=" + clipText2
    'Call TTS_PLAY_SMART3("http://translate.google.com/translate_tts?ie=UTF-8&tl=ja&q=" + URLEncodeUTF8(clipText2))
    Call TTS_PLAY_SMART3("https://translate.google.com/translate_tts?ie=UTF-8&tl=ja-JP&client=tw-ob&q=" + URLEncodeUTF8(clipText2))
'    Call TTS_PLAY_SMART3(Module1.TTS_GLOBAL_URL)
'    Call Module1.toclipboard(clipText2)
End If

frmMain.wb2.Tag = clipText2
#If MYAGENTUSE_ON Then
If Not frmMain.Character Is Nothing Then
    frmMain.Character.Balloon.Style = frmMain.Character.Balloon.Style And (Not 4) '말풍선뜬거 계속 떠있게 하기
End If
#End If


'0.5초 후에 포커싱 하기
'''TmrAfterTTS_focus.Enabled = True
TmrAfterTTS_focus.Interval = 100

End Sub

Sub TTS_PLAY_SMART3(ByVal URL As String)
    If TTS_PLAY_SMART2(URL) = False Then
        Call TTS_PLAY_SMART(URL)
    End If
End Sub

Function TTS_PLAY_SMART4(ByVal str As String) As Boolean
    Dim FileName_tmp As String
    Dim FileName_tmp2 As String
    Dim FileName_ok As String
    Dim FileName_error As String
    
    FileName_tmp = FileSystem.CurDir & "\cache\temp_daum" + str + ".htm"
    FileName_tmp2 = FileSystem.CurDir & "\cache\temp_daum" + str + "2.htm"
    
    FileName_ok = FileSystem.CurDir & "\cache\" & str & ".mp3"
    FileName_error = FileSystem.CurDir & "\cache\~" & str & ".mp3"
    
    Dim okMP3 As Boolean
    
    okMP3 = False
    If FileExists(FileName_error) Then
        TTS_PLAY_SMART4 = False
        Exit Function
    End If
    
    If FileExists(FileName_ok) Then
        WindowsMediaPlayer1.settings.mute = False
        WindowsMediaPlayer1.settings.volume = 100
        WindowsMediaPlayer1.URL = FileName_ok
        TTS_PLAY_SMART4 = True
    Else
    
        Dim URL As String
        URL = "http://small.dic.daum.net/search.do?dic=eng&q=" & str
        
        Dim str1 As String
        str1 = searchFromUrlToContent(FileName_tmp, URL, "URL=", """", "wordid=")
        
        If str1 <> "" And InStr(str1, "wordid=") = 0 Then
            str1 = searchFromFileToContent(FileName_tmp, "<a href=""", """", "wordid=")
        Else
            Debug.Print "test"
        End If
        
        If (str1 <> "") Then
            DeleteFile (FileName_tmp)
            
            Dim str2 As String
            str2 = searchFromUrlToContent(FileName_tmp2, "http://dic.daum.net" + str1, "event, '", "'", "mainPlayer.play")
            If (str2 <> "") Then
                DeleteFile (FileName_tmp2)
                
                Call DownloadFile(str2, FileName_ok)
                
                DoEvents
                
                Call SleepEx(100, True)
                
                If FileExists(FileName_ok) Then
                    TTS_PLAY_SMART4 = True
                    WindowsMediaPlayer1.settings.mute = False
                    WindowsMediaPlayer1.settings.volume = 100
                    WindowsMediaPlayer1.URL = FileName_ok
                Else
                    Debug.Print "file not exits1?"
                End If
            Else
                        Debug.Print "file not exits2?"
            End If
        Else
                    Debug.Print "file not exits3?"
        End If
    End If
    
End Function

Function searchFromUrlToContent(ByVal FileName_tmp As String, ByVal URL As String, ByVal firstPosChar As String, ByVal lastPosChar As String, ByVal search_key As String) As String

'Dim URL As String
        'URL = "http://small.dic.daum.net/search.do?dic=eng&q=" & str
        
        Call DownloadFile(URL, FileName_tmp)
        
        DoEvents
        
        Call SleepEx(100, True)
        
        If FileExists(FileName_tmp) Then
            'wordid=ekw000167646&q=terrify
            Dim file_content As String
            file_content = searchLineInFile(FileName_tmp, search_key)
            'file_content = StrConv(StrConv(file_content, vbUnicode), vbFromUnicode)
            
            Dim pos1 As Long
            'Dim firstPosChar As String
            'firstPosChar = "URL="
            Dim firstPostCharLeng As Long
            firstPostCharLeng = Len(firstPosChar)
            'Dim lastPosChar As String
            'lastPosChar = """"
            
            pos1 = InStr(file_content, firstPosChar)
            Dim pos2 As Long
            pos2 = InStr(pos1 + firstPostCharLeng, file_content, lastPosChar)
            
            Dim keyword As String
            
            If 0 < pos2 - pos1 - firstPostCharLeng Then
                searchFromUrlToContent = Mid(file_content, pos1 + firstPostCharLeng, pos2 - pos1 - firstPostCharLeng)
            Else
                Debug.Print "not found"
            End If
            
            'http://dic.daum.net/word/view.do?wordid=ekw000167646&q=terrify
        End If
        

End Function

Function searchFromFileToContent(ByVal FileName_tmp As String, ByVal firstPosChar As String, ByVal lastPosChar As String, ByVal search_key As String) As String

    If FileExists(FileName_tmp) Then
        'wordid=ekw000167646&q=terrify
        Dim file_content As String
        file_content = searchLineInFileAbout30(FileName_tmp, search_key)
        'file_content = StrConv(StrConv(file_content, vbUnicode), vbFromUnicode)
        
        Dim pos1 As Long
        'Dim firstPosChar As String
        'firstPosChar = "URL="
        Dim firstPostCharLeng As Long
        firstPostCharLeng = Len(firstPosChar)
        'Dim lastPosChar As String
        'lastPosChar = """"
        
        pos1 = InStr(file_content, firstPosChar)
        Dim pos2 As Long
        pos2 = InStr(pos1 + firstPostCharLeng, file_content, lastPosChar)
        
        Dim keyword As String
        
        If 0 < pos2 - pos1 - firstPostCharLeng Then
            searchFromFileToContent = Mid(file_content, pos1 + firstPostCharLeng, pos2 - pos1 - firstPostCharLeng)
        Else
            Debug.Print "not found"
        End If
        
        'http://dic.daum.net/word/view.do?wordid=ekw000167646&q=terrify
    End If
        

End Function



Function searchLineInFile(ByVal f As String, ByVal search As String) As String

  Dim FileHandle As Integer
  FileHandle = FreeFile
  Open f For Input As #FileHandle

  Do While Not EOF(FileHandle)        ' Loop until end of file
    Line Input #FileHandle, searchLineInFile  ' Read line into variable
    If 0 < InStr(searchLineInFile, search) Then
        Exit Do
    End If
    ' Your code here
  Loop

  Close #FileHandle
  

End Function

'검색 성공후 앞에 30자리 이전은 모두 자른다.
Function searchLineInFileAbout30(ByVal f As String, ByVal search As String) As String

  Dim FileHandle As Integer
  Dim retVal As String
  Dim pos1 As Long
  
  FileHandle = FreeFile
  Open f For Input As #FileHandle

  Do While Not EOF(FileHandle)        ' Loop until end of file
    Line Input #FileHandle, retVal  ' Read line into variable
    
    pos1 = InStr(retVal, search)
    
    If 0 < pos1 Then
        
        If 30 < pos1 Then
            retVal = Mid(retVal, pos1 - 30)
        End If
        
        searchLineInFileAbout30 = retVal
        Exit Do
    End If
    ' Your code here
  Loop

  Close #FileHandle
  

End Function

Function TTS_PLAY_SMART2(ByVal URL As String) As Boolean '성공하면 true를 리턴

    Dim sha256 As New CSHA256
    Dim url_hash As String
    url_hash = sha256.sha256(URL) + ".mp3"

    Set sha256 = Nothing

    If FolderExists(FileSystem.CurDir & "\cache") = False Then
        On Error Resume Next
        FileSystem.MkDir (FileSystem.CurDir & "\cache") '//err.Number 75 가 발생되면서 Path/File access error 오류가 발생된다.
        If err.Number = 75 Then
            MsgBox "[" & FileSystem.CurDir & "\cache] 폴더를 수동으로 만드세요.", vbExclamation + vbOKOnly
        End If
    End If

    Dim FileName As String
    FileName = FileSystem.CurDir & "\cache\" & url_hash

    Dim okMP3 As Boolean
    
    okMP3 = False
    If FileExists(FileName) Then
        If 0 < FileSystem.FileLen(FileName) Then
            If isHTML(FileName) Then
                Call DeleteFile(FileName)
            Else
                okMP3 = True
            End If
        Else
            Call DeleteFile(FileName)
        End If
    End If
    

    If okMP3 Then
        'mp3 play
       'Call mciSendString(CommandString, vbNullString, 0, 0)
       WindowsMediaPlayer1.settings.mute = False
       WindowsMediaPlayer1.settings.volume = 100
       WindowsMediaPlayer1.URL = FileName

       TTS_PLAY_SMART2 = True
    Else
       'http://stackoverflow.com/questions/1976152/
       Call DownloadFile(URL, FileName)
       '"C:\Program Files\Internet Explorer\iexplore.exe"
       'Call Shell("C:\Program Files\Internet Explorer\iexplore.exe" & " " & url)

       DoEvents
       
       TTS_PLAY_SMART2 = False
       
       If FileExists(FileName) Then
            If 0 < FileSystem.FileLen(FileName) Then
                If isHTML(FileName) Then
                    Call DeleteFile(FileName)
                Else
                    TTS_PLAY_SMART2 = True
                End If
            Else
                Call DeleteFile(FileName)
            End If
       End If
       
       If TTS_PLAY_SMART2 Then
            WindowsMediaPlayer1.settings.mute = False
            WindowsMediaPlayer1.settings.volume = 100
            WindowsMediaPlayer1.URL = FileName
       End If
    End If
    
End Function

Function IsJAJP(ByVal txt As String) As Boolean
    '%aa%a1~%abf6까지가 일본어
    Dim str1 As String
    str1 = "&H" & Replace(Encode(Left(txt, 1)), "%", "")
    
    Dim v As Long
    
    v = CLng(str1)
    
    If CLng("&Haaa1") <= v And v <= CLng("&Habf6") Then
        IsJAJP = True
    End If
    
End Function

'mp3 파일에도 없고. URL을 호출하여도 응답이 없는 경우는 직접 해당 URL을 메인 웹브라우저에 표시한다.
Sub TTS_PLAY_SMART(ByVal URL As String)
'Call wbQuizMain.Navigate(URL)'그런데 의미 없어 주석 처리함
End Sub

Function isHTML(ByVal filepath) As Boolean
    Dim filetext As String
    Dim handle As Integer
    Dim Pos As Long
    
    handle = FreeFile
    On Error Resume Next
    Open filepath For Input As #handle
               
        Do While Not EOF(handle)
          Line Input #handle, filetext
          If filetext Like "*<html*" Then
            Pos = 1
            Exit Do
          End If
        Loop
        
    Close #handle
    On Error GoTo 0
    '
    If 0 < Pos Then
     isHTML = True
    End If
    
End Function

Public Function FileExists(ByVal fname As String) As Boolean
    Dim TheFile As String
    Dim Results As String
    
    TheFile = fname
    Results = Dir$(TheFile)
    
    If Len(Results) = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function




Private Sub mnuVideoSearch_Click()
'http://www.google.com/search?tbm=vid&hl=ko&source=hp&biw=1302&bih=789&q=%ED%99%94%EC%82%B0&gbv=2&oq=%ED%99%94%EC%82%B0

Dim strQ As String
strQ = clipText

Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.google.com/search?tbm=vid&hl=ko&source=hp&biw=1302&bih=789&q=" + strQ + "&oq=" + strQ + "", vbMaximizedFocus)

End Sub


Private Sub mnuVideo_Click()
Dim strQ As String
strQ = clipText


If strQ = "cock" Then
strQ = "수탉" '야한 이미지가 많아 검색어 변경함
End If

If chk.Enabled Then
chk.Value = vbChecked
End If

Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.google.com/search?tbm=vid&hl=ko&source=hp&biw=1302&bih=789&q=" + strQ + "&oq=" + strQ + "", vbMaximizedFocus)

End Sub

Private Sub mnuGoogleSearch_Click()
Dim strQ As String
strQ = URLEncodeUTF8(clipText)
Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.google.co.kr/#hl=ko&newwindow=1&output=search&sclient=psy-ab&q=" + strQ + "&oq=" + strQ + "")

End Sub

Private Sub mnuNaverSearch_Click()
'http://search.naver.com/search.naver?where=nexearch&query=%EB%B0%95%EC%A0%95%ED%9D%AC&sm=top_hty&fbm=1&ie=utf8
Dim strQ As String
strQ = URLEncodeUTF8(clipText)

If LenB(strQ) = LenH(clipText) Then
    wbQuizMain.Navigate ("http://search.naver.com/search.naver?where=nexearch&query=" + strQ + "&sm=top_hty&fbm=1&ie=utf8")
Else
    Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://search.naver.com/search.naver?where=nexearch&query=" + strQ + "&sm=top_hty&fbm=1&ie=utf8")
End If

End Sub

Private Sub mnuDaumSearch_Click()
'http://search.daum.net/search?w=tot&DA=YZRR&t__nil_searchbox=btn&sug=&q=%EA%B2%B0%ED%98%BC
Dim strQ As String
strQ = URLEncodeUTF8(clipText)
Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://search.daum.net/search?w=tot&DA=YZRR&t__nil_searchbox=btn&sug=&q=" + strQ + "")

End Sub


Private Sub mnuWiki100_Click()
'http://ko.wikipedia.org/w/index.php?search=%EC%9D%98%EB%82%98%EB%AC%B4&title=%ED%8A%B9%EC%88%98%EA%B8%B0%EB%8A%A5%3A%EC%B0%BE%EA%B8%B0
Dim strQ As String

If isEng(clipText) Then
strQ = URLEncodeUTF8(clipText)
    strQ = "chair"
    Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://en.wikipedia.org/w/index.php?search=" & strQ & "&title=" & strQ & "")
Else
    strQ = URLEncodeUTF8(clipText)
    'strQ = URLEncodeUTF8("칼세이건")
    Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://ko.wikipedia.org/w/index.php?search=" & strQ & "&title=" & strQ & "")
End If
End Sub

Private Sub mnuWordListenENUS_Click()

'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.collinsdictionary.com/sounds/e/en_/en_us/en_us_" + LCase(cQuiz.quiz) + ".mp3")

Dim strQ As String
Dim strDoc As String
Dim pos1 As Long
Dim pos2 As Long

strQ = LCase(cQuiz.quiz)

frmMain.wb2.Silent = True 'script오류를 무시한다.

frmMain.wb2.Navigate2 "http://www.collinsdictionary.com/dictionary/american-cobuild-learners/" + Replace(LCase$(clipText), " ", "_") + "?showCookiePolicy=true"

While frmMain.wb2.Busy
    DoEvents
Wend

DoEvents
Call SleepEx(500, True) '0.1초 대기하여 페이지 내에서 리다이렉트하여 새로운 페이지로 바뀌기 까지 다시 기다린다.
DoEvents

strDoc = frmMain.wb2.Application.Document.documentElement.innerhtml

pos1 = InStr(strDoc, "playSoundFromFlash")
If pos1 = 0 Then
    Call SleepEx(500, True)
    DoEvents
    strDoc = frmMain.wb2.Application.Document.documentElement.innerhtml
End If

If pos1 = 0 Then
    Exit Sub '영어단어가 아닌 경우 오류난다.
End If

strDoc = Mid(strDoc, pos1, 256)
pos1 = InStr(strDoc, ", '")
pos2 = InStr(strDoc, ".mp3")

strDoc = Mid(strDoc, pos1 + 3, pos2 - pos1 + 1)
Dim errCnt As Long
While 0 = InStr(strDoc, strQ) And errCnt < 10
errCnt = errCnt + 1
Call SleepEx(100, True)
DoEvents
strDoc = frmMain.wb2.Application.Document.documentElement.innerhtml
pos1 = InStr(strDoc, "playSoundFromFlash")
If 0 < pos1 Then
strDoc = Mid(strDoc, pos1, 256)
pos1 = InStr(strDoc, ", '")
pos2 = InStr(strDoc, ".mp3")
strDoc = Mid(strDoc, pos1 + 3, pos2 - pos1 + 1)
Else
    strDoc = ""
End If
Wend

If 0 < InStr(strDoc, strQ) Then
frmMain.wb3.Silent = True
'frmMain.wb3.Navigate2 "http://www.collinsdictionary.com" + strDoc
Call TTS_PLAY_SMART("http://www.collinsdictionary.com" + strDoc)
'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.collinsdictionary.com" + strDoc)
Else
    If MsgBox("재시도하시겠습니까?", vbQuestion + vbOKCancel, App.Title) = vbOK Then
        
    strDoc = frmMain.wb2.Application.Document.documentElement.innerhtml
    pos1 = InStr(strDoc, "playSoundFromFlash")
    strDoc = Mid(strDoc, pos1, 256)
    pos1 = InStr(strDoc, ", '")
    pos2 = InStr(strDoc, ".mp3")
    strDoc = Mid(strDoc, pos1 + 3, pos2 - pos1 + 1)
        If 0 < InStr(strDoc, strQ) Then
            'frmMain.wb3.Navigate2 "http://www.collinsdictionary.com" + strDoc
            Call TTS_PLAY_SMART("http://www.collinsdictionary.com" + strDoc)
            'Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.collinsdictionary.com" + strDoc)
        End If
    End If
End If

End Sub

Private Sub mnuYoutubeSearch_Click()


Dim strQ As String
strQ = URLEncodeUTF8(clipText)

If strQ = "cock" Then
strQ = "수탉" '야한 이미지가 많아 검색어 변경함
End If

If chk.Enabled Then
chk.Value = vbChecked
End If

Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://www.youtube.com/results?search_query=" + strQ)

End Sub

Private Sub mnuZum100_Click()

Dim strQ As String
strQ = URLEncodeUTF8(clipText)
Call Shell("""C:\Program Files\Internet Explorer\iexplore.exe"" http://search.zum.com/search.zum?method=uni&query=" & strQ & "&qm=f_typing.top")

End Sub

'Private Sub optA_Click()
'cmdNext_Click
'End Sub

Private Sub optA_DblClick()
cmdNext_Click
End Sub

Private Sub optA_GotFocus()

'frmQuiz.SetFocus
picMain.SetFocus
End Sub

Private Sub optA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    lastSelExample = "a"
End If
End Sub

Private Sub optA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static updown_cnt As Integer
Static leftright_cnt As Integer
Static signY As Integer
Static signX As Integer
Static lastY As Single
Static lastX As Single
Static bRunning As Boolean

If bAutoClick Or bRunning Then
    Call motionCatch(1, updown_cnt, leftright_cnt, signY, signX, lastY, lastX, bRunning, X, Y)
End If


End Sub

Private Sub optA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    X1 = X / Screen.TwipsPerPixelX
    Y1 = Y / Screen.TwipsPerPixelY

    If lastSelExample = "a" And X >= 0 And X <= optA.Width And Y >= 0 And Y <= optA.Height Then
        optA.Value = True
'        Call FocusRect1(1)
        cmdNext_Click
    Else
        optA.Value = False
        
    End If
    optA.Value = False
End If
End Sub

Private Sub optAutoClickOff_Click()
If bAutoClick = False Then Exit Sub
bAutoClick = False
optAutoClickOff.Value = True

End Sub


Private Sub optAutoClickOn_Click()
If bAutoClick Then Exit Sub
bAutoClick = True
optAutoClickOn.Value = True

End Sub


Private Sub optB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    lastSelExample = "b"
End If
If optB.Value = True Then optB.Value = False
End Sub

Private Sub optB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static updown_cnt As Integer
Static leftright_cnt As Integer
Static signY As Integer
Static signX As Integer
Static lastY As Single
Static lastX As Single
Static bRunning As Boolean

If bAutoClick Or bRunning Then
    Call motionCatch(2, updown_cnt, leftright_cnt, signY, signX, lastY, lastX, bRunning, X, Y)
End If
End Sub

Private Sub optB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    
    
        X1 = X / Screen.TwipsPerPixelX
        Y1 = Y / Screen.TwipsPerPixelY
    
        If lastSelExample = "b" And X >= 0 And X <= optB.Width And Y >= 0 And Y <= optB.Height Then
            optB.Value = True
'            Call FocusRect1(2)
            cmdNext_Click
        Else
            optB.Value = False
        End If
        
        optB.Value = False
   
End If
End Sub

Private Sub optC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    lastSelExample = "c"
End If
If optC.Value = True Then optC.Value = False
End Sub

Private Sub optC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static updown_cnt As Integer
Static leftright_cnt As Integer
Static signY As Integer
Static signX As Integer
Static lastY As Single
Static lastX As Single
Static bRunning As Boolean

If bAutoClick Or bRunning Then
    Call motionCatch(3, updown_cnt, leftright_cnt, signY, signX, lastY, lastX, bRunning, X, Y)
End If


End Sub

Private Sub optC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
   
    X1 = X / Screen.TwipsPerPixelX
    Y1 = Y / Screen.TwipsPerPixelY

    If lastSelExample = "c" And X >= 0 And X <= optA.Width And Y >= 0 And Y <= optA.Height Then
        optC.Value = True
'        Call FocusRect1(3)
        cmdNext_Click
            
    Else
        optC.Value = False
    End If
    optC.Value = False
End If
End Sub

Private Sub optD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    lastSelExample = "d"
End If
End Sub

Private Sub optD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static updown_cnt As Integer
Static leftright_cnt As Integer
Static signY As Integer
Static signX As Integer
Static lastY As Single
Static lastX As Single
Static bRunning As Boolean

If bAutoClick Or bRunning Then
    Call motionCatch(4, updown_cnt, leftright_cnt, signY, signX, lastY, lastX, bRunning, X, Y)
End If

End Sub


Static Sub motionCatch(idx As Integer, ByRef updown_cnt As Integer, _
    ByRef leftright_cnt As Integer, _
    ByRef signY As Integer, _
    ByRef signX As Integer, _
    ByRef lastY As Single, _
    ByRef lastX As Single, _
    ByRef bRunning As Boolean, _
    ByRef X As Single, _
    ByRef Y As Single)

Static bRun As Boolean
Static lastIdx As Integer
Static lastTimer As Double

If lastIdx <> idx Then
    updown_cnt = 0
    leftright_cnt = 0
Debug.Print "상하좌우 초기화    "
ElseIf Timer > lastTimer + 1 Then '1초이상 머뭇거리면 초기화
    updown_cnt = 0
    leftright_cnt = 0
Debug.Print "상하좌우 초기화    "
End If
lastIdx = idx
lastTimer = Timer

If bRun Then Exit Sub
bRun = True

bRunning = True '함수가 static이 아니면 함수 종료시점에 이 값이 설정된다.

    Dim deltaX As Single
    Dim deltaY As Single
    
    Dim opt As OptionButton
    
Debug.Print idx, X, Y, iCatch, optA.Height / 5, Y - lastY, X - lastX
    
    deltaY = Y - lastY
    deltaX = X - lastX
    
    If Abs(deltaY) < (optA.Height / 6) And Abs(deltaX) < (optA.Height / 6) Then
        bRun = False
        Exit Sub
    Else

    End If
    
    If idx = 1 Then
        Set opt = optA
    ElseIf idx = 2 Then
        Set opt = optB
    ElseIf idx = 3 Then
        Set opt = optC
    ElseIf idx = 4 Then
        Set opt = optD
    ElseIf idx = 5 Then
        Set opt = optE
    End If
        
'Debug.Assert Abs(deltaX) < (optA.Height / 5)
        
    If Abs(deltaY) >= Abs(optA.Height / 6) Then
        If deltaY < 0 Then 'down
            If signY <> -1 Then 'not befor moving down
                updown_cnt = updown_cnt + 1
                signY = -1
Debug.Print "상하방향증가", updown_cnt
            End If
            lastY = Y
        ElseIf deltaY > 0 Then 'move up
            If signY <> 1 Then
                updown_cnt = updown_cnt + 1
                signY = 1
Debug.Print "상하방향증가", updown_cnt
            End If
            lastY = Y
        End If
        
    Else 'If Abs(deltaX) >= Abs(optA.Height /6) Then
        If deltaX < 0 Then 'left
            If signX <> -1 Then
                leftright_cnt = leftright_cnt + 1
                signX = -1
Debug.Print "좌우방향증가", leftright_cnt
            End If
            lastX = X
        Else 'left
            If signX <> 1 Then
                leftright_cnt = leftright_cnt + 1
                signX = 1
Debug.Print "좌우방향증가", leftright_cnt
            End If
            lastX = X
        End If
    
'    Else
'        Debug.Assert False
    End If
    
'    lastX = X
'    lastY = Y
    
    If updown_cnt >= 2 Then
        If opt.Value = False Then
            opt.Value = True
            Call FocusRect1(idx)
            iCatch = idx
        ElseIf iCatch <> idx Then
            iCatch = idx
        Else
'            bRunning = True
Debug.Print "다음버튼클릭!", idx, "선택"
            cmdNext_Click
'            lastY = optA.Height / 2
            iCatch = 0
        End If
        updown_cnt = 0
        leftright_cnt = 0
    ElseIf leftright_cnt >= 2 Then
        If opt.Value = False Then
            opt.Value = True
            Call FocusRect1(idx)
        ElseIf iCatch <> idx Then
            iCatch = idx
        Else
'            bRunning = True
Debug.Print "다음버튼클릭!", idx, "선택"
            cmdNext_Click
'            lastX = optA.Width / 2
            iCatch = 0
        End If
        leftright_cnt = 0
        updown_cnt = 0
    End If
bRunning = False

bRun = False
End Sub



Private Sub optD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    X1 = X / Screen.TwipsPerPixelX
    Y1 = Y / Screen.TwipsPerPixelY

    If lastSelExample = "d" And X >= 0 And X <= optA.Width And Y >= 0 And Y <= optA.Height Then
        optD.Value = True
'        Call FocusRect1(4)
        cmdNext_Click
    Else
        optD.Value = False
        
   End If
    optD.Value = False
End If
End Sub

Private Sub optE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    lastSelExample = "e"
End If

End Sub

Private Sub optE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static updown_cnt As Integer
Static leftright_cnt As Integer
Static signY As Integer
Static signX As Integer
Static lastY As Single
Static lastX As Single
Static bRunning As Boolean

If bAutoClick Or bRunning Then
    Call motionCatch(5, updown_cnt, leftright_cnt, signY, signX, lastY, lastX, bRunning, X, Y)
End If


End Sub

Private Sub optE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    X1 = X / Screen.TwipsPerPixelX
    Y1 = Y / Screen.TwipsPerPixelY

    If lastSelExample = "e" And X >= 0 And X <= optA.Width And Y >= 0 And Y <= optA.Height Then
        optE.Value = True
'        Call FocusRect1(5)
        cmdNext_Click
    Else
        optE.Value = False
        
   End If
   
   optE.Value = False
   
End If
End Sub

Private Sub optB_DblClick()
cmdNext_Click
End Sub

Private Sub optC_DblClick()
cmdNext_Click
End Sub

Private Sub optD_DblClick()
cmdNext_Click
End Sub

Private Sub optE_DblClick()
cmdNext_Click
End Sub

Private Sub optTTSA_Click()
clipText = cQuiz.a
    Call mnuTTS0_Click
End Sub

Private Sub optTTSB_Click()
clipText = cQuiz.B
    Call mnuTTS0_Click
End Sub

Private Sub optTTSC_Click()
clipText = cQuiz.C
    Call mnuTTS0_Click
End Sub

Private Sub optTTSD_Click()
clipText = cQuiz.d
    Call mnuTTS0_Click
End Sub

Private Sub optTTSE_Click()
clipText = cQuiz.e
    Call mnuTTS0_Click
End Sub

Private Sub optTTSQuiz_Click()
    clipText = cQuiz.quiz
    Call mnuTTS0_Click
End Sub

Private Sub PicMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button And vbLeftButton) = vbLeftButton Then
    If (Y < pic1.Top) Then
        optA.Value = False: optB.Value = False: optC.Value = False: optD.Value = False
    End If
    bonDraw = True
    picMain.DrawWidth = cFTP.Profile.FontSize / 5
    
    Call PicMainDrawModeForeColor
    'picMain.ForeColor = vbMagenta
    
    picMain.Line (X, Y)-(X, Y)

Else
    bonDraw = False
    picMain.DrawMode = 13
    clipText = cQuiz.quiz
    Me.PopupMenu mnuPop1
End If

End Sub

Sub PicMainDrawModeForeColor()

    If 128 < GrayScaleFromColor(cFTP.Profile.PenColor) Then
        picMain.DrawMode = 12
        picMain.ForeColor = Shape2.FillColor
    Else
        picMain.DrawMode = 13
        picMain.ForeColor = Shape2.FillColor Xor &HFFFFFF
    End If

End Sub

Private Function GrayScaleFromColor(ByVal lPenColor As Long) As Long

    Dim red As Long, green As Long, blue As Long, Sum As Long
    
    red = RedFromRGB(lPenColor)
    green = GreenFromRGB(lPenColor)
    blue = BlueFromRGB(lPenColor)

    GrayScaleFromColor = Fix(red * 0.3 + green * 0.6 + 0.1 * blue)

End Function


Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then Exit Sub
Static preX As Single
Static preY As Single

Dim Dist As Single

'Call opt_MouseOut

If ((bonDraw And Button) = vbLeftButton) And Shift = 0 Then  '왼쪽버튼 클릭
    
    If preX = 0 And preY = 0 Then
        preX = X
        preY = Y
        Exit Sub
    End If
    Dist = ((X - preX) ^ 2 + (Y - preY) ^ 2) ^ 0.5 '[twip]
    
    If Dist > picMain.DrawWidth * Screen.TwipsPerPixelX Then
    
        picMain.Line -(X, Y)
        preX = X
        preY = Y
    End If
End If

End Sub

Private Sub PicMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bonDraw = False
picMain.DrawMode = 13

End Sub


Private Sub Timer1_Timer()

If pgBar.Value > 10 Then
    Timer1.Interval = 40 '1000
Else
    Timer1.Interval = 40
End If

If Me.Visible = False Then
    Timer1.Enabled = False
    Exit Sub
End If
If pgBar.Value < Trim(Timer1.Interval) / 1000 Then
    pgBar.Value = Trim(Timer1.Interval) / 1000
End If

pgBar.Value = Trim(pgBar.Value) - Trim(Timer1.Interval) / 1000
txtTimer.Text = Fix(pgBar.Value) + 1
If pgBar.Value = 0 Then
    If gMutex Then
        Exit Sub '절묘한 시간에 [이전]버튼을 눌렀는데 이곳을 진입하는 경우가 발생된 사례가 있다.
        '쿼리문 수행 중 doevent 문에서 브랜치되어 타이머가 작동된 경우인데 이경우
        '다시 쿼리문을 수행하기 위한 아래가 수행될 수 있다. 이때는 타이머가 수행되면 안된다.
        '왜냐하면 이전으로 가기 했는데 타이머 작동이 이후 바로되어 이전으로 가지 않는경우도
        '있을 수 있기 때문이다.
    End If
    Timer1.Enabled = False
    cmdNext_Click
End If
'picmain.SetFocus
End Sub


Public Function Read_TU03() As Long
Dim depthAll As Long
Dim depthnow As Long

Con_Open
sSql = "select lastnew,hangsu,od from tu03 where userid='" & gUserid & "' and POCKETNM='" & gPocket & "' AND CHASU=" & gChasu
Set rs = Fn_SQLExec(sSql, , , True).rs
If rs.EOF Then
    MsgBox "해당 시험지는 삭제 되어야 합니다.(상세정보없음)", vbCritical
    Con_Close
    gbReadTu03 = False
    Exit Function
End If
gLastNew = rs(0)
gHangSu = rs(1)
gOrder = rs(2)
If gOrder < 2 Then
    gOrder = 2
End If
Read_TU03 = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow)

'Debug.Assert gLastNew <> 100
Con_Close
gbReadTu03 = True
End Function

Public Function Read_TU02() As Boolean
   'tu02 사용 안함.
   
    If cQuiz.reserve_ymd = "99999999" Then
        cQuiz.reserve_ymd = GETYMD()
    End If
End Function

Sub Write_TU03()
If gbReadTu03 Then
    Con_Open
    sSql = "update tu03 set lastnew=" & gLastNew & " ,  hangsu=" & gHangSu & " ,od=" & gOrder & " where pocketnm='" & gPocket & "' and userid='" & gUserid & "' and chasu=" & gChasu
    Fn_SQLExec (sSql)
    Con_Close
End If
End Sub
Public Function quizDisp(Optional ForHint As Boolean = False, Optional ForDel As Boolean = False) As Boolean
Dim vHint As Variant
Dim maxNum As Variant
Dim i As Long
Dim h0 As Long
Dim h1 As Long
Dim minF As Long
Dim notExistList As Collection
Static inc1 As Long
On Error GoTo ErrTrap

mnu_auto_tts.Checked = cFTP.Profile.bChkTTSuse

If nowVisible Then Exit Function

'If gbMnuExamClick = True And ForHint = False Then
'    Exit Function
'End If

#If MYAGENTUSE_ON Then
If Not frmMain.Character Is Nothing Then
    frmMain.Character.Balloon.Visible = False
    frmMain.Character.Stop
End If
#End If

If cQuiz.bPass Then  'bPass를 false로 바꾼다. 단 2번째로 이곳을 타는경우이다.
    inc1 = inc1 + 1
    If inc1 > 2 Then
        cQuiz.bPass = False
    End If
Else
    inc1 = 0
End If

bonDraw = False '더이상 줄을 긋지 않기 위해
lastSelExample = ""

If cQuiz.forReview = False And ForHint = False And gMainOnResize = False Then
    cQuiz.a = ""
    cQuiz.B = ""
    cQuiz.C = ""
    cQuiz.d = ""
    cQuiz.e = ""
    cQuiz.hint = ""
    Call FocusRect1(0)
End If

'If gIsNew Then
'    optA.Enabled = False
'    optB.Enabled = False
'    optC.Enabled = False
'    optD.Enabled = False
'    optE.Enabled = False
'
'Else
'
'    optA.Enabled = True
'    optB.Enabled = True
'    optC.Enabled = True
'    optD.Enabled = True
'    optE.Enabled = True
'
'End If

If gMainOnResize = False Then
optA.Value = False
optB.Value = False
optC.Value = False
optD.Value = False
optE.Value = False

optA.Enabled = True
optB.Enabled = True
optC.Enabled = True
optD.Enabled = True



optA.Visible = True
optB.Visible = True
'optC.Visible = True
'optD.Visible = True


'cmdHint.Visible = False

cmdRef.Visible = False


optA.FontBold = cFTP.Profile.FontBold
optB.FontBold = cFTP.Profile.FontBold
optC.FontBold = cFTP.Profile.FontBold
optD.FontBold = cFTP.Profile.FontBold
optE.FontBold = cFTP.Profile.FontBold
picMain.FontBold = cFTP.Profile.FontBold
If cFTP.Profile.FontName <> "" Then
    err.Clear
    On Error Resume Next
    picMain.FontName = cFTP.Profile.FontName
    pic1.FontName = cFTP.Profile.FontName
    On Error GoTo 0
    
End If


'subj와 seq값을 얻어옴

'With cQuiz

cQuiz.num = gNowNum
cQuiz.isNew = gIsNew

End If

If cTU01.BreakedLogOn(gSessionId) Then
    MsgBox "연결이 원활하지 못하여 프로그램이 종료됩니다.", vbExclamation, "포켓퀴즈"
    Unload fMainForm
    End
End If

If ((cQuiz.forReview = False) Or cQuiz.isBefore) And gMainOnResize = False Then '전으로 버튼을 클릭한경우(isBefore=true) 인경우도 새로 디스플레이한다.
Con_Open


sSql = "select subj,seq from tp03 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num=" & gNowNum
Set rs = Fn_SQLExec(sSql, , , True).rs

If rs.State = 0 Then
    '위와 같은 경우
    'rs가 닫힌 rs로 리턴되는 경우가 있다.
    '이경우 프로시저를 빠져 나가야 한다.
    Exit Function
End If

'다른곳에서 시험지를 지워서 없을수 있음이 보고됨.
'Debug.Assert rs.RecordCount = 1'카운팅 옵션을 true로 했으므로 카운팅이 안됨.

Dim depthAll As Long
Dim depthnow As Long
Dim befNowNum As Long
Dim preDepthnow As Long
'Dim tmpLastNew As Long
Dim tmpLastNew2 As Long
Dim bTmp As Boolean
befNowNum = gNowNum
If rs.EOF Then
Screen.MousePointer = vbHourglass

While rs.EOF

'    Debug.Assert False
'    MsgBox " 시험을 완료하였습니다."
'    gNowNum = gNowNum + 1
    cQuiz.overFlow = True
    sSql = "select max(num) from tp03 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & ""
    'sSql = "select count(num) from tp03 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & ""
    
    'todo 바로 rs를 실행하지 말고 변수로 받아 체크할 것!
    
    
    maxNum = Fn_SQLExec(sSql).rs(0)
    If Not IsNumeric(maxNum) Then   '삭제된 시험지.
        Call MsgBox("이미 삭제된 시험지 입니다. 시험을 종료합니다.", vbExclamation + vbOKOnly, "삭제된 시험지 알림")
        
        bTmp = cFTP.Profile.setPoP3
        cFTP.Profile.setPoP3 = False
        cmdClose_Click
        cFTP.Profile.setPoP3 = bTmp
        
        Exit Function
    End If
    
    
    Do
        If cQuiz.isBefore Then
            gHangSu = gHangSu - 1
            Debug.Assert gHangSu >= 0
        Else
            gHangSu = gHangSu + 1
        End If
        
        
        gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow, h0, h1)
        
        If gNowNum <= maxNum * gOrder * 2 Then
            While gNowNum > maxNum
                If cQuiz.isBefore Then
                    gHangSu = gHangSu - 1 '1씩빼는 것보다 더 빠른 방법이 있을 듯함.(하지만 maxNum중간에 언제라도 누락된 문제가 있을수있으므로 그냥 1씩 하는게 나을듯함)
                    Debug.Assert gHangSu >= 0
                Else
                    gHangSu = gHangSu + 1 '1씩더하는 것보다 더 빠른 방법이 있을 듯함.(하지만 maxNum중간에 언제라도 누락된 문제가 있을수있으므로 그냥 1씩 하는게 나을듯함)
                    cQuiz.overFlow = True
                End If
                gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow, h0, h1)
            Wend
        Else
            If cQuiz.isBefore Then
                gHangSu = gHangSu - 1
                Debug.Assert gHangSu >= 0
            Else
                gHangSu = gHangSu + 1
            End If
                
            gNowNum = (gHangSu - 1) Mod maxNum
            gNowNum = gNowNum + 1
            If gNowNum = 1 Then
                cQuiz.overFlow = True
            End If
        End If
        
        gIsNew = False '무조건 false이다.
        If gLastNew > (maxNum * 10000#) Or gLastNew > 13421772 Then '(maxNum * 1000) 조건 추가 20050413
                    Screen.MousePointer = vbDefault
                    Me.Hide
                    MsgBox "다음문제의 번호가 너무 커졌습니다.          " + vbCrLf + vbCrLf + "시험을 종료합니다.", vbExclamation
                    Unload Me
                    Write_TU03
                    Exit Function
        End If
                
        Exit Do '여기를 타게 될거므로 여기서 빠져나가고 대신 mod를 수행하였다.
        
        If gNowNum <= maxNum Then
            SSQL1 = "select count(*) from tp03 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num>=" & gNowNum
            Set rs = Fn_SQLExec(SSQL1).rs
            If rs(0) > 0 Then
                Exit Do '여기를 타게 될거므로
            Else
                Debug.Assert False '문제가 띄엄띄엄 만들었을 경우 이다.
            End If
        Else
            '만약 출력할 번호가 너무큰번호면 출력을 못하므로 앞으로 혹은 뒤로 계속이동하면서 적절히 출력할
            '번호를 찾아야 하는데 뒤로가기가 문제가있다.
            If cQuiz.isBefore Then
            
            Else
            
            End If
            
            Do Until gNowNum <= maxNum '종료조건임에 유의
                'gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow, h0, h1)
                If depthAll = depthnow Then
                    'glastnew mod gorder
                    If depthAll = 0 Then
                        If cQuiz.isBefore Then
                            gLastNew = maxNum + 1 '20051003 추가한것인데 정확한 로직연구없이 마구잡이로 코딩한경우임
                        Else
                            gLastNew = gLastNew + (gOrder - (gLastNew Mod gOrder))
                        End If
                    Else
                        Debug.Assert False 'depthall=0 이 아닌경우는? 아직 미구현
'_/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/
'
'  이곳을 진입한경우의 의미
'  depthAll= 7 인경우 해당문제를 해당Order에서 7번+1번 반복하는 경우이며 순차진입이 아닌
'  계산에 의해서 진입되는 경우이다.
'  이경우 올바른 항수가 계산되어지는지는 그리고 Order변화에 대한 대응이 올바른지 검증이 되지 않았다.
'
'_/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/
                    End If
                    
                Else
'                    h0 = h0
                    minF = gLastNew - gOrder ^ depthAll + 1
                    If minF > maxNum Then
                        'gLastNew = gLastNew + 1
                        '1씩 LastNew를 증가시키면 시간이 많이 걸리므로 보완함
                        '20050413
                        '가령 LastNew=6 이어서 depthAll=1 이어서 minF=5 이고 maxNum=2 여서 minF>maxNum 인 조건인경우
                        'order^depthall 한값이 glastnew보다 작거나 같으면
                        'gLastNew값을 gorder^(depthall+1) 한값으로 설정하고
                        '그래도 gLastNew값이 maxNum보다 작거나 같으면
                        'gLastNew값을 gorder^(depthall+2) 한값으로 설정하고
                        '그래도 gLastNew값이 maxNum보다 작거나 같으면 위와같이 반복해서 승수를 증가시키는 방법 으로 하여 gLastNew가 8 이나오게 한다음
                        '아래 20 라인 밑에서는 소그룹첫번째 행번호를찾아가게 구현함 20050413
                        
                        tmpLastNew2 = gOrder ^ depthAll
                        i = 0
                        While tmpLastNew2 <= gLastNew
                            i = i + 1
                            tmpLastNew2 = gOrder ^ (depthAll + i)
                        Wend
                        
                        gLastNew = tmpLastNew2
                        
                    Else
                        'Debug.Assert False '20051003
                    End If
                End If



                If gLastNew > (maxNum * 10000#) Or gLastNew > 13421772 Then '(maxNum * 1000) 조건 추가 20050413
                    Screen.MousePointer = vbDefault
                    Me.Hide
                    MsgBox "다음문제의 번호가 너무 커졌습니다.          " + vbCrLf + vbCrLf + "시험을 종료합니다.", vbExclamation
                    Unload Me
                    Write_TU03
                    Exit Function
                Else
                    If cQuiz.isBefore Then
                        gHangSu = gHangSu - 1
                        
                        Debug.Assert gHangSu >= h0
                        
                        If gHangSu < h0 Then
                            gLastNew = gLastNew - 1
'                            gHangSu = gHangSu - 1
                        End If
                        
                        gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow, h0, h1)
                        Debug.Print gNowNum, gLastNew, gHangSu, depthAll, depthnow
                        'Debug.Assert gHangSu >= 0
                    Else
                        gHangSu = h1 + 1
                        
                        gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow, h0, h1)
                        Debug.Print gNowNum, gLastNew, gHangSu, depthAll, depthnow
                    End If
                    
                End If
            Loop
        End If
    Loop While gNowNum > maxNum
    Write_TU03
    
    If Con.State = 0 Then Con_Open
    
    sSql = "select subj,seq from tp03 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num=" & gNowNum
    Set rs = Fn_SQLExec(sSql).rs
    
    cQuiz.num = gNowNum
    
Wend

End If
Screen.MousePointer = vbDefault

'==========================================================
'다음 클릭에도 비슷하게 구현되어 있다. 여기서 한번 뜨면 거기서는 안떠야 한다.
'==========================================================
Dim Jindo As Single
Dim ans As String

'If cQuiz.overFlow And cQuiz.isBefore = False And cQuiz.bNext = True And cQuiz.bPass = False And cQuiz.isNew = False Then
If cQuiz.overFlow And cQuiz.isBefore = False And cQuiz.bNext = True And cQuiz.bPass = False And gIsNew = False Then
    Jindo = GETjindo(gPocket, gChasu, Con.State)

    If Fix(Jindo) >= 100 Then
        cQuiz.bPass = True
        TmrAfterTTS_focus.Enabled = False
        ans = MsgBox("해당 시험지 문제를 모두 푸셨습니다. 중단하시겠습니까?", vbAbortRetryIgnore + vbQuestion)
        If ans = vbYes Then '삭제
            deletePaper gPocket, gChasu

            frmMain.mnuRefresh_Click ' mnuRefresh_Click
            Unload Me
            quizDisp = False
            Exit Function
        ElseIf ans = vbRetry Or ans = vbIgnore Then
            cQuiz.overFlow = False
'            quizDisp
        ElseIf ans = vbAbort Then
            cmdClose_Click
            quizDisp = False
            
            ans = MsgBox("틀린문제를 확인하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton1)
            
            If ans = vbYes Then
                '틀린문제 시험지를 만든다
                
                Dim bSuccess As Boolean
                bSuccess = frmMain.TvNodeSelect(gPocket & " " & gChasu)
                
                If bSuccess Then
                    
                    Module1.vDataGlobal2 = ""
                    frmMain.mnuMakeSub_Click (0) '틀린문제시험지를 만든다음
                    
                    If Module1.vDataGlobal2 <> "" Then
                        bSuccess = frmMain.TvNodeSelect(Module1.vDataGlobal2)
                        
                        If bSuccess Then
                            '학습생략 후 바로 시작
                            frmMain.mnuQuickStart_Click
                        End If
                        
                    End If
                    
                End If
            End If
            Exit Function
        Else
            Debug.Assert False
        End If
    End If
End If
'==========================================================
'다음 클릭에도 비슷하게 구현되어 있다 종료
'==========================================================



'Screen.MousePointer = vbDefault
cQuiz.subj = rs(0)
cQuiz.seq = rs(1)

sSql = " select A.cat,A.quiz,A.a,A.b,A.c,A.d,A.e,A.hint,A.ans,A.resid,A.mode, "
sSql = sSql & " b.o,b.x,b.update_YMD,B.RESERVE_YMD,B.GANGYEK,A.seq from  vq01 A, TU02 B "
sSql = sSql & " where A.subj='" & cQuiz.subj & "' and A.seq= " & cQuiz.seq & " AND B.USERID='" & gUserid & "'"
sSql = sSql & " AND A.SUBJ=B.SUBJ AND A.SEQ=b.SEQ "

Call rs.Close '? 과연 ?
Set rs = Fn_SQLExec(sSql, , , True).rs

Set notExistList = New Collection

Dim bRVal2 As Boolean

On Error Resume Next
Call notExistList.Add(True, CStr(gNowNum))
While rs.EOF
    '100문제 초과했을경우...
    If notExistList.count > 100 Then
        On Error GoTo ErrTrap
        Call err.Raise(vbObjectError + 100, , "시험지정보가 올바르지 않습니다. 시험지를 삭제하세요.")
        MsgBox "문제가 없습니다. 시험지를 삭제하세요", vbExclamation
        GoTo ErrTrap
    End If
    
    bRVal2 = notExistList(CStr(gNowNum))
    If err.Number <> 0 Then
        err.Clear
        Call notExistList.Add(True, cQuiz.subj & cQuiz.seq)
        
        gHangSu = gHangSu + 1
        gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow)
        '====================================================================================
        'FFF
        'gnownum 을 구하고 나서 대입할때 c.num의 최대값 이상이 발견되는 경우에는 의미가 없으므로 db에 저장할
        '필요가 없을 것 같다. 혹은 하나씩 증가하는 값 보다는 bisection 알고리즘을 적용한 값으로 하든지..
        '전자가 더 나은 방법일것 같지만 하여간 이부분에서의 성능 개선이 요구된다.
        '200601014 by 박규선
        Debug.Assert False '위와 같은 성능개선 필요
        '그런데 다른곳에서 사용하고 나서 이부분의 진도를 보여주므로 write_TU03 이 이루어지므로
        '다른곳에서는 문제를 90%풀었는데도 불구하고 이곳에서 10%부터 계속하여 진행되는 경우가 있을 수 있다.
        '이부분의 진입은 테스트 를 하였으며 이부부에서 진입되었을때 그러한 경우 진도관리가 현명하게
        '합리적으로 이루어지도록 성능과 관련하여 해결될 필요가 있다.
        '====================================================================================
        
        Call notExistList.Add(True, CStr(gNowNum))
        Write_TU03
        gNowNum = Read_TU03
         
        sSql = " select A.cat,A.quiz,A.a,A.b,A.c,A.d,A.e,A.hint,A.ans,A.resid,A.mode, "
        sSql = sSql & vbCrLf & " b.o,b.x,b.update_YMD,B.RESERVE_YMD,B.GANGYEK,A.seq from  vq01 A, TU02 B , tp03 c "
        sSql = sSql & vbCrLf & " where A.subj=c.subj and A.seq=c.seq AND B.USERID='" & gUserid & "'"
        sSql = sSql & vbCrLf & " AND A.SUBJ=B.SUBJ AND A.SEQ=b.SEQ and c.subj='" & cQuiz.subj & "'"
        sSql = sSql & vbCrLf & " and c.userid='" & gUserid & "' and c.pocketnm='" & gPocket & "'"
        sSql = sSql & vbCrLf & " and c.chasu=" & gChasu & " and c.num=" & gNowNum & ""
        Debug.Assert Con.State = 1
        Set rs = Fn_SQLExec(sSql).rs
    Else
        gHangSu = gHangSu + 1
        gNowNum = cMNIQ.getF(gLastNew, gHangSu, gOrder, gIsNew, depthAll, depthnow)
'        Call notExistList.Add(True, CStr(gNowNum))
    End If
Wend

On Error GoTo ErrTrap

Dim lstchr As String


cQuiz.chk = 0
chk.Value = vbUnchecked
cQuiz.Correct_chk = False '채점은 아직 안함.
cQuiz.cat = rs(0)
cQuiz.quiz = rs(1)

cQuiz.o = rs("O")
cQuiz.X = rs("X")
cQuiz.update_ymd = rs("UPDATE_YMD")
cQuiz.reserve_ymd = rs("RESERVE_YMD")
cQuiz.gangyek = rs("GANGYEK")
frmMain.lblTitle(1).Caption = frmMain.lblTitle(1).Tag & "[" & iString(cQuiz.o, 1) & iString(cQuiz.X, 2) & "]"
frmMain.lblTitle(1).ForeColor = calForgetcolor2(cQuiz.o, cQuiz.X, cQuiz.update_ymd, cQuiz.reserve_ymd, cQuiz.gangyek)

vHint = rs("hint")
If IsNull(vHint) Or vHint = "" Then
    cQuiz.hint = "@"
Else
    cQuiz.hint = vHint
End If

cQuiz.ans = rs("ans")
cQuiz.resid = rs("resid")
cQuiz.mode = rs("mode")


lstchr = Right(cQuiz.quiz, 1)

Dim strLeft1 As String, lenStrLeft1 As Long
Dim seq As Long

If cQuiz.mode <> "2" Then
    'mode = ox 가 아니면
        
    cQuiz.a = rs(2)
    
    If lstchr <> "-" And lstchr <> "~" And cQuiz.mode = "1" And 0 < InStr(cQuiz.quiz, " ") Then '답에 지렁이가 있는 경우
        lstchr = Left(cQuiz.a, 1)
    End If
    
    cQuiz.A_seq = rs(16)
    
    If cQuiz.mode = "1" Then
        strLeft1 = Left(cQuiz.a, 1)
        lenStrLeft1 = LenB(strLeft1)
    End If
    
    err.Clear
    On Error GoTo 0
    On Error Resume Next
    cQuiz.B = rs(3)
'    On Error GoTo 0
    If Len(cQuiz.B) = 0 Or err.Number <> 0 Then
      
      cQuiz.B = slectnotin(cQuiz.subj, cQuiz.seq, lstchr, True, cQuiz.mode, seq, cQuiz.a)
      cQuiz.B_seq = seq
      If cQuiz.mode = "1" And LenH(strLeft1) = 2 And strLeft1 = Left(cQuiz.B, 1) And lenStrLeft1 = 2 Then
        cQuiz.B = slectnotin(cQuiz.subj, cQuiz.seq, lstchr, True, cQuiz.mode, seq, cQuiz.a) '요구 요구하다, 등 유사답안이 등장하는 것을 방지하기 위함이다.
        cQuiz.B_seq = seq
      End If
    End If
    
    err.Clear
    On Error GoTo 0
    On Error Resume Next
    cQuiz.C = rs(4).Value
    If Len(cQuiz.C) = 0 Or err.Number <> 0 Then
          cQuiz.C = slectnotin(cQuiz.subj, cQuiz.seq, lstchr, , , seq, cQuiz.a)
          cQuiz.C_seq = seq
          If cQuiz.mode = "1" And LenH(strLeft1) = 2 And strLeft1 = Left(cQuiz.C, 1) And lenStrLeft1 = 2 Then
            cQuiz.C = slectnotin(cQuiz.subj, cQuiz.seq, lstchr, True, cQuiz.mode, seq, cQuiz.a) '요구 요구하다, 등 유사답안이 등장하는 것을 방지하기 위함이다.
            cQuiz.C_seq = seq
          End If
    End If
    
    err.Clear
    On Error GoTo 0
    On Error Resume Next
    cQuiz.d = rs(5).Value
    If Len(cQuiz.d) = 0 Or err.Number <> 0 Then
      cQuiz.d = slectnotin(cQuiz.subj, cQuiz.seq, lstchr, , , seq, cQuiz.a)
      cQuiz.D_seq = seq
      If cQuiz.mode = "1" And LenH(strLeft1) = 2 And strLeft1 = Left(cQuiz.d, 1) And lenStrLeft1 = 2 Then
        cQuiz.d = slectnotin(cQuiz.subj, cQuiz.seq, lstchr, True, cQuiz.mode, seq, cQuiz.a) '요구 요구하다, 등 유사답안이 등장하는 것을 방지하기 위함이다.
        cQuiz.D_seq = seq
      End If
    End If
    
    If cQuiz.mode = "5" Then
      '5지선다
      err.Clear
      On Error GoTo 0
      On Error Resume Next
      cQuiz.e = rs(6)
      If Len(cQuiz.e) = 0 Or err.Number <> 0 Then
         cQuiz.e = slectnotin(cQuiz.subj, cQuiz.seq, lstchr, , , seq, cQuiz.a)
         cQuiz.E_seq = seq
      End If
    End If
    
    On Error GoTo 0
    On Error GoTo ErrTrap
    
   
    
    If UCase(cQuiz.ans) = "A" Or UCase(cQuiz.ans) = "B" Or UCase(cQuiz.ans) = "C" Or UCase(cQuiz.ans) = "D" Or UCase(cQuiz.ans) = "E" Then
        optA.Caption = "A)"
        optB.Caption = "B)"
        optC.Caption = "C)"
        optD.Caption = "D)"
        optE.Caption = "E)"
        
    Else
        optA.Caption = "1)"
        optB.Caption = "2)"
        optC.Caption = "3)"
        optD.Caption = "4)"
        optE.Caption = "5)"
    
    End If

Else 'mode=2 O X 퀴즈
        optA.Caption = "O)"
        optB.Caption = "X)"
        
        

End If

If cQuiz.ans = "1" Or cQuiz.ans = "O" Then
    cQuiz.ans = "A"
ElseIf cQuiz.ans = "2" Or cQuiz.ans = "X" Then
    cQuiz.ans = "B"
ElseIf cQuiz.ans = "3" Then
    cQuiz.ans = "C"
ElseIf cQuiz.ans = "4" Then
    cQuiz.ans = "D"
ElseIf cQuiz.ans = "5" Then
    cQuiz.ans = "E"
End If


'.isNew = gIsNew
Con_Close
'End If

'5지선다형 및 4지선다형은 보기를 다 보이고 단답식(1)은 4지선다형으로 한다 그리고 ox형은 2로 하고 o x 로 한다.

If cQuiz.mode = "5" Then
    optC.Enabled = True
    optC.Visible = True
    optD.Enabled = True
    optD.Visible = True


    optE.Enabled = True
    optE.Visible = True
    
    
    
    
    imgA1.Visible = True
    imgB1.Visible = True
    imgC1.Visible = True
    imgD1.Visible = True
    imgE1.Visible = True
        
ElseIf cQuiz.mode = "4" Or cQuiz.mode = "1" Then
    optC.Enabled = True
    optC.Visible = True
    optD.Enabled = True
    optD.Visible = True
    
    optE.Enabled = False
    optE.Visible = False
    
    
    
    imgA1.Visible = True
    imgB1.Visible = True
    imgC1.Visible = True
    imgD1.Visible = True
    imgE1.Visible = False
        
ElseIf cQuiz.mode = "2" Then
    optC.Enabled = False
    optC.Visible = False
    optD.Enabled = False
    optD.Visible = False
    optE.Enabled = False
    optE.Visible = False

    imgA1.Visible = True
    imgB1.Visible = True
    imgC1.Visible = False
    imgD1.Visible = False
    imgE1.Visible = False
            
End If

optTTSQuiz.Visible = True
optTTSA.Visible = optA.Visible
optTTSB.Visible = optB.Visible
optTTSC.Visible = optC.Visible
optTTSD.Visible = optD.Visible
optTTSE.Visible = optE.Visible


cQuiz.memo = getTableVal("hint", "th01", "where userid='" & gUserid & "' and subj = '" & cQuiz.subj & "' and seq = " & cQuiz.seq & " ")


'공유된 메모에서 하나를 찾아온다.

'=================================================================
'정책에서 제외시켰다.
'개인적인 기록으로 간주하도록 하며 공유할 경우
'부작용이 우려되므로 개인적인 메모외에는 공유하지 않는다.
'=================================================================
'If Len(.memo) = 0 Then
'    .memo = getTableVal("hint", "th01", "where bshare='1' and subj = '" & .subj & "' and seq = " & .seq & " ")
'End If

If Not (lfrmMemo Is Nothing) Then
    Call lfrmMemo.setVal(gUserid, cQuiz.subj, cQuiz.seq, cQuiz.memo, cFTP, cQuiz)
End If


'메모에 내용이 있으면
If Len(cQuiz.memo) > 0 Then
    'cmdMemo.Picture = LoadResPicture(residxxx, 0)
    cmdMemo.BackColor = rgb(Round15(rndVal(0, 239)), 210, 100)  ' vbBlue
    
'    If oTooltip.ToolCount = 0 Then
'        Call oTooltip.AddTool(cmdMemo, Replace(.memo, vbTab, Space(4)))
'
'    Else
    strToolTip = Replace(cQuiz.memo, vbTab, Space(4))
        oTooltip.ToolText(cmdMemo) = strToolTip
'    End If
    
    Call oTooltip.setTitle("", 0)
Else
'메모에 내용이 없으면

'------------------------------------------------------------------------------
'명언을 보여준다.
'------------------------------------------------------------------------------
   cmdMemo.BackColor = vbButtonFace
   oTooltip.ToolText(cmdMemo) = " "
   strToolTip = "" 'selRndTip()
   Call oTooltip.setTitle(strToolTip, 1)
'------------------------------------------------------------------------------
'명언을 보여준다. 끝
'------------------------------------------------------------------------------
End If

If isExistsEE() Then
    cmdEE.BackColor = rgb(Round15(rndVal(0, 239)), 210, 100)  ' vbBlue
Else
    cmdEE.BackColor = vbButtonFace
End If

If Not (lfrmMemo Is Nothing) Then
    If lfrmMemo.Visible = False Then
        lfrmMemo.rtxt1.Text = cQuiz.memo
    End If
    Call lfrmMemo.save
    Call lfrmMemo.setVal(gUserid, cQuiz.subj, cQuiz.seq, cQuiz.memo, cFTP, cQuiz)
End If

Dim K As Integer
Dim N As Integer
Dim str4 As String


    If cQuiz.forReview = False Then
        Read_TU03
        cQuiz.isNew = gIsNew
    End If

End If


Randomize

Dim chkSwapBySubject As Boolean

    cQuiz.A_ans = "A"
    cQuiz.B_ans = "B"
    cQuiz.C_ans = "C"
    cQuiz.D_ans = "D"
    cQuiz.E_ans = "E"

'ForHint = false <== swap oK but forhint=true==>틀린문제를 디스플레이하는경우 섞지 않는다.
If (cQuiz.mode = "1" And cQuiz.forReview = False And cQuiz.isNew = False) Or (gMainOnResize = False And ForHint = False And Not (cQuiz.Correct_chk And cQuiz.Correct = False) And (cQuiz.isNew = False And cQuiz.NoExistABCDinHint)) Then  '힌트에 ABCD가 없으면 문제를 섞는다.(20040910) 단 새로운 문제가 아닌경우이다.
    chkSwapBySubject = True
    If cQuiz.subj = "토익990" Then
        If InStr(cQuiz.quiz, "(A)") > 0 Then
            '토익990 의 경우 예외처리한다.
            '문제에 (A) 가 들어가면 섞지 않도록 한다. 20071207 PGS
            chkSwapBySubject = False
        End If
    End If
    
    If chkSwapBySubject Then
        If cQuiz.mode = "1" Or cQuiz.mode = "4" Then
            If Not ForHint Then
                
                
                K = Fix(Rnd * 1000) Mod 4
                
                For N = 0 To K
                    Call cQuiz.swap
                Next
            End If
        ElseIf cQuiz.mode = "5" Then
            If Not ForHint Then
                
                
                K = Fix(Rnd * 1000) Mod 5
                
                For N = 0 To K
                    Call cQuiz.swap
                Next
            End If
        End If
    End If
End If

pic1.CurrentX = 0
pic1.CurrentY = 0

pic1.cls
picMain.cls

picMain.CurrentX = 0
picMain.CurrentY = pic1.Top - pic1.TextHeight("A")


cmdDel.Enabled = cQuiz.isNew
If cQuiz.isNew Then
    'frmMain.characterPlay ""
    picMain.ForeColor = vbBlue
    picMain.Print vbCrLf & "    ●"
    chk.Enabled = False
ElseIf cQuiz.forReview Then
    'frmMain.characterPlay ""
    picMain.ForeColor = vbMagenta
    picMain.Print vbCrLf & "    ●"
    chk.Enabled = False
Else
    If frmQuiz.cQuiz.reserve_ymd <> "99999999" And frmQuiz.cQuiz.reserve_ymd >= date2Str(DateAdd("d", Module1.ALLOW_AFFECT_DATE, Now)) And frmQuiz.cQuiz.update_ymd >= date2Str(DateAdd("d", Module1.ALLOW_AFFECT_DATE30, Now)) Then
        picMain.ForeColor = vbBlack '유효하지 않은 학습 문항이다.
        picMain.Print vbCrLf & "    ●"
        chk.Enabled = True
    Else
        chk.Enabled = True
    End If
End If

picMain.ForeColor = vbBlack
pic1.ForeColor = vbBlack
str4 = "[" & cQuiz.cat & "]"

picMain.CurrentX = pic1.Left / 24
picMain.CurrentY = pic1.Top - 10 - picMain.TextHeight("A")

picMain.Print str4


'picmain.CurrentX = picmain.TextWidth(.cat & "]] ")
'picmain.CurrentY = picmain.CurrentY + picmain.TextHeight("A") / 2
pic1.FontBold = True

Dim str_url As String

'자동줄바꿈은 :가 있는 경우만 이다.
'If True Or InStr(.cat, ":") > 0 Then '없는 경우도 이제 줄바꿈 된다.
    If 5 < Len(cQuiz.quiz) Then
        If Left(cQuiz.quiz, 5) = "@http" Then
            str_url = Mid(cQuiz.quiz, 2) '@제거
            cQuiz.quiz = ""
            
            Call wbQuizMain.Navigate(str_url)
            
        End If
    End If
    pic1.Print autoCRLF(cQuiz.quiz, picMain.Width - pic1.Left * 2, pic1, True)
    'Debug.Assert InStr(cQuiz.hint, vbCrLf & vbCrLf) = 0 '어쩌다가 vbcrlf 가 쫙쫙 붙는 경우가 있다.
    If Not cQuiz.forReview Then
        cQuiz.hint2 = autoCRLF(cQuiz.hint, picMain.Width - pic1.Left * 2, pic1, True)
    End If
'Else
'    pic1.Print .quiz
'End If

'picmain.CurrentX = 500
pic1.CurrentY = pic1.CurrentY + pic1.TextHeight("A")

imgA0.Top = pic1.CurrentY
imgA0.Left = 0
imgA0.Width = pic1.Width

optA.Left = 200
optA.Top = imgA0.Top + pic1.Top '+ (pic1.TextHeight("A") - optA.Height) / 2

optTTSA.Left = optA.Width + 200
optTTSA.Top = optA.Top

optTTSQuiz.Left = 200
optTTSQuiz.Top = optA.Top - optTTSQuiz.Height * 1.5

If (gIsNew Or cQuiz.forReview) And (cQuiz.ans = "A" Or cQuiz.ans = "1" Or cQuiz.ans = "O") Then
    pic1.ForeColor = vbBlue
    pic1.FontBold = True
    optA.Value = True
Else
    pic1.ForeColor = vbBlack
    pic1.FontBold = cFTP.Profile.FontBold
End If



'자동줄바꿈은 categori에 ":" 문자가 있는경우이다.
'If InStr(.cat, ":") > 0 Then
    Debug.Assert InStr(cQuiz.a, vbCrLf & vbCrLf) = 0
    cQuiz.a2 = autoCRLF(cQuiz.a, picMain.Width - pic1.Left * 2, pic1, True) '값을 어싸인 하지 않고 바로 출력한다.
pic1.Print cQuiz.a2 'autoCRLF(cQuiz.a, picMain.Width - pic1.Left * 2, pic1, True)
'End If

'pic1.Print .a


pic1.CurrentY = pic1.CurrentY + pic1.TextHeight("A") / 2
imgA0.Height = pic1.CurrentY - imgA0.Top

imgB0.Top = pic1.CurrentY
imgB0.Left = 0
imgB0.Width = pic1.Width

optB.Left = 200
optB.Top = imgB0.Top + pic1.Top '+ (imgA0.Height - optA.Height) / 2

optTTSB.Left = optA.Width + 200
optTTSB.Top = optB.Top

If (gIsNew Or cQuiz.forReview) And (cQuiz.ans = "B" Or cQuiz.ans = "2" Or cQuiz.ans = "X") Then
    pic1.ForeColor = vbBlue
    pic1.FontBold = True
    optB.Value = True
Else
    pic1.ForeColor = vbBlack
    pic1.FontBold = cFTP.Profile.FontBold
End If

'자동줄바꿈은 categori에 ":" 문자가 있는경우이다.
'If InStr(.cat, ":") > 0 Then
    cQuiz.b2 = autoCRLF(cQuiz.B, picMain.Width - pic1.Left * 2, pic1, True)
pic1.Print cQuiz.b2 'autoCRLF(cQuiz.B, picMain.Width - pic1.Left * 2, pic1, True)
'End If
    
'pic1.Print .B

pic1.CurrentY = pic1.CurrentY + pic1.TextHeight("A") / 2
imgB0.Height = pic1.CurrentY - imgB0.Top

imgC0.Top = pic1.CurrentY
imgC0.Left = 0
imgC0.Width = pic1.Width

optC.Left = 200
optC.Top = imgC0.Top + pic1.Top '+ (imgA(0).Height - optA.Height) / 2

optTTSC.Left = optA.Width + 200
optTTSC.Top = optC.Top

If (gIsNew Or cQuiz.forReview) And (cQuiz.ans = "C" Or cQuiz.ans = "3") Then
    pic1.ForeColor = vbBlue
    pic1.FontBold = True
    optC.Value = True
Else
    pic1.ForeColor = vbBlack
    pic1.FontBold = cFTP.Profile.FontBold
End If


'자동줄바꿈은 categori에 ":" 문자가 있는경우이다.
'If InStr(.cat, ":") > 0 Then
cQuiz.C2 = autoCRLF(cQuiz.C, picMain.Width - pic1.Left * 2, pic1, True)
'End If
    

pic1.Print cQuiz.C2

pic1.CurrentY = pic1.CurrentY + pic1.TextHeight("A") / 2
imgC0.Height = pic1.CurrentY - imgC0.Top


imgD0.Top = pic1.CurrentY
imgD0.Left = 0
imgD0.Width = pic1.Width

optD.Left = 200
optD.Top = imgD0.Top + pic1.Top '+ (imgA(0).Height - optA.Height) / 2

optTTSD.Left = optA.Width + 200
optTTSD.Top = optD.Top

If (gIsNew Or cQuiz.forReview) And (cQuiz.ans = "D" Or cQuiz.ans = "4") Then
    pic1.ForeColor = vbBlue
    pic1.FontBold = True
    optD.Value = True
Else
    pic1.ForeColor = vbBlack
    pic1.FontBold = cFTP.Profile.FontBold
End If

'자동줄바꿈은 categori에 ":" 문자가 있는경우이다.
'If InStr(.cat, ":") > 0 Then
cQuiz.d2 = autoCRLF(cQuiz.d, picMain.Width - pic1.Left * 2, pic1, True)
'End If

pic1.Print cQuiz.d2

If cQuiz.mode = "5" Then
    pic1.CurrentY = pic1.CurrentY + pic1.TextHeight("A") / 2
    imgD0.Height = pic1.CurrentY - imgD0.Top
    
    imgE0.Top = pic1.CurrentY
    imgE0.Left = 0
    imgE0.Width = pic1.Width
    
    optE.Left = 200
    optE.Top = imgE0.Top + pic1.Top '+ (imgA(0).Height - optA.Height) / 2
    
    optTTSE.Left = optA.Width + 200
    optTTSE.Top = optE.Top

    If (gIsNew Or cQuiz.forReview) And cQuiz.ans = "E" Then
        pic1.ForeColor = vbBlue
        pic1.FontBold = True
        optE.Value = True
    Else
        pic1.ForeColor = vbBlack
        pic1.FontBold = cFTP.Profile.FontBold
    End If
    
    '자동줄바꿈은 categori에 ":" 문자가 있는경우이다.
    'If True Or InStr(.cat, ":") > 0 Then
    cQuiz.e2 = autoCRLF(cQuiz.e, picMain.Width - pic1.Left * 2, pic1, True)
    'End If
    
    pic1.Print cQuiz.e2
    
End If


'cmdHint.Left = pic1.CurrentX + pic1.Left
'cmdHint.Top = pic1.CurrentY + pic1.Top - 90 + pic1.Top

If cQuiz.mode = "4" Or cQuiz.mode = "1" Then
    imgD0.Height = pic1.TextHeight(cQuiz.d2) ' cmdHint.Top - imgC0.Top
ElseIf cQuiz.mode = "5" Then
    imgE0.Height = pic1.TextHeight(cQuiz.e2) '  cmdHint.Top - imgD0.Top
End If


'End With



If cQuiz.isNew Or cQuiz.forReview Then
    cmdHint.Visible = True
    
End If

If cQuiz.hint = "@" Then
    cmdHint.Visible = False
Else
    cmdHint.Visible = True
End If




If ForHint = False Or ForDel = True Then
    If cQuiz.isNew Then
        pgBar.Value = pgBar.Min
        pgBar.Max = cFTP.Profile.TimeOutSecStudy
        pgBar.Value = pgBar.Max
        
        If Not cFTP.Profile.bSetTimeOutStudy Then
            Timer1.Enabled = False
        Else
'            If cFTP.Profile.CntOfStreatOutNow < cFTP.Profile.CntOfStreatOutSetting Then
                Timer1.Enabled = True
'            Else
'                Timer1.Enabled = False
'            End If
        End If
    
    ElseIf cQuiz.forReview Then
'        cmdHint.Visible = True
        
        pgBar.Value = pgBar.Min
        pgBar.Max = cFTP.Profile.TimeOutSecStudy
        pgBar.Value = pgBar.Max
        
        If Not cFTP.Profile.bSetTimeOutStudy Then
            Timer1.Enabled = False
        Else
            If cFTP.Profile.CntOfStreatOutNow < cFTP.Profile.CntOfStreatOutSetting Then
                Timer1.Enabled = True
            Else
                Timer1.Enabled = False
            End If
        End If
    Else '폼 리사이즈
        
        pgBar.Value = pgBar.Min
        
        pgBar.Max = cFTP.Profile.TimeOutSec
        If Not gQuizOnResize Then
            pgBar.Value = pgBar.Max '이거빼면 뒤로가길할때 오류남.
        Else
            If pgBar.Max < gPgBarSaveValue Then
                pgBar.Value = pgBar.Max
            Else
                pgBar.Value = gPgBarSaveValue
            End If
            
        End If
        
        If cFTP.Profile.bSetTimeOut Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End If
End If

If picMain.Visible Then
    picMain.SetFocus
End If

imgA1.Left = 0
imgA1.Top = imgA0.Top + pic1.Top
imgA1.Height = imgA0.Height

imgB1.Left = 0
imgB1.Top = imgB0.Top + pic1.Top
imgB1.Height = imgB0.Height

imgC1.Left = 0
imgC1.Top = imgC0.Top + pic1.Top
imgC1.Height = imgC0.Height

imgD1.Left = 0
imgD1.Top = imgD0.Top + pic1.Top
imgD1.Height = imgD0.Height

imgE1.Left = 0
imgE1.Top = imgE0.Top + pic1.Top
imgE1.Height = imgE0.Height


cmdHint.Left = pic1.Left - cmdHint.Width

If imgE1.Visible Then
    cmdHint.Top = optE.Height + optE.Top + imgE0.Height
ElseIf imgD1.Visible Then
    cmdHint.Top = optD.Height + optD.Top + imgD0.Height
ElseIf imgC1.Visible Then
    cmdHint.Top = optC.Height + optC.Top + imgC0.Height
ElseIf imgB1.Visible Then
    cmdHint.Top = optB.Height + optB.Top + imgB0.Height
End If

If cmdHint.Top > pic1.Height Then
    cmdHint.Top = pic1.Height - cmdHint.Height
End If

If cQuiz.resid <> "0" Then
    cmdRef.Left = cmdHint.Left + cmdHint.Width + cmdHint.Height
    cmdRef.Top = cmdHint.Top '+ pic1.Top
    cmdRef.Visible = True
    cmdRef_Click
Else
    cmdRef.Visible = False
    frmRef.Visible = False
End If

If frmHint.Visible And cmdHint.Visible And cQuiz.isNew = True Then
    cmdHint_Click
Else
    frmHint.Visible = False
End If

picMain.PaintPicture pic1.Image, pic1.Left, pic1.Top
pic1.Visible = False

cQuiz.dLastSec = Timer()

quizDisp = True

'tts 설정

If InStr(cQuiz.subj, "한영") > 0 And (cQuiz.isNew Or cQuiz.Correct_chk) Then
    '한영이라고 하더라도 발음은 영어로 발음하게한다.
    If cQuiz.ans = "A" Then
        clipText = Trim(cQuiz.a)
    ElseIf cQuiz.ans = "B" Then
        clipText = Trim(cQuiz.B)
    ElseIf cQuiz.ans = "C" Then
        clipText = Trim(cQuiz.C)
    ElseIf cQuiz.ans = "D" Then
        clipText = Trim(cQuiz.d)
    ElseIf cQuiz.ans = "E" Then
        clipText = Trim(cQuiz.e)
    Else
        clipText = cQuiz.quiz
    End If
Else
    clipText = cQuiz.quiz
End If

If mnu_auto_tts.Checked And gQuizOnResize = False Then '리사이즈시에는 재생안함.
    If Not ForHint Then '힌트<하얀점을위한 resize힌트>를 클릭해서 들어온 경우는 재생 안함 ( 이미 재생한 경우이므로 )
        '복습문제 재생 안함 설정인 경우 재생 안함
        If cFTP.Profile.bChkTTSuse1 And cQuiz.isNew Then
            Call mnuTTS0_Click
        ElseIf Not cQuiz.isNew Then 'isNew가 아닌것과 forReview가 참인것이 같지 않는 이유는 forReview는 자주색원이 표시되는 경우이다.
            If cFTP.Profile.bChkTTSuse4 Then
                Call mnuTTS0_Click
            ElseIf cFTP.Profile.bChkTTSuse2 And cQuiz.forReview Then            'bPass가 합격 불합격을 의미하는 것이 아님.
                Call mnuTTS0_Click
            ElseIf cFTP.Profile.bChkTTSuse3 And (cQuiz.reserve_ymd < Format(Now(), "yyyymmdd") Or cQuiz.reserve_ymd = "99999999") Then
                Call mnuTTS0_Click
            ElseIf cFTP.Profile.bChkTTSuse1 And cFTP.Profile.bChkTTSuse2 And cFTP.Profile.bChkTTSuse3 And cFTP.Profile.bChkTTSuse4 Then
                Call mnuTTS0_Click
            End If
        End If
        
    End If
End If

If cQuiz.isNew Or cQuiz.forReview Then
    frmMain.characterPlay "" 'TTS 실행 후 움직이게 한다.
End If

'새로운 문제인 경우만 사전 팝업설정
If 0 < Len(clipText) Then
    If cQuiz.isNew Then
        clipText = forTTS(clipText) '(A) _ 공백등을 지운다.
        If mnu_auto_dic.Checked Then
            Call mnuEndic_Click
        ElseIf Len(clipText) = LenH(clipText) Then '영문인경우
            Call mnuEndic_Click
        End If
    ElseIf cQuiz.forReview And cQuiz.isBefore = False Then
        clipText = forTTS(clipText) '(A) _ 공백등을 지운다.
        If Len(clipText) = LenH(clipText) Then '영문인경우
            Call mnuEndic_Click
        End If
    End If
End If

If cQuiz.mode = "1" Then '4지선다
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
ElseIf cQuiz.mode = "2" Then
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = False
ElseIf cQuiz.mode = "4" Then
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False
ElseIf cQuiz.mode = "5" Then
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = True

End If


If gMainOnResize = True Then
    If optA.Value = True Then
        imgA1_Click
    ElseIf optB.Value = True Then
        imgB1_Click
    ElseIf optC.Value = True Then
        imgC1_Click
    ElseIf optD.Value = True Then
        imgD1_Click
    ElseIf optE.Value = True Then
        imgE1_Click
    End If
End If

'pgBarJindo.Max = frmQuiz.TotalQuizCount
If gHangSu <= pgBarJindo.Max Then
    pgBarJindo.Value = gHangSu
Else
    pgBarJindo.Value = pgBarJindo.Max
End If

pgBarJindo.ToolTipText = "" & gHangSu & " / " & pgBarJindo.Max

txtJindo.Text = "" & gHangSu & " / " & frmQuiz.TotalQuizCount
'//Call MsgBox("" & gHangSu & "/" & frmQuiz.TotalQuizCount, vbOKOnly, "")

Exit Function
ErrTrap:
Screen.MousePointer = vbDefault
'
MsgBox err.Description, vbCritical
If IDEMODE Then
    If MsgBox("재시도?IDEMODE", vbExclamation + vbYesNo) = vbYes Then
        Resume
    End If
End If
Me.Hide
'cmdClose_Click
Unload Me
quizDisp = False
'
''
End Function

Function forTTS(str As String) As String
forTTS = str
forTTS = Replace(forTTS, "(A)", "")
forTTS = Replace(forTTS, "(B)", "")
forTTS = Replace(forTTS, "(C)", "")
forTTS = Replace(forTTS, "(D)", "")
forTTS = Replace(forTTS, vbNewLine, " ")
forTTS = Replace(forTTS, "_", "")
forTTS = Replace(forTTS, "    ", " ")
forTTS = Replace(forTTS, "   ", " ")
forTTS = Replace(forTTS, "  ", " ")
forTTS = Replace(forTTS, "  ", " ")
forTTS = Replace(forTTS, "  ", " ")

End Function

Function Incress(rst As Integer) As Boolean '성공 True
Dim ymd As String
Dim sql1 As String
Dim sql_tu02_bonus As String
Dim sql_rt01_bonus As String
Dim affected As Integer
Dim cnt As Long
Static oldcnt1 As Long
Static oldcnt2 As Long
Static oldcnt5 As Long

Dim preGanGyek  As Long

Incress = False
'On Error Resume Next
ymd = Format(Now, "YYYYMMDD")

If rst = 1 Then
    cFTP.Profile.CntOfStreatOutNow = 0
Else
    If pgBar.Value = 0 Then
        cFTP.Profile.CntOfStreatOutNow = cFTP.Profile.CntOfStreatOutNow + 1
    Else
        cFTP.Profile.CntOfStreatOutNow = 0
    End If
End If

Read_TU02

If rst = 1 Then
    '-----------------------
    '맞은경우
    '-----------------------
    '1 예정일자가 오늘보다 큰경우 eg> 오늘 20040717 예정일:20041125
    '  1.1 예정일=20041125+(update_ymd-ymd)
    '  1.2 인터벌: 인터벌+(update_ymd-ymd)
    '-----------------------
    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
    '  2.0 만약 예정일과 오늘과의 차이가 인터벌보다 크면 그값의 배를 인터벌로 함
    '  2.1 예정일=ymd+(인터벌) (X)
    '  2.1 예정일=ymd  (O) 2004.08.16   2004.08.26 YMD+인터벌*1이하의 값
    '  2.1 예정일=ymd+(인터벌) 예정일이 오늘보다 하루작은경우가 정상적인 경우이므로 기존인터벌만큼 증가시킴 2004.08.27
    '  2.2 인터벌: (인터벌+1)*2
    '  2.3 예정일이 오늘날짜에 비해 10년 차이가 나는 경우는 인터벌을 무조건 2로 셋팅하고 5년 차이가 나면 인터벌을 예전의 인터벌*0.5 로 하고 2.5년 차이가 나면 인터벌을 기존인터벌*0.25 로 조정한다. 또한 거기에 정답율을 적용한다.
    '-----------------------
    '3. 예정일자가 오늘날짜인경우는 인터벌과 예정일을 변화시키지 않음.
    '-----------------------
    
    '1 예정일자가 오늘보다 크거나 같고 수정을 오늘 안한경우 eg> 오늘 20040717 예정일:20041125
'    Debug.Assert cQuiz.reserve_ymd >= "20050101"
    If cQuiz.reserve_ymd >= ymd And cQuiz.update_ymd <> ymd Then
    
    '  1.1 예정일=20041125+(update_ymd-ymd)
        cQuiz.reserve_ymd = date2Str(DateAdd("d", Abs(DateDiff("d", str2Date(cQuiz.update_ymd), str2Date(ymd))), str2Date(cQuiz.reserve_ymd)))
        
    '  1.2 인터벌: 인터벌+(update_ymd-ymd)
        cQuiz.gangyek = cQuiz.gangyek + Abs(DateDiff("d", str2Date(cQuiz.update_ymd), str2Date(ymd)))
    
        '보너스인경우 간격을 2배 높임
        If (cQuiz.dLastSec = 0) Then
            cQuiz.gangyek = (cQuiz.gangyek + 1) * 2
        End If
    
    '-----------------------
    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
    ElseIf cQuiz.reserve_ymd < ymd And cQuiz.update_ymd <> ymd Then
       preGanGyek = Abs(DateDiff("d", str2Date(cQuiz.reserve_ymd), str2Date(ymd)))
'       If preGanGyek > cQuiz.gangyek Then
'           cQuiz.gangyek = preGanGyek * 2
'       End If
       
       '보너스인경우 2초내에 문제를 맞춘경우

        If (cQuiz.dLastSec = 0) Then
        
        '  2.1 예정일=ymd'+(인터벌)
            If cQuiz.reserve_ymd = date2Str(str2Date(ymd) - 1) Then
                cQuiz.reserve_ymd = date2Str(DateAdd("d", cQuiz.gangyek, str2Date(ymd)))
            Else
                cQuiz.reserve_ymd = date2Str(DateAdd("d", CLng(cQuiz.gangyek * 1#) + 2, str2Date(ymd)))  '증가시키기 전 간격의 값으로 설정
            End If
        '  2.2 인터벌: (인터벌+1)*2
            cQuiz.gangyek = (Abs(DateDiff("d", str2Date(cQuiz.reserve_ymd), str2Date(ymd))) + 1) * 4 '간격은 나중에 문제를 풀데 맞힌경우 적용된다.
        
        
        Else
        '  2.1 예정일=ymd'+(인터벌)
            If cQuiz.reserve_ymd = date2Str(str2Date(ymd) - 1) Then
                cQuiz.reserve_ymd = date2Str(DateAdd("d", cQuiz.gangyek, str2Date(ymd)))
            Else
                'cQuiz.reserve_ymd = date2Str(DateAdd("d", CLng(cQuiz.gangyek * 0.5) + 1, str2Date(ymd))) '증가시키기 전 간격의 값으로 설정
                '위에 0.5대신중감쇠(重減衰 heavy damping) 적용 기계진동학(출판사 문운당:文運堂) 73 page , heavydamping의 함수 설명 참고
                cQuiz.reserve_ymd = date2Str(DateAdd("d", CLng(cQuiz.gangyek * heavydamping(-0.005, DateDiff("d", str2Date(ymd), str2Date(cQuiz.reserve_ymd)))) + 1, str2Date(ymd)))
                '증가시키기 전 간격의 값으로 설정
            End If
        '  2.2 인터벌: (인터벌+1)*2
            cQuiz.gangyek = (Abs(DateDiff("d", str2Date(cQuiz.reserve_ymd), str2Date(ymd))) + 1) * 2
        End If
        
    '-----------------------
    Else
        '예정일이 오늘인경우 업데이트를 안한다는것은 문제가 된다. 업데이트일이 오늘인경우 안한다면 모를까.
        
    End If
    
'    sSql = "update tu02 set gangyek=" & cQuiz.gangyek & ",reserve_ymd='" & cQuiz.reserve_ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
Else
    '-----------------------
    '틀린경우(예정일을 과거일로 돌리지 않도록 유의한다. 왜냐하면 여러번 틀리고 맞은것을 반복하는경우 예정일이 너무 미래로 가는경우가 있다.(위에 맞은경우 루틴참고)
    '-----------------------
    '1 예정일자가 오늘보다 같거나 큰경우 eg> 오늘 20040717 예정일:20041125
    '  1.1 예정일=ymd
    '  1.2 인터벌: 인터벌 그대로
    '-----------------------
    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
    '  1.1 예정일=ymd
    '  1.2 인터벌: 인터벌-1(단 인터벌>=0)
    '-----------------------
    '틀린경우
    '-----------------------
    '1 예정일자가 오늘보다 같거나 큰경우 eg> 오늘 20040717 예정일:20041125
    If cQuiz.reserve_ymd >= ymd Then
    '  1.1 예정일=오늘 날짜
        cQuiz.reserve_ymd = ymd 'date2Str(DateAdd("d", 0, str2Date(ymd)))
    '  1.2 인터벌: 인터벌
        'cQuiz.gangyek = cQuiz.gangyek
    '-----------------------
    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
    ElseIf cQuiz.reserve_ymd < ymd Then
        
    '  1.2 인터벌: 인터벌-1(단, 오늘여러번 틀린경우는 제외한다.)
        If (cQuiz.update_ymd <> ymd) Then
            cQuiz.gangyek = cQuiz.gangyek - 1 '빼는 경우를 좀더 다양화 해야 할 것 같음
            If cQuiz.gangyek < 0 Then
                cQuiz.gangyek = 0
            End If
        End If
    '  1.1 예정일=오늘 날짜
        cQuiz.reserve_ymd = ymd 'date2Str(DateAdd("d", 0, str2Date(ymd)))
            
    '-----------------------
    End If
    
    Dim user_ans_seq As String
    Dim user_ans_wrong As String
    
    If cQuiz.mode = "1" Then '모드가 1인 경우만 한다.
        Dim ymd2 As String
        ymd2 = date2Str(DateAdd("d", 1, str2Date(cQuiz.reserve_ymd)))
        
        If frmQuiz.cQuiz.user_ans = 1 Then
            user_ans_seq = frmQuiz.cQuiz.A_seq
        ElseIf frmQuiz.cQuiz.user_ans = 2 Then
                user_ans_seq = frmQuiz.cQuiz.B_seq
        ElseIf frmQuiz.cQuiz.user_ans = 3 Then
            user_ans_seq = frmQuiz.cQuiz.C_seq
        ElseIf frmQuiz.cQuiz.user_ans = 4 Then
            user_ans_seq = frmQuiz.cQuiz.D_seq
        ElseIf frmQuiz.cQuiz.user_ans = 5 Then
            user_ans_seq = frmQuiz.cQuiz.E_seq
        End If
        
    End If
    
    If user_ans_seq <> "" Then '틀린문제는 오늘, 오답문제는 내일로 예약일자를 셋팅한다. 그래야 복습일이 가까워지기 때문이다. 하지만 복습예정일 가까워 졌지 다른 팩터 복습일 간격이 바뀌는 것은 아니다. 단, 복습예정일이 바꾸려고 하는 날보다 커야 한다. 그렇기 때문에 오늘 틀린 문제중 하나를 오답으로 고르면 업데이트 안될 수 있는 건 당연하다.
        'sql_tp03_bonus = "update tp03 set update_ymd='20151126+1' where userid='박규선' and SUBJ='특목고입시용' and seq=713"
        sql_tu02_bonus = "update tu02 set reserve_ymd='" & ymd2 & "' where userid='" & gUserid & "' and SUBJ='" & frmQuiz.cQuiz.subj & "' and seq=" & user_ans_seq & " and reserve_ymd<>'99999999' and reserve_ymd>'" & ymd2 & "'"
        sql_rt01_bonus = "insert into rt01(subj,userid,quiz_seq,fail_seq,i_ymd,reserve_ymd) value('" & frmQuiz.cQuiz.subj & "','" & gUserid & "'," & frmQuiz.cQuiz.seq & "," & user_ans_seq & ",date_format(current_date,'%Y%m%d'),date_format(date_add(current_date,interval 7 day),'%Y%m%d'))"
    End If
    
    If frmQuiz.cQuiz.mode = 4 Or frmQuiz.cQuiz.mode = 5 Or frmQuiz.cQuiz.mode = 2 Then
    
        If frmQuiz.cQuiz.user_ans = 1 Then
            user_ans_wrong = frmQuiz.cQuiz.A_ans
        ElseIf frmQuiz.cQuiz.user_ans = 2 Then
            user_ans_wrong = frmQuiz.cQuiz.B_ans
        ElseIf frmQuiz.cQuiz.user_ans = 3 Then
            user_ans_wrong = frmQuiz.cQuiz.C_ans
        ElseIf frmQuiz.cQuiz.user_ans = 4 Then
            user_ans_wrong = frmQuiz.cQuiz.D_ans
        ElseIf frmQuiz.cQuiz.user_ans = 5 Then
            user_ans_wrong = frmQuiz.cQuiz.E_ans
        End If
        
        If user_ans_wrong <> "" Then
            sql_rt01_bonus = "insert into rt01(subj,userid,quiz_seq,user_ans,i_ymd,reserve_ymd) value('" & frmQuiz.cQuiz.subj & "','" & gUserid & "'," & frmQuiz.cQuiz.seq & ",'" & user_ans_wrong & "' ,date_format(current_date,'%Y%m%d'),date_format(date_add(current_date,interval 7 day),'%Y%m%d'))"
        End If
    End If
    
'    sSql = "update tu02 set gangyek=" & cQuiz.gangyek & ",reserve_ymd='" & cQuiz.reserve_ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
End If

If rst = 1 Then
    sSql = "update tu02 set o=o+1,update_ymd='" & ymd & "',reserve_ymd='" & cQuiz.reserve_ymd & "' , gangyek=" & cQuiz.gangyek & "  where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
    sql1 = "update tp03 set o=o+1 ,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
Else
    sSql = "update tu02 set x=x+1,update_ymd='" & ymd & "',reserve_ymd='" & cQuiz.reserve_ymd & "', gangyek=" & cQuiz.gangyek & "  where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
    sql1 = "update tp03 set x=x+1 ,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
End If
Con_Open

'업데이트 일자가 오늘인경우는 tu02를 반영하지 않는다. 20050114
'하루에 여러번 맞고 틀린것의 의미를 두지 않는다. 하루처음 풀었던 문제가 틀렸느냐 맞았느냐에
'중요도를 두어서 정말 잊혀지는 문제있는 안잊고 항상 (매일) 아는 문제인지를 구분할 수 있다.
'즉 하루에 한문제만을 100번 풀었다고 해서 다 맞았다고 해서 혹은 다 틀렸다고 했을때의 의미는 한개 틀렸다고 본다는 식이다.
If (cQuiz.update_ymd <> ymd) Then
    affected = Fn_SQLExec(sSql, , True).nrow
    
    Debug.Assert affected = 1
    If affected = 0 Then
        MsgBox "메뉴에서 시험지>어카운트갱신 을 실행 시킨 후 다시 문제를 푸세요-F7 버튼-", vbExclamation
        
        Con_Close
        cmdClose_Click
        Exit Function
    End If
ElseIf rst <> 1 Then '틀린경우이며 날짜가 같은 경우는 복습일을 모래로 한다. 내일이 아닌 모래이다.
   sSql = "update tu02 set reserve_ymd='" & date2Str(DateAdd("d", 2, str2Date(ymd))) & "', gangyek=" & cQuiz.gangyek & "  where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
   Call Fn_SQLExec(sSql, , True)
End If
affected = Fn_SQLExec(sql1).nrow

'다른곳에서 시험지 지운경우 틀린수를 하나 증가시키려하지만
'반영된것은 0이다. tp03 테이블
Debug.Assert affected = 1

If cQuiz.chk > 0 Then
    sSql = "update tu02 set chk=chk+1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
    sql1 = "update tp03 set chk=chk+1 ,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "'  and chasu=" & gChasu & "  and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
    
    affected = Fn_SQLExec(sSql).nrow
    Debug.Assert affected = 1
    
    affected = Fn_SQLExec(sql1).nrow
    Debug.Assert affected = 1

    cQuiz.chk = 0
    chk.Value = vbUnchecked
End If

If TP.chkm > 0 And rst = 1 Then
    sSql = "update tu02 set chk=chk-1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
    'sql1 = "update tp03 set chk=chk-1 ,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "'  and chasu=" & gChasu & "  and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
    affected = Fn_SQLExec(sSql, , True).nrow
    
    If err.Number <> 0 And err.Number <> -2147467259 Then
        Debug.Assert False
    Else
        Debug.Assert affected = 1
    End If
    
    'affected = Fn_SQLExec(sql1, , True).nrow

End If

If TP.xm > 0 And rst = 1 Then
    
    sSql = "update tu02 set x=x-1,o=o-1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
    'sql1 = "update tp03 set x=x-1,o=o-1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "'  and chasu=" & gChasu & "  and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
    affected = Fn_SQLExec(sSql, , True).nrow
    'affected = Fn_SQLExec(sql1, , True).nrow
    If err.Number <> 0 And err.Number <> -2147467259 Then
        Debug.Assert False
'        MsgBox "이미 Pass한 시험지를 여러 번 푼것같습니다.", vbExclamation
    Else
        Debug.Assert affected = 1

    End If
    Debug.Assert affected = 1

End If

If Len(sql_tu02_bonus) <> 0 Then
    Fn_SQLExec (sql_tu02_bonus)
End If

If Len(sql_rt01_bonus) <> 0 Then
    Fn_SQLExec (sql_rt01_bonus)
End If


'---------------------------------------------------------
' 틀린 수 이벤트 발생
'---------------------------------------------------------
If rst = 0 Then
    
    If cFTP.Profile.bAlarm1 Then
        sSql = "select count(*) from tu02 where userid='" & gUserid & "' and x>o "
        
        cnt = Fn_SQLExec(sSql).rs(0)
        If (cnt Mod cFTP.Profile.CntOfAlarm1) = 0 And cnt > 0 And cnt <> oldcnt1 Then
            oldcnt1 = cnt
            MsgBox "모르는 문제수: [" & cnt & "]          ", vbInformation
'            oldcnt1 = cnt
        Else
            oldcnt1 = 0
        End If
        
    End If
    
    If cFTP.Profile.bAlarm2 Then
        sSql = "select count(*) from tu02 where userid='" & gUserid & "' and x=o and x>0 "
    
        cnt = Fn_SQLExec(sSql).rs(0)
        If (cnt Mod cFTP.Profile.CntOfAlarm2) = 0 And cnt > 0 And cnt <> oldcnt2 Then
            oldcnt2 = cnt
            MsgBox "잘 모르는 문제수: [" & cnt & "]          ", vbInformation
        Else
            oldcnt2 = 0
            
        End If
    End If

'    sSql = "select count(*) from tu02 where userid='guserid & "' and o>x and x>0" '실수
   
    
End If

If cQuiz.chk > 0 And cQuiz.isNew = False Then
    
    If cFTP.Profile.bAlarm5 Then
        sSql = "select count(*) from tu02 where userid='" & gUserid & "' and chk>0 "
        
        cnt = Fn_SQLExec(sSql).rs(0)
        If (cnt Mod cFTP.Profile.CntOfAlarm5) = 0 And cnt > 0 And cnt <> oldcnt5 Then
            oldcnt5 = cnt
            MsgBox "Check된 문제수: [" & cnt & "]           ", vbInformation
        Else
            oldcnt5 = 0
        End If
    End If

End If

'---------------------------------------------------------
' 틀린 수 이벤트 발생 끝
'---------------------------------------------------------


Con_Close

If (rst = 0) Then
    'Call cTg01.add_x(gUserid, cQuiz.subj)
ElseIf rst = 1 Then
    'Call cTg01.add_o(gUserid, cQuiz.subj)
    
    If cQuiz.isNew = False And cQuiz.chk > 0 Then
        Call cTg01.add_chk(gUserid, cQuiz.subj)
    End If

End If

Incress = True

End Function
'사용되지 않고 있음
'Function Incress2(rst As Integer) As Boolean '성공 True
'Dim ymd As String
'Dim sql1 As String
'Dim affected As Integer
'Dim Cnt As Long
'Static oldcnt1 As Long
'Static oldcnt2 As Long
'Static oldcnt5 As Long
'
'Dim preGanGyek As Long
'
'Incress2 = False
''On Error Resume Next
'ymd = Format(Now, "YYYYMMDD")
'
'If rst = 1 Then
'    cFTP.Profile.CntOfStreatOutNow = 0
'Else
'    If pgBar.Value = 0 Then
'        cFTP.Profile.CntOfStreatOutNow = cFTP.Profile.CntOfStreatOutNow + 1
'    Else
'        cFTP.Profile.CntOfStreatOutNow = 0
'    End If
'End If
'
'Read_TU02
'
'If rst = 1 Then
'    '-----------------------
'    '맞은경우
'    '-----------------------
'    '1 예정일자가 오늘보다 같거나 큰경우 eg> 오늘 20040717 예정일:20041125
'    '  1.1 예정일=20041125+(update_ymd-ymd)
'    '  1.2 인터벌: 인터벌+(update_ymd-ymd)
'    '-----------------------
'    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
'    '  1.0 만약 예정일과 오늘과의 차이가 인터벌보다 크면 그값의 배를 인터벌로 함
'    '  1.1 예정일=ymd+(인터벌)
'    '  1.2 인터벌: (인터벌+1)*2
'    '
'    '-----------------------
'    '1 예정일자가 오늘보다 같거나 큰경우 eg> 오늘 20040717 예정일:20041125
'    If cQuiz.reserve_ymd >= ymd Then
'
'    '  1.1 예정일=20041125+(update_ymd-ymd)
'        cQuiz.reserve_ymd = date2Str(DateAdd("d", Abs(DateDiff("d", str2Date(cQuiz.update_ymd), str2Date(ymd))), str2Date(cQuiz.reserve_ymd)))
'
'    '  1.2 인터벌: 인터벌+(update_ymd-ymd)
'        cQuiz.gangyek = cQuiz.gangyek + Abs(DateDiff("d", str2Date(cQuiz.update_ymd), str2Date(ymd)))
'
'
'    '-----------------------
'    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
'    ElseIf cQuiz.reserve_ymd < ymd Then
'       preGanGyek = Abs(DateDiff("d", str2Date(cQuiz.reserve_ymd), str2Date(ymd)))
''       If preGanGyek > cQuiz.gangyek Then
''           cQuiz.gangyek = preGanGyek * 2
''       End If
'        cQuiz.gangyek = cQuiz.gangyek '?DDD
'        'date2Str (DateAdd("d", CLng(cQuiz.gangyek * heavydamping(-0.005, DateDiff("d", str2Date(ymd), cQuiz.reserve_ymd))) + 1, str2Date(ymd)))
'
'    '  1.1 예정일=ymd+(인터벌)
'        cQuiz.reserve_ymd = date2Str(DateAdd("d", cQuiz.gangyek, str2Date(ymd)))
'    '  1.2 인터벌: (인터벌+1)*2
'        cQuiz.gangyek = (cQuiz.gangyek + 1) * 2
'
'    '-----------------------
'    End If
'
'    sSql = "update tu02 set gangyek=" & cQuiz.gangyek & ",reserve_ymd='" & cQuiz.reserve_ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
'Else
'    '-----------------------
'    '틀린경우
'    '-----------------------
'    '1 예정일자가 오늘보다 같거나 큰경우 eg> 오늘 20040717 예정일:20041125
'    '  1.1 예정일=예정일-1
'    '  1.2 인터벌: 인터벌 그대로
'    '-----------------------
'    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
'    '  1.1 예정일=ymd
'    '  1.2 인터벌: 인터벌-1
'    '-----------------------
'    '-----------------------
'    '틀린경우
'    '-----------------------
'    '1 예정일자가 오늘보다 같거나 큰경우 eg> 오늘 20040717 예정일:20041125
'    If cQuiz.reserve_ymd >= ymd Then
'    '  1.1 예정일=오늘 날짜 -1
'        cQuiz.reserve_ymd = date2Str(DateAdd("d", -1, str2Date(cQuiz.reserve_ymd)))
'    '  1.2 인터벌: 인터벌
'        cQuiz.gangyek = cQuiz.gangyek
'    '-----------------------
'    '2 예정일자가 오늘보다 작은 경우 eg> 오늘 20040717 예정일: 20040615
'    ElseIf cQuiz.reserve_ymd < ymd Then
'    '  1.1 예정일=오늘 날짜-1
'        cQuiz.reserve_ymd = ymd
'
'    '  1.2 인터벌: 인터벌-1
'        cQuiz.gangyek = cQuiz.gangyek - 1
'        If cQuiz.gangyek < 0 Then
'            cQuiz.gangyek = 0
'        End If
'    '-----------------------
'    End If
'    sSql = "update tu02 set gangyek=" & cQuiz.gangyek & ",reserve_ymd='" & cQuiz.reserve_ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
'End If
'
'Con_Open
'
'affected = Fn_SQLExec(sSql, , True).nrow
'
'Debug.Assert affected = 1
'If affected = 0 Then
'    MsgBox "메뉴에서 시험지>어카운트갱신 을 실행 시킨 후 다시 문제를 푸세요-F7 버튼-", vbExclamation
'
'    Con_Close
'    cmdClose_Click
'    Exit Function
'ElseIf affected < 0 Then
'    Con_Close
'    cmdClose_Click
'    Exit Function
'End If
'
'affected = Fn_SQLExec(sql1).nrow
'
'Debug.Assert affected = 1
'
'If cQuiz.chk > 0 Then
'    sSql = "update tu02 set chk=chk+1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
'    sql1 = "update tp03 set chk=chk+1 ,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "'  and chasu=" & gChasu & "  and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
'
'
'    affected = Fn_SQLExec(sSql).nrow
'    Debug.Assert affected = 1
'
'    affected = Fn_SQLExec(sql1).nrow
'    Debug.Assert affected = 1
'
'    cQuiz.chk = 0
'    chk.Value = vbUnchecked
'End If
'
'If TP.chkm > 0 And rst = 1 Then
'    sSql = "update tu02 set chk=chk-1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
'    'sql1 = "update tp03 set chk=chk-1 ,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "'  and chasu=" & gChasu & "  and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
'    affected = Fn_SQLExec(sSql, , True).nrow
'
'    If err.Number <> 0 And err.Number <> -2147467259 Then
'        Debug.Assert False
'    Else
'        Debug.Assert affected = 1
'    End If
'
'    'affected = Fn_SQLExec(sql1, , True).nrow
'
'End If
'
'If TP.xm > 0 And rst = 1 Then
'
'    sSql = "update tu02 set x=x-1,o=o-1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq
'    'sql1 = "update tp03 set x=x-1,o=o-1,update_ymd='" & ymd & "' where userid='" & gUserid & "' and pocketnm='" & gPocket & "'  and chasu=" & gChasu & "  and subj='" & cQuiz.subj & "' and num=" & cQuiz.num
'    affected = Fn_SQLExec(sSql, , True).nrow
'    'affected = Fn_SQLExec(sql1, , True).nrow
'    If err.Number <> 0 And err.Number <> -2147467259 Then
'        Debug.Assert False
''        MsgBox "이미 Pass한 시험지를 여러 번 푼것같습니다.", vbExclamation
'    Else
'        Debug.Assert affected = 1
'
'    End If
'    Debug.Assert affected = 1
'
'End If
'
'
''---------------------------------------------------------
'' 틀린 수 이벤트 발생
''---------------------------------------------------------
'If rst = 0 Then
'
'    If cFTP.Profile.bAlarm1 Then
'        sSql = "select count(*) from tu02 where userid='" & gUserid & "' and x>o "
'
'        Cnt = Fn_SQLExec(sSql).rs(0)
'        If (Cnt Mod cFTP.Profile.CntOfAlarm1) = 0 And Cnt > 0 And Cnt <> oldcnt1 Then
'            oldcnt1 = Cnt
'            MsgBox "모르는 문제수: [" & Cnt & "]          ", vbInformation
''            oldcnt1 = cnt
'        Else
'            oldcnt1 = 0
'        End If
'
'    End If
'
'    If cFTP.Profile.bAlarm2 Then
'        sSql = "select count(*) from tu02 where userid='" & gUserid & "' and x=o and x>0 "
'
'        Cnt = Fn_SQLExec(sSql).rs(0)
'        If (Cnt Mod cFTP.Profile.CntOfAlarm2) = 0 And Cnt > 0 And Cnt <> oldcnt2 Then
'            oldcnt2 = Cnt
'            MsgBox "잘 모르는 문제수: [" & Cnt & "]          ", vbInformation
'        Else
'            oldcnt2 = 0
'
'        End If
'    End If
'
''    sSql = "select count(*) from tu02 where userid='guserid & "' and o>x and x>0" '실수
'
'
'End If
'
'If cQuiz.chk > 0 And cQuiz.isNew = False Then
'
'    If cFTP.Profile.bAlarm5 Then
'        sSql = "select count(*) from tu02 where userid='" & gUserid & "' and chk>0 "
'
'        Cnt = Fn_SQLExec(sSql).rs(0)
'        If (Cnt Mod cFTP.Profile.CntOfAlarm5) = 0 And Cnt > 0 And Cnt <> oldcnt5 Then
'            oldcnt5 = Cnt
'            MsgBox "Check된 문제수: [" & Cnt & "]           ", vbInformation
'        Else
'            oldcnt5 = 0
'        End If
'    End If
'
'End If
'
''---------------------------------------------------------
'' 틀린 수 이벤트 발생 끝
''---------------------------------------------------------
'
'
'Con_Close
'
'If (rst = 0) Then
'    Call cTg01.add_x(gUserid, cQuiz.subj)
'ElseIf rst = 1 Then
'    Call cTg01.add_o(gUserid, cQuiz.subj)
'
'    If cQuiz.isNew = False And cQuiz.chk > 0 Then
'        Call cTg01.add_chk(gUserid, cQuiz.subj)
'    End If
'
'End If
'
'Incress2 = True
'
'End Function



Function slectnotin(subj As String, nm As Long, lstchr As String, Optional 초기화 As Boolean = False, Optional mode As String, Optional ByRef refSeq As Long, Optional ByRef ans_a As String) As String
Dim rs1 As New ADODB.Recordset
Dim cnt As Long
Dim i As Long
Dim str1 As String

Static bogi_list As New Collection
Static bogi_list_seq As New Collection

If 초기화 Then
    Set bogi_list = Nothing
    Set bogi_list = New Collection
    
    Set bogi_list_seq = Nothing
    Set bogi_list_seq = New Collection
    
Else
    '성능향상
    slectnotin = bogi_list(1)
    bogi_list.Remove (1)
    
    refSeq = "" + bogi_list_seq(1)
    bogi_list_seq.Remove (1)
    
    Exit Function 'slectnotin 과 refSeq가 쌍이다.(답,순서)
End If


Randomize
If lstchr = "~" Then
    sSql = "select a,quiz,seq from vq01 where subj='" & subj & "' and seq not in (" & nm & ") and (a like '-%' or a like '~%') and a<>'" & ans_a & "'"
ElseIf lstchr = "-" Then
    sSql = "select a,quiz,seq from vq01 where subj='" & subj & "' and seq not in (" & nm & ") and (a like '-%' or a like '~%') and a<>'" & ans_a & "'"
Else
    sSql = "select a,quiz,seq from vq01 where subj='" & subj & "' and seq not in (" & nm & ") and a<>'" & ans_a & "'"
End If
'Set rs1 = New ADODB.Recordset

Dim result As ut_bRecordSet

result = Fn_SQLExec(sSql)

cnt = result.nrow

Set rs1 = result.rs

' rs1.Open sSql, Con, 1, 3
' cnt = rs1.RecordCount
Dim refSeq_cache As String

If 초기화 Then

   Dim mmB As String, mmC As String, mmD As String, mmE As String
   
   Dim NB As Long, nC As Long, nD As Long, nE As Long
   Dim nCnt As Long, nT As Long
   nCnt = 0
   
   NB = Round15(rndVal(0, cnt - 1))
   
   nC = Round15(rndVal(0, cnt - 1))
   Do While (NB = nC)
      nC = Round15(rndVal(0, cnt - 1))
      nCnt = nCnt + 1
      If nCnt > 10 Then Exit Do
   Loop
   
   nCnt = 0
   
   nD = Round15(rndVal(0, cnt - 1))
   Do While (NB = nD Or nC = nD)
      nD = Round15(rndVal(0, cnt - 1))
      nCnt = nCnt + 1
      If nCnt > 10 Then Exit Do
   Loop
   
   nCnt = 0
   
   nE = Round15(rndVal(0, cnt - 1))
   Do While (NB = nE Or nC = nE Or nD = nE)
      nE = Round15(rndVal(0, cnt - 1))
      nCnt = nCnt + 1
      If nCnt > 10 Then Exit Do
   Loop
   
   If NB > nC Then
      nT = NB
      NB = nC
      nC = nT
   End If
   If NB > nD Then
      nT = NB
      NB = nD
      nD = nT
   End If
   If NB > nE Then
      nT = NB
      NB = nE
      nE = nT
   End If
   
   If nC > nD Then
      nT = nC
      nC = nD
      nD = nT
   End If
   If nC > nE Then
      nT = nC
      nC = nE
      nE = nT
   End If
   
   If nD > nE Then
      nT = nD
      nD = nE
      nE = nT
   End If
     
     rs1.Move NB
     str1 = rs1(0).Value
     refSeq = rs1(2).Value
     
     slectnotin = str1
'     Call bogi_list.Add(str1)
    
     rs1.Move (nC - NB)
     str1 = rs1(0).Value
     refSeq_cache = rs1(2).Value
     
     Call bogi_list.Add(str1)
     Call bogi_list_seq.Add(refSeq_cache)
    
     
     rs1.Move (nD - nC)
     str1 = rs1(0).Value
     refSeq_cache = rs1(2).Value
     
     Call bogi_list.Add(str1)
     Call bogi_list_seq.Add(refSeq_cache)
    
    If mode = "5" Or mode = "E" Then
      rs1.Move (nE - nD)
      str1 = rs1(0).Value
      refSeq_cache = rs1(2).Value
      Call bogi_list.Add(str1)
      Call bogi_list_seq.Add(refSeq_cache)
      
     End If
Else
   i = Round15(rndVal(0, cnt - 1))
   rs1.Move i
   slectnotin = rs1(0)
   refSeq = rs1(2).Value
End If
End Function

Function nansu(nm As Long, bisNew As Boolean, rng As Long) As Long
Dim LRng As Long
Dim preLastNewNm As Long
Dim preHangsu As Long
Dim swapnm As Long
Dim bIsNew2 As Boolean
Dim depth1 As Long
Dim depthnow As Long

If bisNew = False And rng > 0 Then
    Randomize
    LRng = Fix(Rnd * 10000) Mod rng
    preLastNewNm = gLastNew
    preHangsu = gHangSu + LRng
    swapnm = cMNIQ.getF(preLastNewNm, preHangsu, gOrder, bisNew, depth1, depthnow)
    
    If swapnm > nm And bIsNew2 = False And preLastNewNm = gLastNew Then
        Call doswapnm(nm, swapnm)
    End If
End If
End Function

Sub doswapnm(nm1 As Long, nm2 As Long)
Con_Open
On Error GoTo ErrTrap
    Con.BeginTrans
    
    sSql = "update tp03 set num=-1 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " And num = " & nm2 & ""
    If Fn_SQLExec(sSql).Error Then
    
    End If
    sSql = "update tp03 set num=" & nm2 & " where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " And num = " & nm1 & ""
    Fn_SQLExec (sSql)
    sSql = "update tp03 set num=" & nm1 & " where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num=-1"
    Fn_SQLExec (sSql)
Con.CommitTrans
Con_Close
    Exit Sub
ErrTrap:
Debug.Assert False
Con.RollbackTrans
Con_Close

End Sub


Sub delprocess(nm As Long)
Dim affected As Long
'Debug.Assert con.STATE = 1
Con_Open
Con.BeginTrans
sSql = "delete from tp03 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num=" & nm & ""
affected = Fn_SQLExec(sSql).nrow
Debug.Assert affected = 1
sSql = "update tp03 set num=num-1 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num>" & nm & ""
affected = Fn_SQLExec(sSql).nrow
'Debug.Assert affected > 0

If affected = 0 Then
    If nm > 1 And (nm Mod gOrder) <> 1 Then
        Con.RollbackTrans
        Con.BeginTrans
        sSql = "update tp03 set num=" & nm + 1 & " where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num=" & nm - 1 & ""
        affected = Fn_SQLExec(sSql).nrow
        Debug.Assert affected = 1
        
        sSql = "update tp03 set num=" & nm - 1 & " where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num=" & nm & ""
        affected = Fn_SQLExec(sSql).nrow
        Debug.Assert affected = 1
        
        sSql = "update tp03 set num=" & nm & " where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & " and num=" & nm + 1 & ""
        affected = Fn_SQLExec(sSql).nrow
        Debug.Assert affected = 1
        
        Con.CommitTrans
        cmdPre_Click
        
        MsgBox "죄송합니다. 다시 한 번 삭제를 시도하세요.", vbExclamation + vbOKOnly
    Else
        Con.RollbackTrans
        
        MsgBox "죄송합니다. 마지막 문제는 삭제할 수 없습니다.", vbExclamation + vbOKOnly
    End If
    
    Con_Close
    
    
    '해결방법 전문제와 바꿔치기 함.
Else
    Con.CommitTrans
    Con_Close
End If

Dim maxNum As Integer
Con_Open
sSql = "select max(num) from tp03 where userid='" & gUserid & "' and pocketnm='" & gPocket & "' and chasu=" & gChasu & ""
maxNum = Fn_SQLExec(sSql).rs(0)
If gOrder > maxNum Then
    gOrder = maxNum '20040701 추가
    If gOrder = 1 Then
        gOrder = 2
    End If
    Write_TU03
End If
Con_Close
Exit Sub
ErrTrap:
    Debug.Assert False
 Con.RollbackTrans
Con_Close

End Sub

Sub MagicHintView()

frmQuiz.cls
If cFTP.Profile.ShowWhiteSpot And cQuiz.isNew = False Then
    If cQuiz.ans = "A" Then
        frmQuiz.PSet (0, 0), vbWhite 'A
    ElseIf cQuiz.ans = "B" Then
        frmQuiz.PSet (frmQuiz.Width - 70, 0), vbWhite 'B
    ElseIf cQuiz.ans = "C" Then
        frmQuiz.PSet (frmQuiz.Width - 70, frmQuiz.Height - 70 - Toolbar1.Height), vbRed 'C
    Else
        frmQuiz.PSet (0, frmQuiz.Height - 70 - Toolbar1.Height), vbRed 'D
    End If
End If

If cFTP.Profile.ShowWhiteSpot And cFTP.Profile.checkWhenResize And cQuiz.isNew = False Then
    cQuiz.chk = 1
    chk.Value = vbChecked
End If

End Sub


Sub read_tg01()

End Sub

Sub write_tg01()


End Sub

'가상키 코드 값(16진수)  키
'VK_LBUTTON 1
'VK_RBUTTON 2
'VK_CANCEL   3   Ctrl-Break
'VK_MBUTTON 4
'VK_BACK 8   Backspace
'VK_TAB  9   Tab
'VK_CLEAR    0C  NumLock이 꺼져 있을 때의 5
'VK_RETURN   0D  Enter
'VK_SHIFT    0x10  Shift
'VK_CONTROL  0x11  Ctrl
'VK_MENU 0x12  Alt
'VK_PAUSE    0x13  Pause
'VK_CAPITAL  0x14  Caps Lock
'VK_ESCAPE   0x1B  Esc
'VK_SPACE    0x20  스페이스
'VK_PRIOR    0x21  PgUp(page up)
'VK_NEXT 0x22  PgDn (page down)
'VK_END  0x23  End
'VK_HOME 0x24  Home
'VK_LEFT 0x25  왼측 커서 이동키
'VK_UP   0x26  위쪽 커서 이동키
'VK_RIGHT    0x27  오른쪽 커서 이동키
'VK_DOWN     0x28  아래쪽 커서 이동키
'VK_SELECT 0x29
'VK_PRINT    0x2A
'VK_EXECUTE  0x2B
'VK_SNAPSHOT 0x2C  Print Screen
'VK_INSERT   0x2D  Insert
'VK_DELETE   0x2E  Delete
'VK_HELP 2F
'0x30    0 key
'0x31    1 key
'0x32    2 key
'0x33    3 key
'0x34    4 key
'0x35    5 key
'0x36    6 key
'0x37    7 key
'0x38    8 key
'0x39    9 key
'0x41    A key
'0x42    B key
'0x43    C key
'0x44    D key
'0x45    E key
'0x46    F key
'0x47    G key
'0x48    H key
'0x49    I key
'0x4A    J key
'0x4B    K key
'0x4C    L key
'0x4D    M key
'0x4E    N key
'0x4F    O key
'0x50    P key
'0x51    Q key
'0x52    R key
'0x53    S key
'0x54    T key
'0x55    U key
'0x56    V key
'0x57    W key
'0x58    X key
'0x59    Y key
'0x5A    Z key
'VK_LWIN     5B  왼쪽 윈도우 키
'VK_RWIN     5C  오른쪽 윈도우 키
'VK_APP  5D  Application 키
'VK_NUMPAD0~VK_NUMPAD9   60~69   숫자 패드의 0~9
'VK_MULTIPLY     6A  숫자 패드의 *
'VK_ADD  6B  숫자 패드의 +
'VK_SEPARATOR    6C
'VK_SUBTRACT     6D  숫자 패드의 -
'VK_DECIMAL  6E  숫자 패드의 .
'VK_DIVIDE   6F  숫자 패드의 /
'VK_F1 0x70 F1 key
'VK_F2 0x71 F2 key
'VK_F3 0x72 F3 key
'VK_F4 0x73 F4 key
'VK_F5 0x74 F5 key
'VK_F6 0x75 F6 key
'VK_F7 0x76 F7 key
'VK_F8 0x77 F8 key
'VK_F9 0x78 F9 key
'VK_F10 0x79 F10 key
'VK_F11 0x7A F11 key
'VK_F12 0x7B F12 key
'VK_F13~VKF24    7C~87   평션키 F13~F24
'VK_NUMLOCK  90  Num Lock
'VK_SCROLL   91  Scroll Lock
'0x92-96 OEM specific
'0x97-9F Unassigned
'VK_LSHIFT 0xA0 Left SHIFT key
'VK_RSHIFT 0xA1 Right SHIFT key
'VK_LCONTROL 0xA2 Left CONTROL key
'VK_RCONTROL 0xA3 Right CONTROL key
'VK_LMENU 0xA4 Left MENU key
'VK_RMENU 0xA5 Right MENU key
'VK_BROWSER_BACK 0xA6 Browser Back key
'VK_BROWSER_FORWARD 0xA7 Browser Forward key
'VK_BROWSER_REFRESH 0xA8 Browser Refresh key
'VK_BROWSER_STOP 0xA9 Browser Stop key
'VK_BROWSER_SEARCH 0xAA Browser Search key
'VK_BROWSER_FAVORITES 0xAB Browser Favorites key
'VK_BROWSER_HOME 0xAC Browser Start and Home key
'VK_VOLUME_MUTE 0xAD Volume Mute key
'VK_VOLUME_DOWN 0xAE Volume Down key
'VK_VOLUME_UP 0xAF Volume Up key
'VK_MEDIA_NEXT_TRACK 0xB0 Next Track key
'VK_MEDIA_PREV_TRACK 0xB1 Previous Track key
'VK_MEDIA_STOP 0xB2 Stop Media key
'VK_MEDIA_PLAY_PAUSE 0xB3 Play/Pause Media key
'VK_LAUNCH_MAIL 0xB4 Start Mail key
'VK_LAUNCH_MEDIA_SELECT 0xB5 Select Media key
'VK_LAUNCH_APP1 0xB6 Start Application 1 key
'VK_LAUNCH_APP2 0xB7 Start Application 2 key
'0xB8-B9 Reserved
'VK_OEM_1 0xBA Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ';:' key
'VK_OEM_PLUS 0xBB For any country/region, the '+' key
'VK_OEM_COMMA 0xBC For any country/region, the ',' key
'VK_OEM_MINUS 0xBD For any country/region, the '-' key
'VK_OEM_PERIOD 0xBE For any country/region, the '.' key
'VK_OEM_2 0xBF Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '/?' key
'VK_OEM_3 0xC0 Used for miscellaneous characters; it can vary by keyboard.Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '`~' key
'0xC1-D7 Reserved
'0xD8-DA Unassigned
'VK_OEM_4 0xDB Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '[' key
'VK_OEM_5 0xDC Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '' key
'VK_OEM_6 0xDD Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '] ' key
'VK_OEM_7 0xDE Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the 'single-quote/double-quote' key
'VK_OEM_8 0xDF Used for miscellaneous characters; it can vary by keyboard.
'0xE0 Reserved 0xE1 OEM specific
'VK_OEM_102 0xE2 Either the angle bracket key or the backslash key on the RT 102-key keyboard
'0xE3-E4 OEM specific
'VK_PROCESSKEY 0xE5 IME PROCESS key
'0xE6 OEM specific VK_PACKET
'0xE7 Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in  KEYBDINPUT ,  SendInput , WM_KEYDOWN , and  WM_KEYUP
'0 xE8 Unassigned
'0xE9-F5 OEM specific
'VK_ATTN 0xF6 Attn key
'VK_CRSEL 0xF7 CrSel key
'VK_EXSEL 0xF8 ExSel key
'VK_EREOF 0xF9 Erase EOF key
'VK_PLAY 0xFA Play key
'VK_ZOOM 0xFB Zoom key
'VK_NONAME 0xFC Reserved
'VK_PA1 0xFD PA1 key
'VK_OEM_CLEAR 0xFE Clear key


Private Sub TmrAfterTTS_focus_Timer()
If TmrAfterTTS_exit Then Exit Sub
'TmrAfterTTS_focus.Enabled = False

Dim fgHwnd As Long
fgHwnd = GetForegroundWindow()

If frmQuiz.Visible And frmQuiz.hwnd <> fgHwnd Then

'VK_LEFT 0x25  왼측 커서 이동키
'VK_UP   0x26  위쪽 커서 이동키
'VK_RIGHT    0x27  오른쪽 커서 이동키
'VK_DOWN     0x28  아래쪽 커서 이동키

    Dim b0x25 As Boolean, b0x26 As Boolean, b0x27 As Boolean, b0x28 As Boolean
    
    b0x25 = IsKeyHitDown(&H25)
    b0x26 = IsKeyHitDown(&H26)
    b0x27 = IsKeyHitDown(&H27)
    b0x28 = IsKeyHitDown(&H28)
    
    If (b0x25 Or b0x26 Or b0x27 Or b0x28) Then
        
        If GetTickCount() < time_tick1 Then '0.1초 이내에 두번 이상 진입이 불가능하게 한 코드이다.
'''            TmrAfterTTS_focus.Enabled = True
            Exit Sub
        End If
        time_tick1 = GetTickCount() + 150
        
        TmrAfterTTS_pushCount = TmrAfterTTS_pushCount + 1
        
        If 2 < TmrAfterTTS_pushCount And fgHwnd <> trust_firstHWND_wndPlayer Then
            '다른윈도우에서 작업을 정상적으로 하고 있는 경우로 봄
            TmrAfterTTS_focus.Enabled = False
            Exit Sub
        End If
        
    End If
    
    If b0x25 Then
        '이전 문제
        If cmdPre.Enabled Then
            Call cmdPre_Click
        End If
    ElseIf b0x26 Then '위로
      If optA.Value Then
        If optE.Visible Then
            optE.Value = True
        ElseIf optD.Visible Then
            optD.Value = True
        ElseIf optC.Visible Then
            optC.Value = True
            Else
            optB.Value = True
        End If
      ElseIf optB.Value Then
        optA.Value = True
      ElseIf optC.Value Then
        optB.Value = True
      ElseIf optD.Value Then
        optC.Value = True
      ElseIf optE.Value Then
        optD.Value = True
      Else
        If optE.Visible Then
            optE.Value = True
        ElseIf optD.Visible Then
            optD.Value = True
                ElseIf optC.Visible Then
            optC.Value = True
                ElseIf optB.Visible Then
            optB.Value = True
        End If
        
        'frmQuiz.SetFocus
        SetForegroundWindow (frmQuiz.hwnd)
        
        trust_firstHWND_wndPlayer = fgHwnd
        TmrAfterTTS_pushCount = 0
        
      End If
      
    ElseIf b0x27 Then '다음문제
       Call cmdNext_Click
    ElseIf b0x28 Then '아래로
        If optA.Value Then
            optB.Value = True
        ElseIf optB.Value Then
            
                If optC.Visible Then
                    optC.Value = True
                Else
                    optA.Value = True
                End If
             
        ElseIf optC.Value Then
         
            If optD.Visible Then
                optD.Value = True
            Else
                optA.Value = True
            End If
        
        ElseIf optD.Value Then
        
            If optE.Visible Then
            optE.Value = True
            Else
            optA.Value = True
            End If
        Else
            optA.Value = True
            
            'frmQuiz.SetFocus
            SetForegroundWindow (frmQuiz.hwnd)
            trust_firstHWND_wndPlayer = fgHwnd
            TmrAfterTTS_pushCount = 0
        End If
    End If
        
    'frmQuiz.SetFocus
    'TmrAfterTTS_focus.Enabled = False
    TmrAfterTTS_focus.Enabled = True
Else
    DoEvents
    
    If frmQuiz.Visible = False Then
        TmrAfterTTS_focus.Enabled = False
    Else
        If frmQuiz.hwnd <> fgHwnd Then
            TmrAfterTTS_focus.Enabled = True
        Else
            TmrAfterTTS_focus.Enabled = False
        End If
    End If
End If

End Sub

Private Sub tmrTG_Timer()
Static Sec60 As Integer
Static OldNum As Long
Static SecAll As Long
Static SecPart As Long
Static MinPart As Long
Static HourPart As Long
Static DayPart As Long
Static reSecAll As Long

Dim frmMsgBox As frmMsgBox60sec

DoEvents

'tmrTG.Enabled = False
'tmrHourGlass.Enabled = False

If frmMain Is Nothing Then Exit Sub

If False Then
    tmrTG.Enabled = False
End If

If cQuiz.num = OldNum Then
    '같으면 이벤트가 발생하지 않음
    SecAll = SecAll + 1
    reSecAll = reSecAll + 1
Else
    '초기화
    Sec60 = 0
    SecAll = 0
    reSecAll = 0
End If

If SecAll < 60 Then
    frmMain.sbStatusBar.Panels(4).Text = "" & SecAll & "초"
ElseIf SecAll < 3600 Then '1시간이내
    MinPart = SecAll \ 60
    SecPart = SecAll - MinPart * 60
    frmMain.sbStatusBar.Panels(4).Text = "" & MinPart & "분 " & SecPart & "초"
ElseIf SecAll < 86400 Then '3600 * 24 Then
    '세션 타임아웃을 시킨다.(너무오랫동안 접속이 없으면 프로그램을 종료)
    'End
    
    If reSecAll > 3600 Then
    
        Set frmMsgBox = New frmMsgBox60sec
        Load frmMsgBox
        
        frmMsgBox.setMsg "프로그램이 너무 오랫동안 구동되어 종료됩니다."
        
        TmrAfterTTS_focus.Enabled = False
        TmrAfterTTS_exit = True
        frmMsgBox.Show vbModal, frmMain
        TmrAfterTTS_exit = False
        
        If frmMsgBox.btn_flag = 0 Then
            Unload fMainForm
            End
        Else
            reSecAll = 0
            Set frmMsgBox = Nothing
        End If

    End If
    
'    HourPart = SecAll \ 3600
'    MinPart = (SecAll - HourPart * 3600) \ 60
'    SecPart = SecAll - HourPart * 3600 - MinPart * 60
'    frmMain.sbStatusBar.Panels(4).Text = "" & HourPart & "시간 " & MinPart & "분 " & SecPart & "초"
Else
    DayPart = SecAll \ (86400)
    HourPart = (SecAll - DayPart * 86400) \ 3600
    MinPart = (SecAll - DayPart * 86400 - HourPart * 3600) \ 60
    SecPart = SecAll - DayPart * 86400 - HourPart * 3600 - MinPart * 60
    frmMain.sbStatusBar.Panels(4).Text = "" & DayPart & "일 " & HourPart & "시간 " & MinPart & "분 " & SecPart & "초"
    
    If reSecAll > 3600 Then
    
        Set frmMsgBox = New frmMsgBox60sec
        Load frmMsgBox
        
        frmMsgBox.setMsg "프로그램이 너무 오랫동안 구동되어 종료됩니다."
        
        TmrAfterTTS_focus.Enabled = False
        TmrAfterTTS_exit = True
        frmMsgBox.Show vbModal, frmMain
        TmrAfterTTS_exit = False
        
        If frmMsgBox.btn_flag = 0 Then
            Unload fMainForm
            End
        Else
            reSecAll = 0
            Set frmMsgBox = Nothing
        End If

    End If
End If
If Timer1.Enabled Then
    Call cTg01.addsec(gUserid, cQuiz.subj)
ElseIf cFTP.Profile.bSetTimeOutStudy = False And cQuiz.isNew Then
'새로운것 공부하는시간 어느정도 허용
    If Sec60 < 60 Then
        Call cTg01.addsec(gUserid, cQuiz.subj)
    Else
        Sec60 = 61
    End If
    Sec60 = Sec60 + 1
    
ElseIf cFTP.Profile.bSetTimeOut = False And cQuiz.isNew = False Then
'시간안가고 공부하는것 (복습) 어느정도 공부시간으로 허용
    If Sec60 < 60 Then
        Call cTg01.addsec(gUserid, cQuiz.subj)
    Else
        Sec60 = 61
    End If
    Sec60 = Sec60 + 1
    
End If

OldNum = cQuiz.num

End Sub


Public Function selRndTip() As String

Static colTip As New Collection
Dim lRs As ADODB.Recordset
Dim i As Long
Dim cnt As Long
Dim strTip As String


If colTip.count <> 0 Then
    i = Round15(rndVal(1, colTip.count))
    selRndTip = colTip(i)
    Exit Function
End If

Randomize

sSql = "select * from ttp1 "

Dim result As ut_bRecordSet

result = Fn_SQLExec(sSql)

Set lRs = result.rs

cnt = result.nrow

Do Until lRs.EOF
    strTip = lRs(1) & " ─ " & lRs(2)
    colTip.Add strTip
    lRs.MoveNext
Loop

    
    i = Round15(rndVal(1, colTip.count))
    selRndTip = colTip(i)

End Function

Sub opt_MouseOut()
    iCatch = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'툴바 사용시 유의: align속성은 런타임시 폼로드시 포지션시켜주도록 한다 기본속성으로
'정하면 폼(윈도우)객체가 언로드시 툴바의 left top값을 읽어오는 부분에서 오류가 간혹
'발생될 수도 있다. 2010.01.05 by PGS
Dim idx As Integer
idx = Button.Index

If idx = 1 Then
    If optA.Visible Then
        optA.Value = True
        cmdNext_Click
    End If
ElseIf idx = 2 Then
    If optB.Visible Then
        optB.Value = True
        cmdNext_Click
    End If
ElseIf idx = 3 Then
    If optC.Visible Then
        optC.Value = True
        cmdNext_Click
    End If
ElseIf idx = 4 Then
    If optD.Visible Then
        optD.Value = True
        cmdNext_Click
    End If
ElseIf idx = 5 Then
    If optE.Visible Then
        optE.Value = True
        cmdNext_Click
    End If
   
End If
End Sub



Private Function isExistsEE() As Boolean
    Con_Open
    sSql = "select exists(select 1 from th01 where subj='" & cQuiz.subj & "' and seq=" & cQuiz.seq & " and userid='@') result"
    Set rs = Fn_SQLExec(sSql, , , True).rs
    If Not rs.EOF Then
        If 1 = rs("result") Then
            isExistsEE = True
        Else
            isExistsEE = False
        End If
    End If
    Con_Close
End Function


Public Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean

   DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS

End Function


Private Function FolderExists(sFullPath As String) As Boolean
Dim myFSO As Object
Set myFSO = CreateObject("Scripting.FileSystemObject")
FolderExists = myFSO.FolderExists(sFullPath)
End Function


Public Function IsKeyHitDown(System_Int32 As Integer) As Boolean

    Dim v As Integer
    v = GetAsyncKeyState(System_Int32)
    
    If v <> 0 Then
        IsKeyHitDown = True
    Else
        IsKeyHitDown = False
    End If
    
End Function

Private Sub wbQuizMain_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Debug.Print URL

End Sub

Private Sub wbQuizMain_GotFocus()
wbQuizFocused = True
End Sub

Private Sub wbQuizMain_LostFocus()
wbQuizFocused = False
End Sub

