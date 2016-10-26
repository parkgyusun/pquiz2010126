VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "POCKETQUIZ"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10860
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox imgSplitter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4725
      Left            =   7740
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4725
      ScaleWidth      =   105
      TabIndex        =   9
      Top             =   1260
      Width           =   105
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   5
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.FileListBox objListFile 
      Height          =   285
      Left            =   5580
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox AnimationListBox 
      Enabled         =   0   'False
      Height          =   255
      Left            =   660
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   5700
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Timer tmrHourGlass 
      Interval        =   1000
      Left            =   6165
      Top             =   1260
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   5190
      Left            =   0
      TabIndex        =   6
      Top             =   315
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   9155
      _Version        =   393217
      Indentation     =   88
      Style           =   7
      ImageList       =   "imlToolbarIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5805
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2280
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10860
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10860
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   2085
         TabIndex        =   3
         Tag             =   "1044"
         Top             =   60
         Width           =   3210
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   "1043"
         Top             =   60
         Width           =   2010
      End
      Begin VB.Shape shpTitle 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         FillColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   5625
         Top             =   45
         Width           =   555
      End
      Begin VB.Shape shpTitle 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         FillColor       =   &H0000FFFF&
         Height          =   240
         Index           =   1
         Left            =   6435
         Top             =   45
         Width           =   555
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1845
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7470
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11906
            Text            =   "상태"
            TextSave        =   "상태"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "오후 10:12"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1305
            TextSave        =   "오후 10:12"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "경과시간"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   5160
      Left            =   2070
      TabIndex        =   4
      Top             =   345
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   9102
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   1005
      Index           =   0
      Left            =   9810
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   1005
      ExtentX         =   1764
      ExtentY         =   1764
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
   Begin SHDocVwCtl.WebBrowser wb2 
      Height          =   1000
      Left            =   -1000
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   -1000
      Width           =   1000
      ExtentX         =   1764
      ExtentY         =   1764
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
   Begin SHDocVwCtl.WebBrowser wb3 
      Height          =   1000
      Left            =   -1000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   -1000
      Width           =   1000
      ExtentX         =   1764
      ExtentY         =   1764
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
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   1005
      Index           =   1
      Left            =   9510
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1005
      ExtentX         =   1764
      ExtentY         =   1764
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
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   1005
      Index           =   2
      Left            =   9000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1005
      ExtentX         =   1764
      ExtentY         =   1764
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
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1001"
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "1002"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "1003"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "1004"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "1005"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "1006"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "1007"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "1008"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "1009"
      Visible         =   0   'False
      Begin VB.Menu mnuEditUndo 
         Caption         =   "1010"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "1011"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "1012"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "1013"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "1014"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "1015"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "1016"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1017"
      Visible         =   0   'False
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1018"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1019"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1020"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1021"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1022"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1023"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "1024"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "1025"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "1026"
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "1027"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1028"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "1029"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "1030"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1031"
      End
   End
   Begin VB.Menu mnuPocketQuizP 
      Caption         =   "열공모드(&C)"
      Begin VB.Menu mnuRefresh 
         Caption         =   "새로고침"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSchedule 
         Caption         =   "학습계획(시험지 만들기)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuDeletePaperAuto 
         Caption         =   "시험지 자동 정리"
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "일정보기"
      End
      Begin VB.Menu mnuMakePaper 
         Caption         =   "시험지만들기"
         Shortcut        =   ^{F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuAccount 
         Caption         =   "어카운트 갱신"
         Shortcut        =   {F7}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUserUpdate 
         Caption         =   "회원정보변경"
      End
      Begin VB.Menu mnuEnv 
         Caption         =   "환경설정(&R)"
      End
      Begin VB.Menu mnuBar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApproach 
         Caption         =   "학습상태점검"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuChgUsr 
         Caption         =   "사용자변경"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "종료(&X)"
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "즐공모드(&G)"
      Begin VB.Menu mnuCharUse 
         Caption         =   "캐릭터사용"
      End
      Begin VB.Menu mnuEndSetting 
         Caption         =   "종료시간설정"
      End
      Begin VB.Menu mnuGame1 
         Caption         =   "게임(넷트리스)"
      End
      Begin VB.Menu mnuSepa9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "검색"
      End
   End
   Begin VB.Menu mnuTv 
      Caption         =   "트리뷰"
      Begin VB.Menu mnuEXAM 
         Caption         =   "시험보기(&S)"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMake 
         Caption         =   "시험지만들기"
         Begin VB.Menu mnuReviewCreate 
            Caption         =   "복습문제생성"
         End
         Begin VB.Menu mnuBar200501301 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "틀린문제"
            Index           =   0
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "체크문제"
            Index           =   1
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "안푼문제"
            Index           =   2
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "맞은문제"
            Index           =   3
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "틀린+체크"
            Index           =   4
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "틀린+체크+안품"
            Index           =   5
         End
      End
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReTry 
         Caption         =   "재시험"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "삭제(&X)"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "시험지명변경"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuQuickStart 
         Caption         =   "학습생략 후 바로시작"
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "학습생략"
      End
      Begin VB.Menu mnuBar55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsetPocket 
         Caption         =   "클립보드 문항추가"
      End
      Begin VB.Menu mnuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperty 
         Caption         =   "속성(&R)"
      End
   End
   Begin VB.Menu mnuEdit2 
      Caption         =   "폰트크기"
      Begin VB.Menu mnuPt 
         Caption         =   "10pt"
         Index           =   0
      End
      Begin VB.Menu mnuPt 
         Caption         =   "14pt"
         Index           =   1
      End
      Begin VB.Menu mnuPt 
         Caption         =   "20pt"
         Index           =   2
      End
      Begin VB.Menu mnuPt 
         Caption         =   "22pt"
         Index           =   3
      End
      Begin VB.Menu mnuPt 
         Caption         =   "24pt"
         Index           =   4
      End
      Begin VB.Menu mnuPt 
         Caption         =   "26pt"
         Index           =   5
      End
      Begin VB.Menu mnuPt 
         Caption         =   "28pt"
         Index           =   6
      End
      Begin VB.Menu mnuPt 
         Caption         =   "30pt"
         Index           =   7
      End
      Begin VB.Menu mnuPt 
         Caption         =   "32pt"
         Index           =   8
      End
      Begin VB.Menu mnuPt 
         Caption         =   "34pt(한자권장)"
         Index           =   9
      End
      Begin VB.Menu mnuPt 
         Caption         =   "36pt"
         Index           =   10
      End
      Begin VB.Menu mnuPt 
         Caption         =   "38pt"
         Index           =   11
      End
      Begin VB.Menu mnuPt 
         Caption         =   "40pt"
         Index           =   12
      End
   End
   Begin VB.Menu mnu_help_main 
      Caption         =   "도움말"
      Begin VB.Menu mnu_help 
         Caption         =   "도움말"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "프로그램정보"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const MYAGENTUSE_ON = False
'#Const MYAGENTUSE_ON = True

'Microsoft Agent Control 2.0(C:\Windows\MSAgent\agentctl.dll) 누락시킴 2012.12.26 Windows7이상에서 호환성 결여문제로...박규선
Option Explicit
Option Base 0

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

'============================
'Agent 설정
'----------------------------
Const BalloonOn = 1
Const SizeToText = 2
Const AutoHide = 4
Const AutoPace = 8
'============================

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Dim cMNIQ As New clsMNIQ

Dim CharLoaded As Boolean

#If MYAGENTUSE_ON Then
Public Character As IAgentCtlCharacterEx
#End If

Public frmEndSetting As frmEndSetting
Public frmVq01 As frmVq01

Const sglSplitLimit = 50
Public cQuiz As New clsQuiz
Dim frmMsgBox As frmMsgBox60sec

#If MYAGENTUSE_ON Then
'Private WithEvents MyAgent As AgentObjectsCtl.Agent
#End If

Private Sub Command1_Click()

End Sub

Private Sub Form_GotFocus()

If frmQuiz.Visible Then
 frmQuiz.TmrAfterTTS_focus.Enabled = False
End If

End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)

If mnuEdit2.Visible Then
    On Error Resume Next
    Call frmQuiz.picMain.SetFocus
    Call frmQuiz.Form_KeyDown(keycode, Shift)
End If


End Sub

Private Sub Form_Load()

#If MYAGENTUSE_ON Then
    mnuCharUse.Enabled = True
#Else
    mnuCharUse.Enabled = False
    mnuCharUse.Visible = False
#End If

mnuGame1.Visible = False

Set frmEndSetting = New frmEndSetting

frmEndSetting.Hide
Set frmEndSetting.parent = Me
If cFTP.Profile.endSettingLogin = "1" Then
    
    frmEndSetting.txtEndMin.Text = cFTP.Profile.endSettingLoginDefaultMin
    TmrAfterTTS_exit = True
    frmEndSetting.Show vbModal
    TmrAfterTTS_exit = False
    frmEndSetting.isForce = True
End If


'mnuCharUse_Click '사용

lblTitle(0).BackStyle = 0 '투명
lblTitle(1).BackStyle = 0 '투명

Call shpTitle(0).Move(0, 0)
Call shpTitle(1).Move(lblTitle(1).Left, 0)

    Set status = Me.sbStatusBar.Panels(1)

    mnuEdit2.Visible = False
    
    Me.Caption = Me.Caption & " ver." & App.Major & "." & App.Minor & "." & App.Revision
'    LoadResStrings Me
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    imgSplitter.Left = GetSetting(App.Title, "Settings", "Splitter", imgSplitter.Left)
    If imgSplitter.Left < 950 Then
       imgSplitter.Left = 950
    End If
    mnuTv.Visible = False
    mnuRefresh_Click
    
    'mnuApproach.Enabled = Not isempty_tg01()
    
    
    cFTP.Profile.save
    
    If cFTP.Profile.FontSize = 8 Then
        mnuPt(0).Checked = True
    ElseIf cFTP.Profile.FontSize = 9 Then
        mnuPt(1).Checked = True
    ElseIf cFTP.Profile.FontSize = 10 Then
        mnuPt(2).Checked = True
    ElseIf cFTP.Profile.FontSize = 11 Then
        mnuPt(3).Checked = True
    ElseIf cFTP.Profile.FontSize = 12 Then
        mnuPt(4).Checked = True
    ElseIf cFTP.Profile.FontSize = 13 Then
        mnuPt(5).Checked = True
    ElseIf cFTP.Profile.FontSize = 14 Then
        mnuPt(6).Checked = True
    Else
        mnuPt(7).Checked = True
    End If
    
'    tvTreeView.Indentation = 100
If tvTreeView.Nodes.count = 1 Then
    Call mnuSchedule_Click
End If

If gLogonCnt = 0 Then
    Call mnuUserUpdate_Click
End If

Con_Close

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cFTP.Profile.setPoP3 Then
    If MsgBox("프로그램을 종료하시겠습니까?", vbYesNo + vbQuestion) = vbNo Then Cancel = 1
    Exit Sub
End If

Screen.MousePointer = vbHourglass
Con_Close


If InStr(LCase(STRCON), "mdb") > 0 Then
compactDB
End If

Screen.MousePointer = vbDefault

Call cTU01.logout(gUserid)  '로그아웃으로 설정한다.

End Sub

'Private Sub Form_Paint()
'    lvListView.View = val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
'    Select Case lvListView.View
'        Case lvwIcon
'            tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
'        Case lvwSmallIcon
'            tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
'        Case lvwList
'            tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
'        Case lvwReport
'            tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
'    End Select
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


frmEndSetting.isForce = True
Unload frmEndSetting


    'close all sub forms
    For i = Forms.count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        
        SaveSetting App.Title, "Settings", "Splitter", imgSplitter.Left
        
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
    
    cFTP.Profile.save
    Con_Close
    If Not gChangUser Then '사용자 변경에 의한 언로드인경우는 컨넥션을 끊지 않는다.
        Con.Close
    End If

If Not gChangUser Then
'사용자가 변경되는 경우는 이것을 행하면 안됨.
    End '정상종료가 안되는 경우 해결=>전역변수,환경변수,프로파일=>스태틱을 다루는 일이 중요함
End If
End Sub
Sub compactDB()
On Error GoTo ErrTrap
  Dim ldbengine As New DBEngine
  Dim file1 As String
  Dim file2 As String
  file1 = getPerfectFile(App.Path, MDBFILENM)
  file2 = getPerfectFile(App.Path, TEMPMDB)
  If cFTP.Profile.compactDB Then
        
        
        ldbengine.CompactDatabase file1, file2, , , GPWD
        Kill file1
        Call FileCopy(file2, file1)
        Kill file2
    
    End If
Exit Sub
ErrTrap:
If cFTP.Profile.setPoP2 Then '경미한 에러 보이지 않음
    MsgBox err.Description, vbExclamation
End If

  
End Sub




Private Sub Form_Resize()
    On Error Resume Next
    gMainOnResize = True
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left + 50 '서숙련차장님 오류 발견 제안
    If mnuEdit2.Visible Then
        Call frmQuiz.Move(lvListView.Left, lvListView.Top, lvListView.Width, lvListView.Height)
'        frmQuiz.Top = 0
'        frmQuiz.Width = lvListView.Width
'        frmQuiz.Height = lvListView.Height
    End If
    
    gMainOnResize = False
    
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
gMainOnResize = True
'With imgSplitter
        picSplitter.Move imgSplitter.Left, imgSplitter.Top, 60, imgSplitter.Height - 20
    'End With
    picSplitter.Visible = True
    mbMoving = True
    picSplitter.ZOrder 0
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = x + imgSplitter.Left - 25
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit - 1500 Then
            picSplitter.Left = Me.Width - sglSplitLimit - 1500
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left + 25
    picSplitter.Visible = False
    mbMoving = False
    gMainOnResize = False '마우스다운에서 true 설정과 짝이며 메인창 resize 이벤트에서도 동일 설정값이 있다.
End Sub


Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub


Sub SizeControls(x As Single)
    On Error Resume Next
    

    '너비를 설정합니다.
    If x < 50 Then
        x = 50
    End If
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    tvTreeView.Width = x
    imgSplitter.Left = x - 50
    lvListView.Left = x + 30
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40

    shpTitle(0).Width = tvTreeView.Width
    shpTitle(1).Left = lvListView.Left + 20
    shpTitle(1).Width = lvListView.Width - 40

    
    '위쪽을 설정합니다.
  

'    If tbToolBar.Visible Then
'        tvTreeView.Top = tbToolBar.Height + picTitles.Height
'    Else
'        tvTreeView.Top = picTitles.Height
'    End If

  lvListView.Top = tvTreeView.Top
    

    '높이를 설정합니다.
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    

    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
    
    If mnuEdit2.Visible Then
        frmQuiz.Move lvListView.Left, lvListView.Top, lvListView.Width, lvListView.Height
'        frmQuiz.PicMain.Width = frmQuiz.Width - frmQuiz.PicMain.Left * 2
'        frmQuiz.PicMain.Height = frmQuiz.Height - frmQuiz.PicMain.Top * 2
      
    End If
    
    Dim frm As Form
    Dim ctl As Object
    
    For Each frm In Forms
        If frm.Tag = "FORM" Then
            If frm.Visible Then
            
            'Set frm = ctl
                frm.Move lvListView.Left, lvListView.Top, lvListView.Width, lvListView.Height
            End If
        End If
    Next
    

End Sub

Private Sub mnu12pt_Click()

End Sub


Private Sub mnu_about_Click()
frmAbout.Show
End Sub

Private Sub mnu_help_Click()
frmHelp.Show
End Sub

Private Sub mnuAccount_Click()

If Con.State = 1 Then
    insert_totalaccount2
Else
    Con_Open
    insert_totalaccount2
    Con_Close
End If
End Sub

Private Sub mnuApproach_Click()

If mnuEdit2.Visible Then
    
End If

Call cTg01.addsec(gUserid, "", True)

Load frmApproach
SetParent frmApproach.hwnd, Me.hwnd

frmApproach.Visible = True

SizeControls imgSplitter.Left

End Sub

Private Sub mnuCharUse_Click()
    'merlin.acs
On Error Resume Next
#If MYAGENTUSE_ON Then
MyAgent.Tag = "mutex"
If mnuCharUse.Checked = False Then
    mnuCharUse.Checked = True
Else
    mnuCharUse.Checked = False
End If

If mnuCharUse.Checked = False Then
    If Character Is Nothing Then
        MsgBox "(캐릭터 파일).acs파일이 없습니다.", vbInformation + vbOKOnly, "알림"
        Exit Sub
    End If
    Character.Hide ' True
    Exit Sub
End If
#End If

On Error Resume Next

#If MYAGENTUSE_ON Then
Set Character = Nothing
MyAgent.Characters.Unload "CharacterID"
#End If


On Error GoTo errHandler
Dim ri As Long

objListFile.FileName = "*.acs" 'if file extionname is .acs_ is listed .. this is error

'Debug.Print objListFile.ListCount

If objListFile.ListCount = 0 Then
    Exit Sub
End If

Dim FileName As String


ri = makeRandInt(1, objListFile.ListCount)
FileName = objListFile.List(ri - 1)
While InStrRev(LCase$(FileName), ".acs") <> Len(FileName) - 3
    'file exition name acs_ error defense codding
    ri = makeRandInt(1, objListFile.ListCount)
    FileName = objListFile.List(ri - 1)
Wend

'MsgBox filename
On Error Resume Next
#If MYAGENTUSE_ON Then
MyAgent.Characters.Load "CharacterID", App.Path + "\" + FileName

Set Character = MyAgent.Characters("CharacterID")

If Not (Character Is Nothing) Then
    CharLoaded = True
Else
    CharLoaded = False
End If

AnimationListBox.Clear
Dim AnimationName As Variant
For Each AnimationName In Character.AnimationNames
    AnimationListBox.AddItem AnimationName
Next

'Character.LanguageID = &H409 '영어
Character.LanguageID = &H412 '한국어

Character.SoundEffectsOn = False

Character.Left = (Me.Left + Me.Width) / Screen.TwipsPerPixelX - 100 ' 100
Character.Top = (Me.Top + Me.Height) / Screen.TwipsPerPixelX - 100 ' 100

'-- Show the character
Character.Show


Character.Balloon.Style = Character.Balloon.Style And (Not AutoHide)
Character.Balloon.Style = Character.Balloon.Style Or AutoPace
Character.Balloon.Style = Character.Balloon.Style Or SizeToText
#End If

Exit Sub

errHandler:
    
End Sub

Private Sub mnuChgUsr_Click()
Call cTU01.logout(gUserid)
gChangUser = True
Main
gChangUser = False
End Sub

Private Sub mnuDelete_Click()
Dim vData As Variant
Dim Pos As Integer
Dim selPocket As String
Dim selChasu  As String

vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '실행중인작업있는경우
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))

If cFTP.Profile.setPoP1 Then '삭제시 질문
    If MsgBox("[" & selPocket & "] 시험지를 삭제 하시겠습니까? (차수 = " & selChasu & ")", vbQuestion + vbYesNo) = vbYes Then
        If mnuEdit2.Visible = True Then
            If selPocket = gPocket Then
              MsgBox "현재 시험보고 있는 중입니다.", vbExclamation, "삭제불가"
            Else
                deletePaper selPocket, selChasu
                mnuRefresh_Click
            End If
        Else
            LockWindowUpdate (Me.hwnd)
            deletePaper selPocket, selChasu
            mnuRefresh_Click
            LockWindowUpdate (0&)
        End If
    End If
Else
    LockWindowUpdate (Me.hwnd)
    deletePaper selPocket, selChasu
    mnuRefresh_Click
    LockWindowUpdate (0&)
End If
End Sub

Private Sub mnuDeletePaperAuto_Click()

Dim churi_cnt  As Long
churi_cnt = 0
Con_Open
sSql = "select if(dd.od=1,2,dd.od) selOrder,dd.hangsu -1 tmpHangSu,cc.cnt RSTOTAL,ab.pocketnm,ab.chasu,cc.userid from "
sSql = sSql + " (select b.pocketnm ,b.chasu,b.userid from tp02 a inner join tp02 b on (a.pcode=0 and a.code=b.pcode and a.userid=b.userid) where b.userid='" & gUserid & "'"
sSql = sSql + "  and b.hidden=0 order by b.pcode, b.create_ymd,b.code) ab"
sSql = sSql + " inner join (select count(*) cnt,pocketnm,chasu,userid from tp03 group by pocketnm,chasu,userid) cc"
sSql = sSql + " on (ab.pocketnm=cc.pocketnm and ab.chasu=cc.chasu and ab.userid=cc.userid)"
sSql = sSql + " inner join tu03 dd on (ab.pocketnm=dd.pocketnm and ab.chasu=dd.chasu and ab.userid=dd.userid) where dd.hangsu>dd.od and ab.pocketnm not like '!%'"
Set rs = Fn_SQLExec(sSql).rs

'1단계인것중에 진도 다 뺀것은 자식까지 다 지운다 1단계인지는 부모의 부모는 0인 것이다
'+----------+-----------+---------+---------------------+-------+--------+
'| selOrder | tmpHangSu | RSTOTAL | pocketnm            | chasu | userid |
'+----------+-----------+---------+---------------------+-------+--------+
'|       11 |        92 |     107 | 1 회차 20150816 ~   |     1 | 박규선 |
'|       10 |       171 |      81 | 1 회차 20151230 ~   |     1 | 박규선 |
'|       10 |        97 |      84 | 1 회차 20160107 ~   |     1 | 박규선 |
'|       12 |        43 |     128 | 2 회차  ~           |     4 | 박규선 |
'|       20 |       400 |     377 | 2 회차 20150803 ~   |     1 | 박규선 |
'|       10 |       200 |      20 | 2 회차 20151218 ~   |     2 | 박규선 |
'|       10 |       200 |      80 | 2 회차 20151231 ~   |     1 | 박규선 |
'|       12 |        12 |     120 | 3 회차  ~           |     4 | 박규선 |
'|       19 |       276 |     360 | 3 회차 20150806 ~   |     1 | 박규선 |
'|       10 |       131 |      80 | 3 회차 20160101 ~   |     1 | 박규선 |
'|       15 |       201 |     196 | 6 회차 20150815 ~   |     1 | 박규선 |
'|       10 |       200 |      90 | 73 회차 20140924 ~  |     1 | 박규선 |
'|       10 |        94 |      90 | 74 회차 20140927 ~  |     1 | 박규선 |
'|       10 |        20 |      90 | 80 회차 20141015 ~  |     1 | 박규선 |
'|      126 |     31752 |     126 | 복습문제20160107    |     8 | 박규선 |
'|       10 |       200 |      10 | 복습문제20160114    |     1 | 박규선 |
'|      289 |    167042 |     289 | 복습문제20160123    |     1 | 박규선 |
'|      283 |    160178 |     283 | 복습문제20160209    |     2 | 박규선 |
'|       20 |       800 |      20 | 복습문제20160216    |     1 | 박규선 |
'+----------+-----------+---------+---------------------+-------+--------+
'19 rows in set (0.50 sec)
'19개 가 그 대상이며 거기서 조사해서 적절히 삭제하면 되겠다.
'중요한것(느낌표가 처음에 들어간 시험지)는 삭제하지 않는다.

Dim tmpHangsu As Long, totalHangSu As Long

Do Until rs.EOF()
    totalHangSu = cMNIQ.geth0(rs(2) + 1, 0, rs(0)) - 1
    
    If rs(1) >= totalHangSu Then
        Call deletePaper(rs(3), rs(4))
        churi_cnt = churi_cnt + 1
    End If
    rs.MoveNext
Loop
Con_Close

If churi_cnt = 0 Then
    MsgBox "정리 완료 [" & churi_cnt & "]건", vbOKOnly, Me.Caption
Else
    MsgBox "정리 완료 [" & churi_cnt & "]건", vbOKOnly, Me.Caption
    mnuRefresh_Click
End If

End Sub

Private Sub mnuInsetPocket_Click()

Dim vData As Variant
Dim Pos As Integer
Dim selPocket As String
Dim selChasu  As String

vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '실행중인작업있는경우
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))
'
'If selPocket = gPocket Then
'    MsgBox "현재 시험보고 있는 중입니다.", vbExclamation, "삭제불가"
'Else
    
    Call PocketAdd(selPocket, selChasu)
    
'End If

End Sub

Private Sub PocketAdd(ByVal selPocket As String, ByVal selChasu As String)

Dim str As String

str = Clipboard.GetText()

Dim arr() As String

arr = Split(str, ",")

Dim col_arr As New Collection

Dim i As Long

On Error Resume Next
For i = LBound(arr) To UBound(arr)
    arr(i) = Trim(arr)
    Call col_arr.Add(arr(i), arr(i))
Next
On Error GoTo 0

If col_arr.count <> UBound(arr) - LBound(arr) + 1 Then
    '중복오류가 생겼음 문제 해결
    ReDim arr(0 To col_arr.count() - 1) As String
    
    For i = 0 To col_arr.count() - 1
        arr(i) = col_arr(i + 1)
    Next
    
End If

Dim max_num As Long

max_num = getTq03_max_num(selPocket, selChasu)

Dim affected As Long

Dim pre_max_num As Long
pre_max_num = max_num '1인경우 해당 1은 먼저 지우도록 해야 한다.

If pre_max_num = 1 Then '기존에 한건이 1번이 었으면 그 한 건은 지운다.
    sSql = "delete from tp03 where pocketnm='" + selPocket + "' and chasu=" + selChasu + " and num = 1"
    affected = Fn_SQLExec(sSql).nrow
    max_num = max_num - 1
End If

Dim lRs As ADODB.Recordset

sSql = "select ta.subj,ta.seq,ta.quiz from vq01 ta inner join tu02 tb on (ta.subj=tb.subj and ta.seq=tb.seq) where tb.userid='" + gUserid + "' and ta.mode='1' and ta.quiz " + vbNewLine + "in (" & selectSeriesArray(arr) & ")"

Set lRs = Fn_SQLExec(sSql).rs

Dim churi As Long
Dim quiz_tmp As String

Do Until lRs.EOF()
    max_num = max_num + 1
    churi = churi + 1
    sSql = "insert into tp03(pocketnm,chasu,userid,num,subj,seq,o,X,chk,update_ymd) values('" & selPocket & "'," & selChasu & ",'" & gUserid & "'," & max_num & ",'" & lRs(0) & "'," & lRs(1).Value & ",0,0,0,date_format(current_date,'%Y%m%d'))"
    
    err.Clear
    On Error Resume Next
    quiz_tmp = lRs(2).Value
    If IsNull(quiz_tmp) = False Then
        Call col_arr.Remove(quiz_tmp) 'Sunday가 2개조회되어 없는 데이터를 지우려하는 경우 5번 오류가 발생되어 비정상적으로 죽을 수 있어 위에서 On Error Resume Next를 했다.
    End If
    
    If err.Number <> 5 Then
        affected = Fn_SQLExec(sSql).nrow
    Else
        churi = churi - 1
    End If
    lRs.MoveNext
Loop

If 0 < col_arr.count Then
    
    Call MsgBox("" & churi & "건 정상 처리" & vbNewLine & col_arr.count & "건은 처리 되지 못하였습니다.(이유: 기초 데이터 에 존재하지 않음)" & vbNewLine & "미처리된 내역은 클립보드에 넣었습니다.", vbExclamation, Me.Caption)
    
    Dim strMichuri As String
    
    Dim j As Long
    For j = 1 To col_arr.count
        If Len(strMichuri) = 0 Then
            strMichuri = col_arr(j)
        Else
            strMichuri = strMichuri + "," + col_arr(j)
        End If
        
        If 0 < Len(col_arr(j)) Then
            sSql = "insert into TR01(quiz_required,userid,ymd,chk) values('" + col_arr(j) + "','" + gUserid + "',date_format(current_date,'%Y%m%d'),0) "
            affected = Fn_SQLExec(sSql).nrow
        End If
            
    Next
    
    '입력이 추후 요구 될 수 있는 문항임
    'create table TR01 (seq int(11) NOT NULL auto_increment,quiz_required mediumtext,userid varchar(50),ymd varchar(8),chk int(2),PRIMARY KEY(seq)) ENGINE=MyISAM ;
    '
    
    Clipboard.Clear
    Clipboard.SetText (strMichuri)
    
Else
    Call MsgBox("" & churi & "건 정상 처리되었습니다.", vbExclamation, Me.Caption)
End If

If 0 < churi Then
    '시험지명을 바꾼다.
    If InStr(selPocket, "!") = 0 Then
        sSql = "update tp02 set pocketnm=concat('!',pocketnm) where pocketnm='" & selPocket & "' and chasu=" & selChasu & " and userid='" & gUserid & "'"
        affected = Fn_SQLExec(sSql).nrow
        
        sSql = "update tp03 set pocketnm=concat('!',pocketnm) where pocketnm='" & selPocket & "' and chasu=" & selChasu & " and userid='" & gUserid & "'"
        affected = Fn_SQLExec(sSql).nrow
        
        sSql = "update tu03 set pocketnm=concat('!',pocketnm) where pocketnm='" & selPocket & "' and chasu=" & selChasu & " and userid='" & gUserid & "'"
        affected = Fn_SQLExec(sSql).nrow
        
        selPocket = "!" & selPocket
    Else
        
    End If
    
    '시험보기를 바로 진행한다.
    
    If MsgBox("시험보기를 하시겠습니까? [" & selPocket & "]", vbQuestion + vbYesNo) = vbYes Then
        vDataGlobal = selPocket & " " & selChasu
    Else
        vDataGlobal = ""
    End If
    
    If InStr(selPocket, "!") > 0 Or vDataGlobal <> "" Then
        mnuRefresh_Click
    End If
End If

Debug.Print (sSql)

End Sub

Function getTq03_max_num(selPocket, selChasu) As Long
    sSql = "select max(num) num from tp03 where pocketnm='" + selPocket + "' and chasu=" + selChasu
    getTq03_max_num = Fn_SQLExec(sSql).rs(0)
End Function

Private Sub mnuEndSetting_Click()
'frmEndSetting.txtEndMin.Enabled = False
frmEndSetting.txtEndMin = cFTP.Profile.endSettingSetMin
'frmEndSetting.txtEndMin.Enabled = True
TmrAfterTTS_exit = True
frmEndSetting.Show vbModal
TmrAfterTTS_exit = False
frmEndSetting.isForce = True
End Sub

Private Sub mnuEnv_Click()
    Load frmProperty
    
    Set frmProperty.frmMain = Me
    
    If mnuEdit2.Visible Then
        Set frmProperty.frmQuiz = frmQuiz
    End If
    frmProperty.cboFont.Text = frmQuiz.picMain.FontName
 
    frmProperty.Show
End Sub

Public Sub mnuEXAM_Click()
gbMnuExamClick = True
Dim vData As Variant
Dim Pos As Integer
Dim MNU As Menu

If mnuEdit2.Visible Then
    MsgBox "기존 시험을 종료 후 실행하세요!.", vbExclamation
    gbMnuExamClick = False
    Exit Sub
End If

If Module1.vDataGlobal = "" Then
    vData = getPN_CH(tvTreeView.SelectedItem.key)
Else
    vData = Module1.vDataGlobal
End If
Module1.vDataGlobal = ""

If vData = "" Then
    gbMnuExamClick = False
    Exit Sub '실행중인작업있는경우
End If
Pos = InStrRev(vData, " ")

If InStr(tvTreeView.SelectedItem.key, "_") Then
    
    gPocket = Left(vData, Pos - 1)
    gChasu = Trim(Mid(vData, Pos))
    
'    Debug.Assert con.State = 1 '?
    Con_Open
    sSql = "select chkm, xm from tp02 where pocketnm='" & gPocket & "' and chasu=" & gChasu & " and userid='" & gUserid & "'"
    
    Set rs = Fn_SQLExec(sSql).rs
        
    chkm = rs(0)
    xm = rs(1)
    
    If Me.mnuEdit2.Visible = True Then
        Unload frmQuiz
        mnuEdit2.Visible = False
    End If
    Load frmQuiz
    
    frmQuiz.wbQuizMain.Navigate ("about:blank")
    
    Set frmQuiz.frmMain = Me
    mnuApproach.Enabled = True

    Me.mnuEdit2.Visible = True
    
    If cFTP.Profile.FontSize = 0 Then
        cFTP.Profile.FontSize = 14
    End If
    
    If cFTP.Profile.TimeOutSec = 0 Then
        cFTP.Profile.TimeOutSec = 10
    End If
    
    If cFTP.Profile.TimeOutSecStudy = 0 Then
        cFTP.Profile.TimeOutSecStudy = 10
    End If
    
    
    setfontsize_menu_bysize cFTP.Profile.FontSize
    
    Call SetParent(frmQuiz.hwnd, Me.hwnd)
'    frmQuiz.Visible = True
'    frmQuiz.PicMain.Width = frmQuiz.Width - frmQuiz.PicMain.Left * 2
'    frmQuiz.PicMain.Height = frmQuiz.Height - frmQuiz.PicMain.Top * 2
'    frmQuiz.Visible = True
    picSplitter.ZOrder 0
    SizeControls imgSplitter.Left
'    frmQuiz.Left = lvListView.Left + imgSplitter.Width
'    frmQuiz.Top = lvListView.Top + imgSplitter.Width
'    frmQuiz.ZOrder 0
''     frmQuiz.Width = lvListView.Width
'    frmQuiz.Height = lvListView.Height
   lblTitle(1).Caption = "시험명: " & gPocket & "      응시자: " & gUserid
   lblTitle(1).Tag = lblTitle(1).Caption
   
     gNowNum = frmQuiz.Read_TU03
     If gbReadTu03 = False Then
        gbMnuExamClick = False
        Exit Sub
     End If
     Call LockWindowUpdate(frmQuiz.hwnd)
     
     If cFTP.Profile.endSettingQuiz = "1" Then
'        frmEndSetting.txtEndMin.Enabled = False
        frmEndSetting.txtEndMin.Text = cFTP.Profile.endSettingQuizDefaultMin
'        frmEndSetting.txtEndMin.Enabled = True
TmrAfterTTS_exit = True
        frmEndSetting.Show vbModal
        TmrAfterTTS_exit = False
        frmEndSetting.isForce = True
    End If
     
     frmQuiz.nowVisible = True
     frmQuiz.Visible = True '폼을보이면리사이즈이벤트가 자동 발생함.
     frmQuiz.nowVisible = False
'     For Each mnu In mnuPt
'        If mnu.Checked Then
'            mnuPt_Click mnu.Index
'        End If
'     Next
'    frmQuiz.Read_TU03
Dim RSTOTAL As Long, selOrder As Long

sSql = "SELECT COUNT(*) cnt FROM TP03 WHERE POCKETNM='" & gPocket & "' AND CHASU=" & gChasu & " AND USERID='" & gUserid & "' "
RSTOTAL = Fn_SQLExec(sSql).rs(0)

sSql = "select od from tu03 where POCKETNM='" & gPocket & "' and userid='" & gUserid & "' AND CHASU=" & gChasu
selOrder = Fn_SQLExec(sSql).rs(0)

frmQuiz.TotalQuizCount = cMNIQ.geth0(RSTOTAL + 1, 0, selOrder) - 1
frmQuiz.pgBarJindo.Max = frmQuiz.TotalQuizCount
    
    If frmQuiz.quizDisp = False Then
        gbMnuExamClick = False
        Exit Sub 'ddd여기서 한 번 더 호출함...
    End If
    Call LockWindowUpdate(0&)
    frmQuiz.Visible = True
    'frmQuiz.cmdNext.SetFocus
    frmQuiz.picMain.SetFocus
    
    frmQuiz.mnu_auto_tts.Checked = cFTP.Profile.bChkTTSuse
    
'    frmQuiz.optA.SetFocus ' .pic1.SetFocus
End If




Con_Close
gbMnuExamClick = False
End Sub

Public Function getPN_CH(key As String) As String

'##1
Dim conbef As Integer
If Con.State = 1 Then
    conbef = 1
Else
    Con_Open
    conbef = 0
End If

Dim SSQL1 As String

SSQL1 = "SELECT * FROM TP02 WHERE USERID='" & gUserid & "' AND CODE=" & Left(key, Len(key) - 1) & ""
Dim s1 As String
Dim rs1 As ADODB.Recordset
Set rs1 = Fn_SQLExec(SSQL1).rs
If rs1 Is Nothing Then Exit Function
If rs1.EOF = False Then
    s1 = rs1("POCKETNM") & " " & rs1("CHASU")
    getPN_CH = s1
End If
'##2
If conbef = 1 Then
Else
    Con_Close
End If


End Function

Private Sub mnuExit2_Click()
Unload Me
End Sub

Private Sub mnuGame1_Click()
'기능정의
'1. 상단화살표키를 두 번 누른것은 오늘의 review노트에 기록된다.
'2. 나중에 리뷰노트에 있는 것을 볼 수 있다.
'3. 중국어도 가능하다.(다국어)
'4. 게임종료시 리뷰노트를 보시겠습니까? 라는 질문이 나온다.
'5. 리뷰노트를 보는 경우는 웹페이지와 바로가기가 있게 된다.
'6. 틀린것으로 카운팅되는 경우
' 6.1 pass된 경우
' 6.2 위로 한칸 쌓아진 경우
' 틀린것으로 카운팅되는 것과 동시에 내일 추천학습 항목으로 기록된다.
' 틀린경우나 맞은 경우 하루 한번만(최초 1회) 카운팅 된다.
' 트레이닝 모드인 경우 시간기록(소요)-> 답맞춘경우 평균소요시간을 알 수 있도록 저장된다.
' 훌륭한 디자인(입체 모양,혹은 쎄련된 색상의 단순하면서도)
' 훌륭한 게임음악
' 배틀 가능하게 하기
' 공격 아이템(상대방 얼게 하기 3초간)
' 상대방 미리보기 아이템 5번 삭제하기
' 소리재생은 둘다 다이렉트 사운드로 재생하게 한다.
' 상대방 소리재생 막기
' 상대방 위로 증가시키게 하기
' 바이러스를먹은 경우는 내 아이템의 한개가 하나가 가려진다.
' 상대방에게 가려지게 하는 바이러스를 줄수 있다.
' 배틀에서 이기면 승점이 부여된다.
' 게임은 점점 빨라진다.
' 상대가 져야 배틀의 승리가 된다.
' 힌트사용시 정답 알려준다.
' 틀리면 그냥 쌓인다.
' 맞추면 사라지며 맞추는데 소요된 시간이 기록되고 키보드아래화살표를 이용한 경우와 space를 이용하는 경우 소리 안나고 나게 한다.
' 미리 대기하는 것이 있다.
' 3개의 생명이 있게 된다.
' 트레이닝 모드와 게임모드로 나눈다.
' 트레이닝 모드로 하여 맞추는 경우 빨리 맞추게끔 하는 훈련을 시키고
' 등장하면서 소리나오기 하는 기능은 위(↑)로 버튼을 한번 누르는 것으로 한다.
' 1층이 올라오는 경우도 있다.
' 아무런키 조작이 없이 구멍에 빠지면 한개의 생명을 잃게 된다.
' 위로 쌇아 올린게 꼭대기에 닳으면 생명이 하나 잃게되며 다시 게임이 시작된다.
' 늪에 빠져도 생명을 하나 잃게 되며 다시 게임이 시작된다.
' 떨어지는 속도 산출방식: 목적지까지의 거리와 시간의 두배로 떨어진다.
' 사라지는 효과는 같이 한몸을 이루는테두리가 생긴후 사라진다.
' 틀린횟수가 기록된다
' 맞춘 횟수가 기록된다.
' 총소요시간(맞추는데 소요된 시간의 누계)와 맞춘횟수에 의해 평균 소요시간[초]을 구할 수 있다.
' 틀린것으로 그냥 쌓인 경우는 다음에 맞을때 어떻게 맞든 소리가 나온다.
' 배틀게임으로 내가 모르는 문제인 어려운 것을 상대방에게 보낼 수 있다.
' 영어단어 말고도 산술연산을 문제로 할 수 있다. (가령 3+3,7*2)
' 중국어단어가 나올 수 도 있다.
' 200원에 3개의 생명이 가능 하다. 마이너스 포인트로 게임을 할 수 있다. 단, 가입시 신용지수가 10000원으로 간주한다.
' 이름은 트리스인데 이지트리스? 넷트리스? OOO트리스 배틀트리스
' 팀플이 가능해야 한다. 최대 8명 4vs4
' 팀플의 경우는 우리팀이 하단에 있고 상대팀이 상단에 있으며 팀에게 서로 메시지를 주고 받을 수 있다.
' 팀플의 경우는 탁구처럼 한명씩 돌아가면서 게임을 진행한다.
' 트레이너는 10만원에 개인별로 판매하도록 하는데 pc마다 고유하게 설치하며 서버로 이용이 가능하고 또한
' 여러명이 사용 할 수 또한 있다.
' 즉 판매후 사후관리를 하지 않아 더 좋은 장점이 있으며 온라인 베틀의 장은 베틀 서버를 마련해서 배틀순위에 따라 상을 주도록
' 하며 배틀은 하루 1게임만 가능하며 최대팀 4명에게 상을 주도록 한다.
' 상은 현금과 유사한 도서 상품권을 주도록 한다.
' 한팀에게만 주는것에서 점차 늘려가능 방식으로 하면 좋을 듯함.

Load frmGame01
SetParent frmGame01.hwnd, Me.hwnd
Set frmGame01.parent = Me
frmGame01.Visible = True
SizeControls imgSplitter.Left

End Sub


Private Sub mnuMakePaper_Click()
'If Not IDEMODE Then
'    mnuSchedule_Click
'Else
    Load frmMakePaper
    Set frmMakePaper.parent = Me
    
    frmMakePaper.Show
    
'End If
End Sub

Public Sub mnuMakeSub_Click(Index As Integer)
Dim vData As Variant
Dim Pos As Integer
Dim MNU As Menu

Dim selPocket As String
Dim selChasu As String


'vData = tvTreeView.SelectedItem.KEY
vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '오류발생
Pos = InStrRev(vData, " ")

If InStr(tvTreeView.SelectedItem.key, "_") Then
    selPocket = Left(vData, Pos - 1)
    selChasu = Trim(Mid(vData, Pos))
    
    If Con.State = 1 Then
        If insert_pocketinfo5(selPocket, selChasu, Index) Then
        
            Select Case Index
            Case 0
                Call decress_tp03(tvTreeView.SelectedItem.key, 1, 0)
            Case 1
                Call decress_tp03(tvTreeView.SelectedItem.key, 0, 1)
            Case 4, 5
                Call decress_tp03(tvTreeView.SelectedItem.key, 0, 1)
                Call decress_tp03(tvTreeView.SelectedItem.key, 1, 0)
            End Select
            
            mnuRefresh_Click
            
        End If
    Else
        Con_Open
        If insert_pocketinfo5(selPocket, selChasu, Index) Then
            
            Select Case Index
            Case 0
                Call decress_tp03(tvTreeView.SelectedItem.key, 1, 0)
            Case 1
                Call decress_tp03(tvTreeView.SelectedItem.key, 0, 1)
            Case 4, 5
                Call decress_tp03(tvTreeView.SelectedItem.key, 0, 1)
                Call decress_tp03(tvTreeView.SelectedItem.key, 1, 0)
            End Select
            
            mnuRefresh_Click
        End If
        
        Con_Close
    End If
    
    
End If

End Sub

Private Sub decress_tp03(ByVal key As String, xm As Integer, chkm As Integer)
Dim lkey As String
Dim lsql As String
Dim rs1 As ADODB.Recordset
Dim affected As Long



lkey = Left(key, Len(key) - 1)

lsql = "select pocketnm,chasu from tp02 where code=" & lkey & " and userid='" & gUserid & "'"
Set rs1 = Fn_SQLExec(lsql).rs


If xm > 0 Then
    lsql = "update tp03 set x=x-" & xm & " where userid='" & gUserid & "' "
    lsql = lsql & "and x>=" & xm & " and pocketnm='" & rs1(0) & "' and chasu = " & rs1(1) & " "
Else
    lsql = "update tp03 set chk=chk-" & chkm & " where userid='" & gUserid & "' "
    lsql = lsql & "and chk>=" & chkm & " and pocketnm='" & rs1(0) & "' and chasu = " & rs1(1) & " "
End If

affected = Fn_SQLExec(lsql).nrow
Debug.Assert affected >= 0

End Sub



Private Sub mnuMonth_Click()
Load frmMonth
SetParent frmMonth.hwnd, Me.hwnd
Set frmMonth.parent = Me
frmMonth.Visible = True
SizeControls imgSplitter.Left
End Sub

Private Sub mnuProperty_Click()
Dim cap As String
Dim selPocket As String
Dim selChasu As Long
Dim rs0 As Integer
Dim rs1 As Integer
Dim RS2 As Integer
Dim rs3 As Integer
Dim rs4 As Integer
Dim rs_pass As Integer
Dim rs_solve As Integer
Dim rs_review_far As Integer

Dim selOrder As Integer
Dim isErr As Boolean
Dim errCnt As Integer

Dim rsK As Long
Dim vData As Variant
Dim Pos As Long
Dim RSTOTAL As Variant
Dim Jindo As Variant
Dim CorrectRate As Variant

'vData = tvTreeView.SelectedItem.KEY
vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '실행중인작업있는경우

Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))

cap = ""
Con_Open

sSql = "select count(b.seq) from tP03 a inner join tu02 b on (a.userid=b.userid and a.subj=b.subj and a.seq=b.seq) where a.POCKETNM='" & selPocket & "' and a.userid='" & gUserid & "' and a.chasu=" & selChasu & " and b.reserve_ymd>='" & date2Str(Now()) & "'"
rs_review_far = Fn_SQLExec(sSql).rs(0)

sSql = "select count(b.seq) from tP03 a inner join tu02 b on (a.userid=b.userid and a.subj=b.subj and a.seq=b.seq) where a.POCKETNM='" & selPocket & "' and a.userid='" & gUserid & "' and a.chasu=" & selChasu & " and b.o>b.x"
rs_pass = Fn_SQLExec(sSql).rs(0)

sSql = "select count(b.seq) from tP03 a inner join tu02 b on (a.userid=b.userid and a.subj=b.subj and a.seq=b.seq) where a.POCKETNM='" & selPocket & "' and a.userid='" & gUserid & "' and a.chasu=" & selChasu & " and (b.o>0 or b.x>0)"
rs_solve = Fn_SQLExec(sSql).rs(0)

sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and  o=x AND O=0 and chasu=" & selChasu
rs0 = Fn_SQLExec(sSql).rs(0)

sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and   o>x and chasu=" & selChasu
rs1 = Fn_SQLExec(sSql).rs(0)

sSql = "select count(*) from tP03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' and   o<x and chasu=" & selChasu
RS2 = Fn_SQLExec(sSql).rs(0)

sSql = "SELECT COUNT(*) FROM TP03 WHERE POCKETNM='" & selPocket & "' AND CHASU=" & selChasu & " AND USERID='" & gUserid & "' "
RSTOTAL = Fn_SQLExec(sSql).rs(0)

sSql = "select od from tu03 where POCKETNM='" & selPocket & "' and userid='" & gUserid & "' AND CHASU=" & selChasu
selOrder = Fn_SQLExec(sSql).rs(0)

rsK = RSTOTAL - rs0 - rs1 - RS2
Dim passRatio As Integer, trustRatio As Integer
passRatio = Fix(rs_pass / RSTOTAL * 100)
trustRatio = Fix(rs_review_far / RSTOTAL * 100)

If RSTOTAL = rs_solve And RSTOTAL > 0 Then '★♣●◑▲♨
    
    If passRatio = 100 Then
        cap = cap & vbCrLf & "★★★★★★★" '별3'퍼팩트
    ElseIf passRatio >= 90 Then
        cap = cap & vbCrLf & "★★★★★★☆" '별2
    ElseIf passRatio >= 80 Then
        cap = cap & vbCrLf & "★★★★★☆☆" '별
    ElseIf passRatio >= 70 Then
        cap = cap & vbCrLf & "★★★★☆☆☆" '다이아2
    ElseIf passRatio >= 60 Then
        cap = cap & vbCrLf & "★★★☆☆☆☆" '다이아 완성
    ElseIf passRatio >= 50 Then
        cap = cap & vbCrLf & "★★☆☆☆☆☆" '반쪽 다이아
    Else
        cap = cap & vbCrLf & "★☆☆☆☆☆☆" '똥
    End If
    
    cap = cap & " 합격율: " & passRatio & "%"

    If trustRatio = 100 Then
        cap = cap & vbCrLf & "★★★★★★★"
    ElseIf trustRatio >= 90 Then
        cap = cap & vbCrLf & "★★★★★★☆"
    ElseIf trustRatio >= 80 Then
        cap = cap & vbCrLf & "★★★★★☆☆"
    ElseIf trustRatio >= 70 Then
        cap = cap & vbCrLf & "★★★★☆☆☆"
    ElseIf trustRatio >= 60 Then
        cap = cap & vbCrLf & "★★★☆☆☆☆"
    ElseIf trustRatio >= 50 Then
        cap = cap & vbCrLf & "★★☆☆☆☆☆"
    Else
        cap = cap & vbCrLf & "★☆☆☆☆☆☆"
    End If
    
    cap = cap & " 합격율 신뢰성:" & trustRatio & "%"
    
    cap = cap & vbCrLf & "총평:" & getPyeong(rs_pass, RSTOTAL)

    If trustRatio < 50 Then
        cap = cap & vbCrLf & "복습이 필요합니다."
    End If
 Else
    cap = cap & vbCrLf & "☞ 평가가 완료되지 않았습니다."
 End If

cap = cap & vbCrLf & "-------------------------------------"
cap = cap & vbCrLf & "시험지명:" & selPocket & "(차수: " & selChasu & ")"
cap = cap & vbCrLf & "-------------------------------------"
cap = cap & vbCrLf & "안푼문제: " & rs0
cap = cap & vbCrLf & "맞은문제: " & rs1
cap = cap & vbCrLf & "틀린문제: " & RS2 & "(+" & rsK & ")"
cap = cap & vbCrLf & "-------------------------------------"
cap = cap & vbCrLf & "전체문제: " & RSTOTAL
If RSTOTAL > 0 Then
Jindo = (rs1 + RS2 + rsK) / RSTOTAL * 100
cap = cap & vbCrLf & "-------------------------------------"

cap = cap & vbCrLf & " 풀이완료 : " & Right("       " & Format(Jindo, "##0.00") & " %", 8)

Select Case Jindo

Case 0 To 5
    cap = cap & "□□□□□□□□□□"
Case 5 To 15
    cap = cap & "■□□□□□□□□□"
Case 15 To 25
    cap = cap & "■■□□□□□□□□"
Case 25 To 35
    cap = cap & "■■■□□□□□□□"
Case 35 To 45
    cap = cap & "■■■■□□□□□□"
Case 45 To 55
    cap = cap & "■■■■■□□□□□"
Case 55 To 65
    cap = cap & "■■■■■■□□□□"
Case 65 To 75
    cap = cap & "■■■■■■■□□□"
Case 75 To 85
    cap = cap & "■■■■■■■■□□"
Case 85 To 95
    cap = cap & "■■■■■■■■■□"
Case 95 To 100
    cap = cap & "■■■■■■■■■■"
End Select

'-------------------------------------------------------------------------------
'문항수 진도 계산(S)
'-------------------------------------------------------------------------------
Dim tmpHangsu As Long
Dim totalHangSu As Long

tmpHangsu = getHangsu(selPocket, selChasu) - 1
If selOrder = 1 Then selOrder = 2
totalHangSu = cMNIQ.geth0(RSTOTAL + 1, 0, selOrder) - 1 '다음 신규 학습문항에서 한개의 항수를 뺀값이 전체항수이다.

Jindo = Format(tmpHangsu / totalHangSu * 100, "##0.00") '문항수진도

cap = cap & vbCrLf & vbCrLf & "-------------------------------------"

cap = cap & vbCrLf & " 진  도  : " & Right("       " & Format(Jindo, "##0.00") & " %", 8)

Select Case Jindo

Case 0 To 5
    cap = cap & "□□□□□□□□□□"
Case 5 To 15
    cap = cap & "■□□□□□□□□□"
Case 15 To 25
    cap = cap & "■■□□□□□□□□"
Case 25 To 35
    cap = cap & "■■■□□□□□□□"
Case 35 To 45
    cap = cap & "■■■■□□□□□□"
Case 45 To 55
    cap = cap & "■■■■■□□□□□"
Case 55 To 65
    cap = cap & "■■■■■■□□□□"
Case 65 To 75
    cap = cap & "■■■■■■■□□□"
Case 75 To 85
    cap = cap & "■■■■■■■■□□"
Case 85 To 95
    cap = cap & "■■■■■■■■■□"
Case 95 To 100
    cap = cap & "■■■■■■■■■■"
Case Is > 100
    cap = cap & "■■■■■■■■■■(+)"
End Select
cap = cap & " " & tmpHangsu & "/" & totalHangSu
'-------------------------------------------------------------------------------
'문항수 진도 계산(E)
'-------------------------------------------------------------------------------

End If
cap = cap & vbCrLf & "-------------------------------------"
'CORRECTRATE = rs1 / (rs1 + RS2 + rsK) * 100
CorrectRate = GETCORRECTRATE(selPocket, selChasu, Con.State)
'If (rs1 + RS2 + rsK) > 0 Then
    cap = cap & vbCrLf & "정 답 율: " & Right("        " & Format(CorrectRate, "##0.00") & " %", 8)

Select Case CorrectRate
Case 0 To 5
    cap = cap & "□□□□□□□□□□"
Case 5 To 15
    cap = cap & "■□□□□□□□□□"
Case 15 To 25
    cap = cap & "■■□□□□□□□□"
Case 25 To 35
    cap = cap & "■■■□□□□□□□"
Case 35 To 45
    cap = cap & "■■■■□□□□□□"
Case 45 To 55
    cap = cap & "■■■■■□□□□□"
Case 55 To 65
    cap = cap & "■■■■■■□□□□"
Case 65 To 75
    cap = cap & "■■■■■■■□□□"
Case 75 To 85
    cap = cap & "■■■■■■■■□□"
Case 85 To 95
    cap = cap & "■■■■■■■■■□"
Case 95 To 100
    cap = cap & "■■■■■■■■■■"
End Select


    cap = cap & vbCrLf & "-------------------------------------"

'End If
cap = cap & vbCrLf
cap = cap & vbCrLf & "문제 복습인터벌 ( > 1) ? "

Dim vOrder As Variant
Do
    vOrder = InputBox(cap, "속성", selOrder)
    If vOrder = "" Then Exit Do
Loop Until IsNumeric(vOrder) And vOrder >= 2

If vOrder <> "" Then
    If vOrder <> selOrder Then
        sSql = "update tu03 set od=" & vOrder & " where userid='" & gUserid & "' and POCKETNM='" & selPocket & "' AND CHASU= " & selChasu
        Fn_SQLExec (sSql)
    End If
        
    If selOrder <> vOrder Then
        If mnuEdit2.Visible And selPocket = gPocket And selChasu = gChasu Then
            frmQuiz.Read_TU03
        End If
    End If
End If
Con_Close

End Sub

Private Sub mnuPt_Click(Index As Integer)
setfontsize (Index)

gQuizOnResize = True
gPgBarSaveValue = frmQuiz.pgBar.Value

frmQuiz.quizDisp False
gQuizOnResize = False
End Sub

Public Sub mnuQuickStart_Click()
Call mnuSolve_Click
Call mnuEXAM_Click

End Sub

Public Sub mnuRefresh_Click()
   On Error GoTo Err1
   Dim nodX As Node
   Dim sKey As String
   Dim sNm As String
   Dim nImage As Integer
   Dim nSelectedImage As Integer
   Dim oNodex As Object
   
   'Dim nodeLast As Node
   tvTreeView.Nodes.Clear
'   Set nodX = tvTreeView.Nodes.Add(, , "R", "시험지")
'   Set nodeLast = nodX
    Con_Open
    sSql = "select pocketnm ,chasu,code,pcode,xm,chkm,image,selectimage from tp02 where userid='" & gUserid & "' and hidden=0 order by pcode, create_ymd,code"
    Set rs = Fn_SQLExec(sSql).rs
    
    Set Module1.TreeViewCache = New Collection 'DB에서 찾지 않고 캐시된 로컬에서 빨리 찾기 위함 써버에서 디질필요는 없잖아? 키는 treeview의 키와 같게 한다.
    
    Set oNodex = tvTreeView.Nodes.Add(, tvwChild, _
               "00_", gUserid, 1, 1)
    Do Until rs.EOF()
    
        sKey = CStr(rs(0)) & " " & rs(1)
        
        sNm = rs(0) & "() " ' & "(" & Fix(GETjindo(rs(0), rs(1), Con.State)) & "/" & Fix(GETCORRECTRATE(rs(0), rs(1), Con.State)) & ")"
        'sNm = rs(0) & "(" & Fix(GETjindo(rs(0), rs(1), Con.State)) & "/" & Fix(GETCORRECTRATE(rs(0), rs(1), Con.State)) & ")"
        
        nImage = rs.Fields("image")
        nSelectedImage = rs.Fields("selectimage")
        If Trim(rs.Fields("pcode")) = 0 Then 'All root nodes have 0_ in the parent field
            Set oNodex = tvTreeView.Nodes.Add("00_", tvwChild, rs("code") & "_", _
              sNm, nImage, nSelectedImage)
              Call TreeViewCache.Add(rs("pocketnm") & " " & rs("chasu"), rs("code") & "_")
              oNodex.EnsureVisible 'expend the TreeView so all nodes are visible
        Else 'All child nodes will have the parent key stored in the parent field
            On Error Resume Next
            Set oNodex = tvTreeView.Nodes.Add(Trim(rs.Fields("pcode")) & "_", tvwChild, _
               rs("code") & "_", sNm, nImage, nSelectedImage)
               
               Call TreeViewCache.Add(rs("pocketnm") & " " & rs("chasu"), rs("code") & "_")
               
            oNodex
            If err.Number = 35601 Then
               '낙동강오리알
                Set oNodex = tvTreeView.Nodes.Add("00_", tvwChild, rs("code") & "_", _
                  sNm, nImage, nSelectedImage)
                  
                  Call TreeViewCache.Add(rs("pocketnm") & " " & rs("chasu"), rs("code") & "_")
                  
            End If
            On Error GoTo 0
            oNodex.EnsureVisible 'expend the TreeView so all nodes are visible
        End If
        
'        Set nodeLast = tvTreeView.Nodes.Add("R", tvwChild, skey, sNm, 1, 2)
        rs.MoveNext
    Loop
'    Con_Close
   
   tvTreeView.BorderStyle = vbFixedSingle
   
   If Module1.vDataGlobal <> "" Then
      Dim bSuccess As Boolean
      bSuccess = TvNodeSelect(Module1.vDataGlobal)
      
      If bSuccess Then
        mnuEXAM_Click
      End If
      
   End If
   
   Exit Sub
Err1:
    MsgBox err.Description, vbExclamation, "포켓퀴즈"

End Sub

Public Function TvNodeSelect(pn_ch As String) As Boolean
    Dim i As Long
    'For i = 2 To tvTreeView.Nodes.count
'        If getPN_CH(tvTreeView.Nodes(i).key) = pn_ch Then
'            tvTreeView.Nodes(i).Selected = True
'            TvNodeSelect = True
'            Exit For
'        End If
    'Next
    
    For i = 2 To tvTreeView.Nodes.count
        If Module1.TreeViewCache(tvTreeView.Nodes(i).key) = pn_ch Then
            tvTreeView.Nodes(i).Selected = True
            TvNodeSelect = True
            Exit For
        End If
    Next
    
End Function

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
'    Select Case Button.Key
'        Case "뒤로"
'            '작업: '뒤로' 단추 코드를 추가하십시오.
'            MsgBox "'뒤로' 단추 코드를 추가하십시오."
'        Case "앞으로"
'            '작업: '앞으로' 단추 코드를 추가하십시오.
'            MsgBox "'앞으로' 단추 코드를 추가하십시오."
'        Case "잘라내기"
'            mnuEditCut_Click
'        Case "복사"
'            mnuEditCopy_Click
'        Case "붙여넣기"
'            mnuEditPaste_Click
'        Case "삭제"
'            mnuFileDelete_Click
'        Case "속성"
'            mnuFileProperties_Click
'        Case "큰 아이콘"
'            tvTreeView.Style = tvwTextOnly ' lvListView.View = lvwIcon
'        Case "작은 아이콘"
'            tvTreeView.Style = tvwTreeLines ''lvListView.View = lvwSmallIcon
'        Case "목록 보기"
'            tvTreeView.Style = tvwTreelinesPictureText
'        Case "자세히 보기"
'            tvTreeView.Style = tvwTreelinesPlusMinusPictureText
'    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "버전 " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    '이 프로젝트에 대한 도움말 파일이 없으면 사용자에게 메시지를 표시합니다.
    '사용자는 [프로젝트 속성] 대화 상자에서 응용 프로그램에 대한
    '도움말 파일을 설정할 수 있습니다.
    If Len(App.HelpFile) = 0 Then
        MsgBox "도움말 목차를 표시할 수 없습니다. 이 프로젝트와 연관된 도움말이 없습니다.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If err Then
            MsgBox err.Description, vbCritical
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    '이 프로젝트에 대한 도움말 파일이 없으면 사용자에게 메시지를 표시합니다.
    '사용자는 [프로젝트 속성] 대화 상자에서 응용 프로그램에 대한
    '도움말 파일을 설정할 수 있습니다.
    If Len(App.HelpFile) = 0 Then
        MsgBox "도움말 목차를 표시할 수 없습니다. 이 프로젝트와 연관된 도움말이 없습니다.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If err Then
            MsgBox err.Description, vbCritical
        End If
    End If

End Sub


Private Sub mnuRename_Click()
tvTreeView.StartLabelEdit
End Sub

Private Function changeNodeName(ByVal NewName As String)

Dim vData As String
Dim Pos As Long

Dim selPocket As String
Dim selChasu As Long

Dim Jindo As Double
Dim CorrectRate As Double
Dim maxChasu As Variant

'
LockWindowUpdate (Me.hwnd)
NewName = Left(NewName, 15)

If InStr(NewName, "(") > 0 Then
    NewName = Left(NewName, InStr(NewName, "(") - 1)
End If

vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Function '실행중인작업있는경우
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))

'##1
Dim conbef As Integer
If Con.State = 1 Then
    conbef = 1
Else
    Con_Open
    conbef = 0
End If
        
        sSql = "select max(chasu) from tp02 where userid='" & gUserid & "'"
        maxChasu = Fn_SQLExec(sSql).rs(0) + 1
        
        Con.BeginTrans
    
        sSql = "update tp02 set pocketnm='" & NewName & "' ,chasu=" & maxChasu & " where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
        Fn_SQLExec (sSql)
        
        sSql = "update tp03 set pocketnm='" & NewName & "' ,chasu=" & maxChasu & " where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
        Fn_SQLExec (sSql)
        
        sSql = "update tu03 set pocketnm='" & NewName & "' ,chasu=" & maxChasu & " where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
        Fn_SQLExec (sSql)
        
        If SQLCA.Error Then
            Con.RollbackTrans
            MsgBox "이름 변경 실패", vbExclamation
        Else
            Con.CommitTrans
            
            mnuRefresh_Click
            
        End If
        
'##2
If conbef = 1 Then
Else
    Con_Close
End If
 
LockWindowUpdate (0&)
 
End Function


Private Sub mnuReTry_Click()
Dim vData As String
Dim Pos As Long

Dim selPocket As String
Dim selChasu As Long
Dim rval As Integer

Dim Jindo As Double
Dim CorrectRate As Double

vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '실행중인작업있는경우
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))

Jindo = GETjindo(selPocket, selChasu, Con.State)
CorrectRate = GETCORRECTRATE(selPocket, selChasu, Con.State)

'If Jindo + CorrectRate > 199 Then
'    '초기화 할 수 없습니다.
'    MsgBox "재시험으로 초기화 할 수 없습니다.!(진도와 점수합이 199이상입니다.)", vbExclamation
'Else
    
    If selChasu = gChasu And selPocket = gPocket And mnuEdit2.Visible Then
        MsgBox "해당 시험을 종료한 후 초기화가 가능합니다.", vbExclamation
    Else
        rval = vbYes
        If cFTP.Profile.setPoP2 Then
            rval = MsgBox("초기화 하시겠습니까?", vbQuestion + vbYesNo)
        End If
        
        If rval = vbYes Then
            Con_Open
            sSql = "update tu03 set lastnew=1, hangsu=1 where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
            Fn_SQLExec (sSql)
            
            sSql = "update tp03 set o=0, x=0 where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
            Fn_SQLExec (sSql)
'            mnuRefresh_Click
            Con_Close
            
            status.Text = "초기화 성공!"
            
        End If
    End If
'End If

End Sub

Private Sub mnuReviewCreate_Click()

Dim vData As Variant
Dim Pos As Integer
Dim MNU As Menu

Dim selPocket As String
Dim selChasu As String

'vData = tvTreeView.SelectedItem.KEY
vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '오류발생
Pos = InStrRev(vData, " ")

If InStr(tvTreeView.SelectedItem.key, "_") Then
    selPocket = Left(vData, Pos - 1)
    selChasu = Trim(Mid(vData, Pos))
    
    If insert_pocketinfo7(selPocket, selChasu, 0) Then
        mnuRefresh_Click
    Else
        MsgBox "시험지를 만들지 않았습니다.", vbExclamation + vbOKOnly, "알림"
    End If

End If


End Sub

Private Sub mnuSchedule_Click()
If gMutex Then
    MsgBox "현재 진행중인 작업이 완료된 후 다시 시도해 주세요.", vbExclamation, "알림"
    Exit Sub
End If
Load frmSchedule
SetParent frmSchedule.hwnd, Me.hwnd
Set frmSchedule.parent = Me
frmSchedule.Visible = True
SizeControls imgSplitter.Left
End Sub

Private Sub mnuSearch_Click()
    
'아래 기능은 사용에 제한을 두기 위한 설정으로 가능하다. 가령 정식사용자가 아니라든지.
'    If gUserid <> "영어" And mnuEdit2.Visible Then
'        MsgBox "기존 시험을 종료 후 실행하세요!.", vbExclamation
'        gbMnuExamClick = False
'        Exit Sub
'    End If
    
    frmSearch.Show
    
End Sub

Private Sub mnuSolve_Click()

Dim vData As String
Dim Pos As Long

Dim selPocket As String
Dim selChasu As Long
Dim rval As Integer

Dim Jindo As Double
Dim CorrectRate As Double
Dim totalcnt As Long


vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '실행중인작업있는경우
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))


Jindo = GETjindo(selPocket, selChasu, Con.State)
CorrectRate = GETCORRECTRATE(selPocket, selChasu, Con.State)
totalcnt = GETTotalCnt(selPocket, selChasu, Con.State)

If Jindo + CorrectRate <> 0 Then
    '초기화 할 수 없습니다.
    MsgBox "학습과정을 생략 할 수 없습니다.!(진도가 0이 아닙니다.)", vbExclamation
Else
    
    If selChasu = gChasu And selPocket = gPocket And mnuEdit2.Visible Then
        MsgBox "해당 시험을 종료한 후 가능합니다.", vbExclamation
    Else
        rval = vbYes
        If cFTP.Profile.setPoP2 Then
            rval = MsgBox("학습과정을 생략하시겠습니까?", vbQuestion + vbYesNo)
        End If
        
        If rval = vbYes Then
            Con_Open
            sSql = "update tu03 set lastnew=" & totalcnt & ", hangsu=" & totalcnt + 1 & ", od=" & totalcnt & " where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
            Fn_SQLExec (sSql)
            Con_Close
            status.Text = "학습과정 생략! 완료"
            
        End If
    End If
End If



End Sub

Private Sub mnuUserUpdate_Click()
   '///



Load frmUserUpdate
SetParent frmUserUpdate.hwnd, Me.hwnd

frmUserUpdate.Visible = True
   
SizeControls imgSplitter.Left
   
End Sub

Private Sub mnuViewWebBrowser_Click()
    '작업: 'mnuViewWebBrowser_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewWebBrowser_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewOptions_Click()
    '작업: 'mnuViewOptions_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewOptions_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewRefresh_Click()
    '작업: 'mnuViewRefresh_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewRefresh_Click' 코드를 추가하십시오."
End Sub



Private Sub mnuVAIByDate_Click()
    '작업: 'mnuVAIByDate_Click' 코드를 추가합니다.
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    '작업: 'mnuVAIByName_Click' 코드를 추가합니다.
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    '작업: 'mnuVAIBySize_Click' 코드를 추가합니다.
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    '작업: 'mnuVAIByType_Click' 코드를 추가합니다.
'  lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    'tbToolBar.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuEditInvertSelection_Click()
    '작업: 'mnuEditInvertSelection_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditInvertSelection_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditSelectAll_Click()
    '작업: 'mnuEditSelectAll_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditSelectAll_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditPasteSpecial_Click()
    '작업: 'mnuEditPasteSpecial_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditPasteSpecial_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditPaste_Click()
    '작업: 'mnuEditPaste_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditPaste_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditCopy_Click()
    '작업: 'mnuEditCopy_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditCopy_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditCut_Click()
    '작업: 'mnuEditCut_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditCut_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuEditUndo_Click()
    '작업: 'mnuEditUndo_Click' 코드를 추가하십시오.
    MsgBox "'mnuEditUndo_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileClose_Click()
    '폼을 언로드합니다.
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    '작업: 'mnuFileProperties_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileProperties_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileRename_Click()
    '작업: 'mnuFileRename_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileRename_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileDelete_Click()
    '작업: 'mnuFileDelete_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileDelete_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileNew_Click()
    '작업: 'mnuFileNew_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileNew_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileSendTo_Click()
    '작업: 'mnuFileSendTo_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSendTo_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileFind_Click()
    '작업: 'mnuFileFind_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileFind_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    'With dlgCommonDialog
        dlgCommonDialog.DialogTitle = "열기"
        dlgCommonDialog.CancelError = False
        '작업: Common Dialog 컨트롤의 플래그와 특성을 설정합니다.
        dlgCommonDialog.Filter = "모든 파일(*.*)|*.*"
        dlgCommonDialog.ShowOpen
        If Len(dlgCommonDialog.FileName) = 0 Then
            Exit Sub
        End If
        sFile = dlgCommonDialog.FileName
    'End With
    '작업: 코드를 추가하여 열려 있는 파일을 처리합니다.

End Sub

#If MYAGENTUSE_ON Then
Private Sub MyAgent_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
On Error Resume Next
If MyAgent.Tag <> "" Then
    MyAgent.Tag = ""
    mnuCharUse.Checked = False
    Exit Sub
Else
    mnuCharUse_Click
End If


End Sub
#End If


Private Sub tmrHourGlass_Timer()

Static Sec60 As Integer
Static OldNum As Long
Static SecAll As Long
Static SecPart As Long
Static MinPart As Long
Static HourPart As Long
Static DayPart As Long
Static reSecAll As Long


Dim Timer1 As Double
If gScreenHourGLS Then
    Screen.MousePointer = vbHourglass
Else
    Screen.MousePointer = vbNormal
    If gMutex = False Then
        Timer1 = Timer - gTimer
        If InStr(Command, "/mdb") = 0 Then
            If Timer1 > 20 Or Timer1 < 0 Then
                Fn_SQLExec ("select 1")
                
            End If
        End If
    End If
End If
If False Then
    tmrHourGlass.Enabled = False
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
    
    If reSecAll > 3600 Then
    
        Set frmMsgBox = New frmMsgBox60sec
        Load frmMsgBox
        
        reSecAll = 0
        frmMsgBox.setMsg "프로그램이 너무 오랫동안 구동되어 종료됩니다."
        frmMsgBox.cnt = 10
        
        TmrAfterTTS_exit = True
        frmMsgBox.Show vbModal, frmMain
        TmrAfterTTS_exit = False
        
        If frmMsgBox.btn_flag = 0 Then
            Unload Me
            End
        Else
            reSecAll = 3600 / 2 '다음번에는 30분으로 설정함.
            Set frmMsgBox = Nothing
        End If

    End If

OldNum = cQuiz.num


End Sub

Private Sub tvTreeView_AfterLabelEdit(Cancel As Integer, NewString As String)

If tvTreeView.Tag = gUserid Then
   Cancel = 1
   Exit Sub
End If
If tvTreeView.Tag <> NewString Then
    If Len(Trim(NewString)) > 0 Then
        Call changeNodeName(NewString)
    End If
End If
End Sub

Private Sub tvTreeView_BeforeLabelEdit(Cancel As Integer)
If tvTreeView.SelectedItem.Text = gUserid Then Cancel = 1
tvTreeView.Tag = tvTreeView.SelectedItem.Text
End Sub

Private Sub tvTreeView_DblClick()
On Error Resume Next
Dim str As String
str = tvTreeView.SelectedItem.key
If InStr(str, "_") > 0 Then
    mnuEXAM_Click
End If

End Sub

Private Sub tvTreeView_GotFocus()
If frmQuiz.Visible Then
    frmQuiz.TmrAfterTTS_focus.Enabled = False
End If
End Sub

Private Sub tvTreeView_KeyDown(keycode As Integer, Shift As Integer)

Dim str As String
If keycode = 13 Or keycode = vbKeyDelete Then
    str = tvTreeView.SelectedItem.key
    If str <> "00_" Then '00_는 최상위 키임
        If keycode = 13 Then
            Call tvTreeView_DblClick
        ElseIf keycode = vbKeyDelete Then
            If MsgBox("[" + tvTreeView.SelectedItem.FullPath + "]시험지를" + vbNewLine + "삭제하시겠습니까?", vbQuestion + vbYesNo, "질문") = vbYes Then
                Call mnuDelete_Click
            End If
        End If
    End If
End If
End Sub


Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim str As String
str = tvTreeView.SelectedItem.key
If str = "00_" And Button = vbRightButton Then
   Me.PopupMenu mnuPocketQuizP
ElseIf InStr(str, "_") > 0 And Button = vbRightButton Then
    
    Me.PopupMenu mnuTv
'    If Me.mnuEdit2.Visible Then
'        If Me.mnuPt(0).Visible Then
'            Dim MNU As Menu
'            Debug.Assert False
'            For Each MNU In mnuPt
'                MNU.Enabled = False
'
'            Next
'        End If
'    End If
End If
End Sub
Public Sub setfontsize_menu_bysize(font_size As Integer)
Select Case font_size
Case Is <= 10
    setfontsize 0

Case Is <= 14
    setfontsize 1

Case Else
    setfontsize Int((font_size - 16) / 2) ' + ((font_size - 16) Mod 2)

End Select

End Sub
Public Sub setfontsize(Index As Integer)
Dim MNU As Menu

For Each MNU In mnuPt
    MNU.Checked = False
Next

If Index = 0 Then
    cFTP.Profile.FontSize = 10
ElseIf Index = 1 Then
    cFTP.Profile.FontSize = 14
ElseIf Index >= 2 Then
    cFTP.Profile.FontSize = (Index + 8) * 2
End If

If Index < mnuPt.count Then
    If Index >= 0 Then 'ddd -1로 되어 오류 나서. 조치함.
        mnuPt(Index).Checked = True
    End If
Else
    mnuPt(mnuPt.count - 1).Checked = True
End If

frmQuiz.picMain.FontSize = cFTP.Profile.FontSize
frmQuiz.pic1.FontSize = cFTP.Profile.FontSize
frmQuiz.optA.FontSize = cFTP.Profile.FontSize
frmQuiz.optB.FontSize = cFTP.Profile.FontSize
frmQuiz.optC.FontSize = cFTP.Profile.FontSize
frmQuiz.optD.FontSize = cFTP.Profile.FontSize
frmQuiz.optE.FontSize = cFTP.Profile.FontSize
 
'frmQuiz.optA.Width = frmQuiz.picMain.TextWidth("A       ")
frmQuiz.optA.Width = frmQuiz.picMain.TextWidth("A       ") - frmQuiz.optTTSA.Width
frmQuiz.optB.Width = frmQuiz.optA.Width 'frmQuiz.PicMain.TextWidth("A       ")
frmQuiz.optC.Width = frmQuiz.optA.Width 'frmQuiz.PicMain.TextWidth("A       ")
frmQuiz.optD.Width = frmQuiz.optA.Width 'frmQuiz.PicMain.TextWidth("A       ")
frmQuiz.optE.Width = frmQuiz.optA.Width 'frmQuiz.PicMain.TextWidth("A       ")

frmQuiz.optA.Height = frmQuiz.picMain.TextHeight("Y")
frmQuiz.optB.Height = frmQuiz.optA.Height
frmQuiz.optC.Height = frmQuiz.optA.Height
frmQuiz.optD.Height = frmQuiz.optA.Height
frmQuiz.optE.Height = frmQuiz.optA.Height

frmQuiz.optTTSQuiz.Height = frmQuiz.optA.Height
frmQuiz.optTTSA.Height = frmQuiz.optA.Height
frmQuiz.optTTSB.Height = frmQuiz.optA.Height
frmQuiz.optTTSC.Height = frmQuiz.optA.Height
frmQuiz.optTTSD.Height = frmQuiz.optA.Height
frmQuiz.optTTSE.Height = frmQuiz.optA.Height

frmQuiz.pic1.Left = frmQuiz.picMain.TextWidth("A       ") + frmQuiz.optA.Left * 2
cFTP.Profile.save
End Sub

Private Function isempty_tg01() As Boolean
Con_Open
sSql = "select count(*) from tg01 "
If Fn_SQLExec(sSql).rs(0) = 0 Then
    isempty_tg01 = True
Else
    isempty_tg01 = False
End If
'Con_Close
End Function

Public Function getHangsu(ByRef pPocket As String, ByRef pChasu As Long) As Long
Dim depthAll As Long
Dim depthnow As Long

Con_Open
sSql = "select lastnew,hangsu,od from tu03 where userid='" & gUserid & "' and POCKETNM='" & pPocket & "' AND CHASU=" & pChasu
Set rs = Fn_SQLExec(sSql, , , True).rs
getHangsu = rs(1)
End Function


Public Function makeRandInt(ByRef startNum As Long, ByRef endNum As Long) As Long
    Dim retVal As Single
    Math.Randomize
    retVal = Math.Rnd() * (endNum + 1 - startNum)
    makeRandInt = Int(retVal + startNum)
    
End Function


Public Sub characterPlay(ByRef motion_name As String, Optional ByRef sec_mothon1 As String, Optional ByRef sec_mothon2 As String)
#If MYAGENTUSE_ON Then
'================================
'merlin.acs
'--------------------------------
'Acknowledge
'Alert
'Announce
'Blink
'Confused
'Congratulate
'Congratulate_2
'Decline
'DoMagic1
'DoMagic2
'DontRecognize
'Explain
'GestureDown
'GestureLeft
'GestureRight
'GestureUp
'GetAttention
'GetAttentionContinued
'GetAttentionReturn
'Greet
'Hearing_1
'Hearing_2
'Hearing_3
'Hearing_4
'Hide
'Idle1_1
'Idle1_2
'Idle1_3
'Idle1_4
'Idle2_1
'Idle2_2
'Idle3_1
'Idle3_2
'LookDown
'LookDownBlink
'LookDownReturn
'LookLeft
'LookLeftBlink
'LookLeftReturn
'LookRight
'LookRightBlink
'LookRightReturn
'LookUp
'LookUpBlink
'LookUpReturn
'MoveDown
'MoveLeft
'MoveRight
'MoveUp
'Pleased
'Process
'Processing
'Read
'ReadContinued
'ReadReturn
'Reading
'RestPose
'Sad
'Search
'Searching
'Show
'StartListening
'StopListening
'Suggest
'Surprised
'Think
'Thinking
'Uncertain
'Wave
'Write
'WriteContinued
'WriteReturn
'Writing
'================================

'================================
'SAEKO.ACS
'--------------------------------
'Alert
'CheckingSomething
'Congratulate
'DeepIdle1
'EmptyTrash
'Explain
'GestureDown
'GestureLeft
'GestureRight
'GestureUp
'GetArtsy
'GetAttention
'GetTechy
'GetWizardy
'Goodbye
'Greeting
'Hearing_1
'Hide
'Idle(1)
'Idle(2)
'Idle(3)
'Idle(4)
'Idle(5)
'Idle(6)
'Idle(7)
'Idle(8)
'Idle(9)
'Idle1_1
'LookDown
'LookDownLeft
'LookDownRight
'LookLeft
'LookRight
'LookUp
'LookUpLeft
'LookUpRight
'Print
'Processing
'RestPose
'Save
'Searching
'SendMail
'Show
'Thinking
'Wave
'================================
    
    
    If CharLoaded Then
        If motion_name = ".stop" Then
            Character.Stop
            Exit Sub
        End If
        If LCase(motion_name) = "random" Or motion_name = "" Then
            
            Character.Play AnimationListBox.List(makeRandInt(0, AnimationListBox.ListCount - 1))
        
        Else
            On Error Resume Next
            
            Character.Play motion_name
            
            If err.Number <> 0 Then
               err.Clear
               
               If Not IsMissing(sec_mothon1) Then
                Character.Play sec_mothon1
                If err.Number <> 0 Then
                   err.Clear
                   If Not IsMissing(sec_mothon2) Then
                    Character.Play sec_mothon2
                   End If
                End If
               End If
                
            End If
            
        End If
        
    End If
    #End If
End Sub

#If MYAGENTUSE_ON Then
Public Sub charChuri()
    If Me.Character Is Nothing Then
        Call mnuCharUse_Click
    End If
End Sub
#End If

Private Function getPyeong(i_pass As Integer, i_total As Variant) As String
Dim R As Double
R = i_pass / i_total * 100

If (R >= 90) Then
    getPyeong = " [축하합니다!.시험지를 삭제해도 괜찮습니다.]"
ElseIf R >= 80 Then
    getPyeong = " [정말 잘했군요!]"
ElseIf R >= 70 Then
    getPyeong = " [참 잘했습니다.]"
ElseIf R >= 60 Then
    getPyeong = " [잘한 편입니다.]"
ElseIf R >= 50 Then
    getPyeong = " [잘한편이지만 조금 더 노력이 필요합니다. 꾸준한 복습이 중요한거 아시죠?]"
ElseIf R >= 40 Then
    getPyeong = " [꾸준한 복습으로 조금 더 노력해 보시길 바랍니다.]"
ElseIf R >= 30 Then
    getPyeong = " [부족한 문항 시험보기를 하시길 권장합니다.]"
ElseIf R >= 20 Then
    getPyeong = " [실력이 부족한 편입니다. 노력이 필요한것 같군요.]"
ElseIf R >= 10 Then
    getPyeong = " [노력이 많이 필요합니다.(좌절금지)]"
Else
    getPyeong = " [포기하지 마세요. 잘 할 수 있을 테니까요.]"
End If

getPyeong = getPyeong & vbNewLine & "레벨: " & Fix(R) & "% (" & i_pass & "/" & i_total & ")"

End Function



Private Sub wb2_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'status.Text = "Navigate..."
End Sub

#If MYAGENTUSE_ON Then
Private Sub wb2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'status.Text = "완료"

If Len(wb2.Tag) > 0 Then
'    frmMain.charChuri

    If Not frmMain.Character Is Nothing Then
        frmMain.Character.Speak wb2.Tag
    End If

End If

End Sub
#End If


#If MYAGENTUSE_ON Then
Private Sub wb3_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If Len(wb3.Tag) > 0 Then
        If Not Me.Character Is Nothing Then
            Me.Character.Speak wb3.Tag
        End If
    End If
End Sub
#End If
