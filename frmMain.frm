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
            Text            =   "����"
            TextSave        =   "����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "���� 10:12"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1305
            TextSave        =   "���� 10:12"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "����ð�"
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
      Caption         =   "�������(&C)"
      Begin VB.Menu mnuRefresh 
         Caption         =   "���ΰ�ħ"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSchedule 
         Caption         =   "�н���ȹ(������ �����)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuDeletePaperAuto 
         Caption         =   "������ �ڵ� ����"
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuMakePaper 
         Caption         =   "�����������"
         Shortcut        =   ^{F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuAccount 
         Caption         =   "��ī��Ʈ ����"
         Shortcut        =   {F7}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUserUpdate 
         Caption         =   "ȸ����������"
      End
      Begin VB.Menu mnuEnv 
         Caption         =   "ȯ�漳��(&R)"
      End
      Begin VB.Menu mnuBar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApproach 
         Caption         =   "�н���������"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuChgUsr 
         Caption         =   "����ں���"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "����(&X)"
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "������(&G)"
      Begin VB.Menu mnuCharUse 
         Caption         =   "ĳ���ͻ��"
      End
      Begin VB.Menu mnuEndSetting 
         Caption         =   "����ð�����"
      End
      Begin VB.Menu mnuGame1 
         Caption         =   "����(��Ʈ����)"
      End
      Begin VB.Menu mnuSepa9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "�˻�"
      End
   End
   Begin VB.Menu mnuTv 
      Caption         =   "Ʈ����"
      Begin VB.Menu mnuEXAM 
         Caption         =   "���躸��(&S)"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMake 
         Caption         =   "�����������"
         Begin VB.Menu mnuReviewCreate 
            Caption         =   "������������"
         End
         Begin VB.Menu mnuBar200501301 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "Ʋ������"
            Index           =   0
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "üũ����"
            Index           =   1
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "��Ǭ����"
            Index           =   2
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "��������"
            Index           =   3
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "Ʋ��+üũ"
            Index           =   4
         End
         Begin VB.Menu mnuMakeSub 
            Caption         =   "Ʋ��+üũ+��ǰ"
            Index           =   5
         End
      End
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReTry 
         Caption         =   "�����"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "����(&X)"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "����������"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuQuickStart 
         Caption         =   "�н����� �� �ٷν���"
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "�н�����"
      End
      Begin VB.Menu mnuBar55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsetPocket 
         Caption         =   "Ŭ������ �����߰�"
      End
      Begin VB.Menu mnuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperty 
         Caption         =   "�Ӽ�(&R)"
      End
   End
   Begin VB.Menu mnuEdit2 
      Caption         =   "��Ʈũ��"
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
         Caption         =   "34pt(���ڱ���)"
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
      Caption         =   "����"
      Begin VB.Menu mnu_help 
         Caption         =   "����"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "���α׷�����"
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

'Microsoft Agent Control 2.0(C:\Windows\MSAgent\agentctl.dll) ������Ŵ 2012.12.26 Windows7�̻󿡼� ȣȯ�� �Ῡ������...�ڱԼ�
Option Explicit
Option Base 0

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

'============================
'Agent ����
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


'mnuCharUse_Click '���

lblTitle(0).BackStyle = 0 '����
lblTitle(1).BackStyle = 0 '����

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
    If MsgBox("���α׷��� �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion) = vbNo Then Cancel = 1
    Exit Sub
End If

Screen.MousePointer = vbHourglass
Con_Close


If InStr(LCase(STRCON), "mdb") > 0 Then
compactDB
End If

Screen.MousePointer = vbDefault

Call cTU01.logout(gUserid)  '�α׾ƿ����� �����Ѵ�.

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
    If Not gChangUser Then '����� ���濡 ���� ��ε��ΰ��� ���ؼ��� ���� �ʴ´�.
        Con.Close
    End If

If Not gChangUser Then
'����ڰ� ����Ǵ� ���� �̰��� ���ϸ� �ȵ�.
    End '�������ᰡ �ȵǴ� ��� �ذ�=>��������,ȯ�溯��,��������=>����ƽ�� �ٷ�� ���� �߿���
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
If cFTP.Profile.setPoP2 Then '����� ���� ������ ����
    MsgBox err.Description, vbExclamation
End If

  
End Sub




Private Sub Form_Resize()
    On Error Resume Next
    gMainOnResize = True
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left + 50 '����������� ���� �߰� ����
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
    gMainOnResize = False '���콺�ٿ�� true ������ ¦�̸� ����â resize �̺�Ʈ������ ���� �������� �ִ�.
End Sub


Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub


Sub SizeControls(x As Single)
    On Error Resume Next
    

    '�ʺ� �����մϴ�.
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

    
    '������ �����մϴ�.
  

'    If tbToolBar.Visible Then
'        tvTreeView.Top = tbToolBar.Height + picTitles.Height
'    Else
'        tvTreeView.Top = picTitles.Height
'    End If

  lvListView.Top = tvTreeView.Top
    

    '���̸� �����մϴ�.
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
        MsgBox "(ĳ���� ����).acs������ �����ϴ�.", vbInformation + vbOKOnly, "�˸�"
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

'Character.LanguageID = &H409 '����
Character.LanguageID = &H412 '�ѱ���

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
If vData = "" Then Exit Sub '���������۾��ִ°��
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))

If cFTP.Profile.setPoP1 Then '������ ����
    If MsgBox("[" & selPocket & "] �������� ���� �Ͻðڽ��ϱ�? (���� = " & selChasu & ")", vbQuestion + vbYesNo) = vbYes Then
        If mnuEdit2.Visible = True Then
            If selPocket = gPocket Then
              MsgBox "���� ���躸�� �ִ� ���Դϴ�.", vbExclamation, "�����Ұ�"
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

'1�ܰ��ΰ��߿� ���� �� ������ �ڽı��� �� ����� 1�ܰ������� �θ��� �θ�� 0�� ���̴�
'+----------+-----------+---------+---------------------+-------+--------+
'| selOrder | tmpHangSu | RSTOTAL | pocketnm            | chasu | userid |
'+----------+-----------+---------+---------------------+-------+--------+
'|       11 |        92 |     107 | 1 ȸ�� 20150816 ~   |     1 | �ڱԼ� |
'|       10 |       171 |      81 | 1 ȸ�� 20151230 ~   |     1 | �ڱԼ� |
'|       10 |        97 |      84 | 1 ȸ�� 20160107 ~   |     1 | �ڱԼ� |
'|       12 |        43 |     128 | 2 ȸ��  ~           |     4 | �ڱԼ� |
'|       20 |       400 |     377 | 2 ȸ�� 20150803 ~   |     1 | �ڱԼ� |
'|       10 |       200 |      20 | 2 ȸ�� 20151218 ~   |     2 | �ڱԼ� |
'|       10 |       200 |      80 | 2 ȸ�� 20151231 ~   |     1 | �ڱԼ� |
'|       12 |        12 |     120 | 3 ȸ��  ~           |     4 | �ڱԼ� |
'|       19 |       276 |     360 | 3 ȸ�� 20150806 ~   |     1 | �ڱԼ� |
'|       10 |       131 |      80 | 3 ȸ�� 20160101 ~   |     1 | �ڱԼ� |
'|       15 |       201 |     196 | 6 ȸ�� 20150815 ~   |     1 | �ڱԼ� |
'|       10 |       200 |      90 | 73 ȸ�� 20140924 ~  |     1 | �ڱԼ� |
'|       10 |        94 |      90 | 74 ȸ�� 20140927 ~  |     1 | �ڱԼ� |
'|       10 |        20 |      90 | 80 ȸ�� 20141015 ~  |     1 | �ڱԼ� |
'|      126 |     31752 |     126 | ��������20160107    |     8 | �ڱԼ� |
'|       10 |       200 |      10 | ��������20160114    |     1 | �ڱԼ� |
'|      289 |    167042 |     289 | ��������20160123    |     1 | �ڱԼ� |
'|      283 |    160178 |     283 | ��������20160209    |     2 | �ڱԼ� |
'|       20 |       800 |      20 | ��������20160216    |     1 | �ڱԼ� |
'+----------+-----------+---------+---------------------+-------+--------+
'19 rows in set (0.50 sec)
'19�� �� �� ����̸� �ű⼭ �����ؼ� ������ �����ϸ� �ǰڴ�.
'�߿��Ѱ�(����ǥ�� ó���� �� ������)�� �������� �ʴ´�.

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
    MsgBox "���� �Ϸ� [" & churi_cnt & "]��", vbOKOnly, Me.Caption
Else
    MsgBox "���� �Ϸ� [" & churi_cnt & "]��", vbOKOnly, Me.Caption
    mnuRefresh_Click
End If

End Sub

Private Sub mnuInsetPocket_Click()

Dim vData As Variant
Dim Pos As Integer
Dim selPocket As String
Dim selChasu  As String

vData = getPN_CH(tvTreeView.SelectedItem.key)
If vData = "" Then Exit Sub '���������۾��ִ°��
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))
'
'If selPocket = gPocket Then
'    MsgBox "���� ���躸�� �ִ� ���Դϴ�.", vbExclamation, "�����Ұ�"
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
    '�ߺ������� ������ ���� �ذ�
    ReDim arr(0 To col_arr.count() - 1) As String
    
    For i = 0 To col_arr.count() - 1
        arr(i) = col_arr(i + 1)
    Next
    
End If

Dim max_num As Long

max_num = getTq03_max_num(selPocket, selChasu)

Dim affected As Long

Dim pre_max_num As Long
pre_max_num = max_num '1�ΰ�� �ش� 1�� ���� ���쵵�� �ؾ� �Ѵ�.

If pre_max_num = 1 Then '������ �Ѱ��� 1���� ������ �� �� ���� �����.
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
        Call col_arr.Remove(quiz_tmp) 'Sunday�� 2����ȸ�Ǿ� ���� �����͸� ������ϴ� ��� 5�� ������ �߻��Ǿ� ������������ ���� �� �־� ������ On Error Resume Next�� �ߴ�.
    End If
    
    If err.Number <> 5 Then
        affected = Fn_SQLExec(sSql).nrow
    Else
        churi = churi - 1
    End If
    lRs.MoveNext
Loop

If 0 < col_arr.count Then
    
    Call MsgBox("" & churi & "�� ���� ó��" & vbNewLine & col_arr.count & "���� ó�� ���� ���Ͽ����ϴ�.(����: ���� ������ �� �������� ����)" & vbNewLine & "��ó���� ������ Ŭ�����忡 �־����ϴ�.", vbExclamation, Me.Caption)
    
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
    
    '�Է��� ���� �䱸 �� �� �ִ� ������
    'create table TR01 (seq int(11) NOT NULL auto_increment,quiz_required mediumtext,userid varchar(50),ymd varchar(8),chk int(2),PRIMARY KEY(seq)) ENGINE=MyISAM ;
    '
    
    Clipboard.Clear
    Clipboard.SetText (strMichuri)
    
Else
    Call MsgBox("" & churi & "�� ���� ó���Ǿ����ϴ�.", vbExclamation, Me.Caption)
End If

If 0 < churi Then
    '���������� �ٲ۴�.
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
    
    '���躸�⸦ �ٷ� �����Ѵ�.
    
    If MsgBox("���躸�⸦ �Ͻðڽ��ϱ�? [" & selPocket & "]", vbQuestion + vbYesNo) = vbYes Then
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
    MsgBox "���� ������ ���� �� �����ϼ���!.", vbExclamation
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
    Exit Sub '���������۾��ִ°��
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
   lblTitle(1).Caption = "�����: " & gPocket & "      ������: " & gUserid
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
     frmQuiz.Visible = True '�������̸鸮�������̺�Ʈ�� �ڵ� �߻���.
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
        Exit Sub 'ddd���⼭ �� �� �� ȣ����...
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
'�������
'1. ���ȭ��ǥŰ�� �� �� �������� ������ review��Ʈ�� ��ϵȴ�.
'2. ���߿� �����Ʈ�� �ִ� ���� �� �� �ִ�.
'3. �߱�� �����ϴ�.(�ٱ���)
'4. ��������� �����Ʈ�� ���ðڽ��ϱ�? ��� ������ ���´�.
'5. �����Ʈ�� ���� ���� ���������� �ٷΰ��Ⱑ �ְ� �ȴ�.
'6. Ʋ�������� ī���õǴ� ���
' 6.1 pass�� ���
' 6.2 ���� ��ĭ �׾��� ���
' Ʋ�������� ī���õǴ� �Ͱ� ���ÿ� ���� ��õ�н� �׸����� ��ϵȴ�.
' Ʋ����쳪 ���� ��� �Ϸ� �ѹ���(���� 1ȸ) ī���� �ȴ�.
' Ʈ���̴� ����� ��� �ð����(�ҿ�)-> ������� ��ռҿ�ð��� �� �� �ֵ��� ����ȴ�.
' �Ǹ��� ������(��ü ���,Ȥ�� ��õ� ������ �ܼ��ϸ鼭��)
' �Ǹ��� ��������
' ��Ʋ �����ϰ� �ϱ�
' ���� ������(���� ��� �ϱ� 3�ʰ�)
' ���� �̸����� ������ 5�� �����ϱ�
' �Ҹ������ �Ѵ� ���̷�Ʈ ����� ����ϰ� �Ѵ�.
' ���� �Ҹ���� ����
' ���� ���� ������Ű�� �ϱ�
' ���̷��������� ���� �� �������� �Ѱ��� �ϳ��� ��������.
' ���濡�� �������� �ϴ� ���̷����� �ټ� �ִ�.
' ��Ʋ���� �̱�� ������ �ο��ȴ�.
' ������ ���� ��������.
' ��밡 ���� ��Ʋ�� �¸��� �ȴ�.
' ��Ʈ���� ���� �˷��ش�.
' Ʋ���� �׳� ���δ�.
' ���߸� ������� ���ߴµ� �ҿ�� �ð��� ��ϵǰ� Ű����Ʒ�ȭ��ǥ�� �̿��� ���� space�� �̿��ϴ� ��� �Ҹ� �ȳ��� ���� �Ѵ�.
' �̸� ����ϴ� ���� �ִ�.
' 3���� ������ �ְ� �ȴ�.
' Ʈ���̴� ���� ���Ӹ��� ������.
' Ʈ���̴� ���� �Ͽ� ���ߴ� ��� ���� ���߰Բ� �ϴ� �Ʒ��� ��Ű��
' �����ϸ鼭 �Ҹ������� �ϴ� ����� ��(��)�� ��ư�� �ѹ� ������ ������ �Ѵ�.
' 1���� �ö���� ��쵵 �ִ�.
' �ƹ���Ű ������ ���� ���ۿ� ������ �Ѱ��� ������ �Ұ� �ȴ�.
' ���� ���� �ø��� ����⿡ ������ ������ �ϳ� �ҰԵǸ� �ٽ� ������ ���۵ȴ�.
' �˿� ������ ������ �ϳ� �Ұ� �Ǹ� �ٽ� ������ ���۵ȴ�.
' �������� �ӵ� ������: ������������ �Ÿ��� �ð��� �ι�� ��������.
' ������� ȿ���� ���� �Ѹ��� �̷���׵θ��� ������ �������.
' Ʋ��Ƚ���� ��ϵȴ�
' ���� Ƚ���� ��ϵȴ�.
' �Ѽҿ�ð�(���ߴµ� �ҿ�� �ð��� ����)�� ����Ƚ���� ���� ��� �ҿ�ð�[��]�� ���� �� �ִ�.
' Ʋ�������� �׳� ���� ���� ������ ������ ��� �µ� �Ҹ��� ���´�.
' ��Ʋ�������� ���� �𸣴� ������ ����� ���� ���濡�� ���� �� �ִ�.
' ����ܾ� ���� ��������� ������ �� �� �ִ�. (���� 3+3,7*2)
' �߱���ܾ ���� �� �� �ִ�.
' 200���� 3���� ������ ���� �ϴ�. ���̳ʽ� ����Ʈ�� ������ �� �� �ִ�. ��, ���Խ� �ſ������� 10000������ �����Ѵ�.
' �̸��� Ʈ�����ε� ����Ʈ����? ��Ʈ����? OOOƮ���� ��ƲƮ����
' ������ �����ؾ� �Ѵ�. �ִ� 8�� 4vs4
' ������ ���� �츮���� �ϴܿ� �ְ� ������� ��ܿ� ������ ������ ���� �޽����� �ְ� ���� �� �ִ�.
' ������ ���� Ź��ó�� �Ѹ� ���ư��鼭 ������ �����Ѵ�.
' Ʈ���̳ʴ� 10������ ���κ��� �Ǹ��ϵ��� �ϴµ� pc���� �����ϰ� ��ġ�ϸ� ������ �̿��� �����ϰ� ����
' �������� ��� �� �� ���� �ִ�.
' �� �Ǹ��� ���İ����� ���� �ʾ� �� ���� ������ ������ �¶��� ��Ʋ�� ���� ��Ʋ ������ �����ؼ� ��Ʋ������ ���� ���� �ֵ���
' �ϸ� ��Ʋ�� �Ϸ� 1���Ӹ� �����ϸ� �ִ��� 4���� ���� �ֵ��� �Ѵ�.
' ���� ���ݰ� ������ ���� ��ǰ���� �ֵ��� �Ѵ�.
' �������Ը� �ִ°Ϳ��� ���� �÷����� ������� �ϸ� ���� ����.

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
If vData = "" Then Exit Sub '�����߻�
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
If vData = "" Then Exit Sub '���������۾��ִ°��

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

If RSTOTAL = rs_solve And RSTOTAL > 0 Then '�ڢ��ܢš��
    
    If passRatio = 100 Then
        cap = cap & vbCrLf & "�ڡڡڡڡڡڡ�" '��3'����Ʈ
    ElseIf passRatio >= 90 Then
        cap = cap & vbCrLf & "�ڡڡڡڡڡڡ�" '��2
    ElseIf passRatio >= 80 Then
        cap = cap & vbCrLf & "�ڡڡڡڡڡ١�" '��
    ElseIf passRatio >= 70 Then
        cap = cap & vbCrLf & "�ڡڡڡڡ١١�" '���̾�2
    ElseIf passRatio >= 60 Then
        cap = cap & vbCrLf & "�ڡڡڡ١١١�" '���̾� �ϼ�
    ElseIf passRatio >= 50 Then
        cap = cap & vbCrLf & "�ڡڡ١١١١�" '���� ���̾�
    Else
        cap = cap & vbCrLf & "�ڡ١١١١١�" '��
    End If
    
    cap = cap & " �հ���: " & passRatio & "%"

    If trustRatio = 100 Then
        cap = cap & vbCrLf & "�ڡڡڡڡڡڡ�"
    ElseIf trustRatio >= 90 Then
        cap = cap & vbCrLf & "�ڡڡڡڡڡڡ�"
    ElseIf trustRatio >= 80 Then
        cap = cap & vbCrLf & "�ڡڡڡڡڡ١�"
    ElseIf trustRatio >= 70 Then
        cap = cap & vbCrLf & "�ڡڡڡڡ١١�"
    ElseIf trustRatio >= 60 Then
        cap = cap & vbCrLf & "�ڡڡڡ١١١�"
    ElseIf trustRatio >= 50 Then
        cap = cap & vbCrLf & "�ڡڡ١١١١�"
    Else
        cap = cap & vbCrLf & "�ڡ١١١١١�"
    End If
    
    cap = cap & " �հ��� �ŷڼ�:" & trustRatio & "%"
    
    cap = cap & vbCrLf & "����:" & getPyeong(rs_pass, RSTOTAL)

    If trustRatio < 50 Then
        cap = cap & vbCrLf & "������ �ʿ��մϴ�."
    End If
 Else
    cap = cap & vbCrLf & "�� �򰡰� �Ϸ���� �ʾҽ��ϴ�."
 End If

cap = cap & vbCrLf & "-------------------------------------"
cap = cap & vbCrLf & "��������:" & selPocket & "(����: " & selChasu & ")"
cap = cap & vbCrLf & "-------------------------------------"
cap = cap & vbCrLf & "��Ǭ����: " & rs0
cap = cap & vbCrLf & "��������: " & rs1
cap = cap & vbCrLf & "Ʋ������: " & RS2 & "(+" & rsK & ")"
cap = cap & vbCrLf & "-------------------------------------"
cap = cap & vbCrLf & "��ü����: " & RSTOTAL
If RSTOTAL > 0 Then
Jindo = (rs1 + RS2 + rsK) / RSTOTAL * 100
cap = cap & vbCrLf & "-------------------------------------"

cap = cap & vbCrLf & " Ǯ�̿Ϸ� : " & Right("       " & Format(Jindo, "##0.00") & " %", 8)

Select Case Jindo

Case 0 To 5
    cap = cap & "�����������"
Case 5 To 15
    cap = cap & "�����������"
Case 15 To 25
    cap = cap & "�����������"
Case 25 To 35
    cap = cap & "�����������"
Case 35 To 45
    cap = cap & "�����������"
Case 45 To 55
    cap = cap & "�����������"
Case 55 To 65
    cap = cap & "�����������"
Case 65 To 75
    cap = cap & "�����������"
Case 75 To 85
    cap = cap & "�����������"
Case 85 To 95
    cap = cap & "�����������"
Case 95 To 100
    cap = cap & "�����������"
End Select

'-------------------------------------------------------------------------------
'���׼� ���� ���(S)
'-------------------------------------------------------------------------------
Dim tmpHangsu As Long
Dim totalHangSu As Long

tmpHangsu = getHangsu(selPocket, selChasu) - 1
If selOrder = 1 Then selOrder = 2
totalHangSu = cMNIQ.geth0(RSTOTAL + 1, 0, selOrder) - 1 '���� �ű� �н����׿��� �Ѱ��� �׼��� ������ ��ü�׼��̴�.

Jindo = Format(tmpHangsu / totalHangSu * 100, "##0.00") '���׼�����

cap = cap & vbCrLf & vbCrLf & "-------------------------------------"

cap = cap & vbCrLf & " ��  ��  : " & Right("       " & Format(Jindo, "##0.00") & " %", 8)

Select Case Jindo

Case 0 To 5
    cap = cap & "�����������"
Case 5 To 15
    cap = cap & "�����������"
Case 15 To 25
    cap = cap & "�����������"
Case 25 To 35
    cap = cap & "�����������"
Case 35 To 45
    cap = cap & "�����������"
Case 45 To 55
    cap = cap & "�����������"
Case 55 To 65
    cap = cap & "�����������"
Case 65 To 75
    cap = cap & "�����������"
Case 75 To 85
    cap = cap & "�����������"
Case 85 To 95
    cap = cap & "�����������"
Case 95 To 100
    cap = cap & "�����������"
Case Is > 100
    cap = cap & "�����������(+)"
End Select
cap = cap & " " & tmpHangsu & "/" & totalHangSu
'-------------------------------------------------------------------------------
'���׼� ���� ���(E)
'-------------------------------------------------------------------------------

End If
cap = cap & vbCrLf & "-------------------------------------"
'CORRECTRATE = rs1 / (rs1 + RS2 + rsK) * 100
CorrectRate = GETCORRECTRATE(selPocket, selChasu, Con.State)
'If (rs1 + RS2 + rsK) > 0 Then
    cap = cap & vbCrLf & "�� �� ��: " & Right("        " & Format(CorrectRate, "##0.00") & " %", 8)

Select Case CorrectRate
Case 0 To 5
    cap = cap & "�����������"
Case 5 To 15
    cap = cap & "�����������"
Case 15 To 25
    cap = cap & "�����������"
Case 25 To 35
    cap = cap & "�����������"
Case 35 To 45
    cap = cap & "�����������"
Case 45 To 55
    cap = cap & "�����������"
Case 55 To 65
    cap = cap & "�����������"
Case 65 To 75
    cap = cap & "�����������"
Case 75 To 85
    cap = cap & "�����������"
Case 85 To 95
    cap = cap & "�����������"
Case 95 To 100
    cap = cap & "�����������"
End Select


    cap = cap & vbCrLf & "-------------------------------------"

'End If
cap = cap & vbCrLf
cap = cap & vbCrLf & "���� �������͹� ( > 1) ? "

Dim vOrder As Variant
Do
    vOrder = InputBox(cap, "�Ӽ�", selOrder)
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
'   Set nodX = tvTreeView.Nodes.Add(, , "R", "������")
'   Set nodeLast = nodX
    Con_Open
    sSql = "select pocketnm ,chasu,code,pcode,xm,chkm,image,selectimage from tp02 where userid='" & gUserid & "' and hidden=0 order by pcode, create_ymd,code"
    Set rs = Fn_SQLExec(sSql).rs
    
    Set Module1.TreeViewCache = New Collection 'DB���� ã�� �ʰ� ĳ�õ� ���ÿ��� ���� ã�� ���� ������� �����ʿ�� ���ݾ�? Ű�� treeview�� Ű�� ���� �Ѵ�.
    
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
               '������������
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
    MsgBox err.Description, vbExclamation, "��������"

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
'        Case "�ڷ�"
'            '�۾�: '�ڷ�' ���� �ڵ带 �߰��Ͻʽÿ�.
'            MsgBox "'�ڷ�' ���� �ڵ带 �߰��Ͻʽÿ�."
'        Case "������"
'            '�۾�: '������' ���� �ڵ带 �߰��Ͻʽÿ�.
'            MsgBox "'������' ���� �ڵ带 �߰��Ͻʽÿ�."
'        Case "�߶󳻱�"
'            mnuEditCut_Click
'        Case "����"
'            mnuEditCopy_Click
'        Case "�ٿ��ֱ�"
'            mnuEditPaste_Click
'        Case "����"
'            mnuFileDelete_Click
'        Case "�Ӽ�"
'            mnuFileProperties_Click
'        Case "ū ������"
'            tvTreeView.Style = tvwTextOnly ' lvListView.View = lvwIcon
'        Case "���� ������"
'            tvTreeView.Style = tvwTreeLines ''lvListView.View = lvwSmallIcon
'        Case "��� ����"
'            tvTreeView.Style = tvwTreelinesPictureText
'        Case "�ڼ��� ����"
'            tvTreeView.Style = tvwTreelinesPlusMinusPictureText
'    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "���� " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    '�� ������Ʈ�� ���� ���� ������ ������ ����ڿ��� �޽����� ǥ���մϴ�.
    '����ڴ� [������Ʈ �Ӽ�] ��ȭ ���ڿ��� ���� ���α׷��� ����
    '���� ������ ������ �� �ֽ��ϴ�.
    If Len(App.HelpFile) = 0 Then
        MsgBox "���� ������ ǥ���� �� �����ϴ�. �� ������Ʈ�� ������ ������ �����ϴ�.", vbInformation, Me.Caption
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


    '�� ������Ʈ�� ���� ���� ������ ������ ����ڿ��� �޽����� ǥ���մϴ�.
    '����ڴ� [������Ʈ �Ӽ�] ��ȭ ���ڿ��� ���� ���α׷��� ����
    '���� ������ ������ �� �ֽ��ϴ�.
    If Len(App.HelpFile) = 0 Then
        MsgBox "���� ������ ǥ���� �� �����ϴ�. �� ������Ʈ�� ������ ������ �����ϴ�.", vbInformation, Me.Caption
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
If vData = "" Then Exit Function '���������۾��ִ°��
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
            MsgBox "�̸� ���� ����", vbExclamation
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
If vData = "" Then Exit Sub '���������۾��ִ°��
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))

Jindo = GETjindo(selPocket, selChasu, Con.State)
CorrectRate = GETCORRECTRATE(selPocket, selChasu, Con.State)

'If Jindo + CorrectRate > 199 Then
'    '�ʱ�ȭ �� �� �����ϴ�.
'    MsgBox "��������� �ʱ�ȭ �� �� �����ϴ�.!(������ �������� 199�̻��Դϴ�.)", vbExclamation
'Else
    
    If selChasu = gChasu And selPocket = gPocket And mnuEdit2.Visible Then
        MsgBox "�ش� ������ ������ �� �ʱ�ȭ�� �����մϴ�.", vbExclamation
    Else
        rval = vbYes
        If cFTP.Profile.setPoP2 Then
            rval = MsgBox("�ʱ�ȭ �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo)
        End If
        
        If rval = vbYes Then
            Con_Open
            sSql = "update tu03 set lastnew=1, hangsu=1 where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
            Fn_SQLExec (sSql)
            
            sSql = "update tp03 set o=0, x=0 where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
            Fn_SQLExec (sSql)
'            mnuRefresh_Click
            Con_Close
            
            status.Text = "�ʱ�ȭ ����!"
            
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
If vData = "" Then Exit Sub '�����߻�
Pos = InStrRev(vData, " ")

If InStr(tvTreeView.SelectedItem.key, "_") Then
    selPocket = Left(vData, Pos - 1)
    selChasu = Trim(Mid(vData, Pos))
    
    If insert_pocketinfo7(selPocket, selChasu, 0) Then
        mnuRefresh_Click
    Else
        MsgBox "�������� ������ �ʾҽ��ϴ�.", vbExclamation + vbOKOnly, "�˸�"
    End If

End If


End Sub

Private Sub mnuSchedule_Click()
If gMutex Then
    MsgBox "���� �������� �۾��� �Ϸ�� �� �ٽ� �õ��� �ּ���.", vbExclamation, "�˸�"
    Exit Sub
End If
Load frmSchedule
SetParent frmSchedule.hwnd, Me.hwnd
Set frmSchedule.parent = Me
frmSchedule.Visible = True
SizeControls imgSplitter.Left
End Sub

Private Sub mnuSearch_Click()
    
'�Ʒ� ����� ��뿡 ������ �α� ���� �������� �����ϴ�. ���� ���Ļ���ڰ� �ƴ϶����.
'    If gUserid <> "����" And mnuEdit2.Visible Then
'        MsgBox "���� ������ ���� �� �����ϼ���!.", vbExclamation
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
If vData = "" Then Exit Sub '���������۾��ִ°��
Pos = InStrRev(vData, " ")

selPocket = Left(vData, Pos - 1)
selChasu = Trim(Mid(vData, Pos))


Jindo = GETjindo(selPocket, selChasu, Con.State)
CorrectRate = GETCORRECTRATE(selPocket, selChasu, Con.State)
totalcnt = GETTotalCnt(selPocket, selChasu, Con.State)

If Jindo + CorrectRate <> 0 Then
    '�ʱ�ȭ �� �� �����ϴ�.
    MsgBox "�н������� ���� �� �� �����ϴ�.!(������ 0�� �ƴմϴ�.)", vbExclamation
Else
    
    If selChasu = gChasu And selPocket = gPocket And mnuEdit2.Visible Then
        MsgBox "�ش� ������ ������ �� �����մϴ�.", vbExclamation
    Else
        rval = vbYes
        If cFTP.Profile.setPoP2 Then
            rval = MsgBox("�н������� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo)
        End If
        
        If rval = vbYes Then
            Con_Open
            sSql = "update tu03 set lastnew=" & totalcnt & ", hangsu=" & totalcnt + 1 & ", od=" & totalcnt & " where pocketnm='" & selPocket & "' and userid='" & gUserid & "' and chasu=" & selChasu
            Fn_SQLExec (sSql)
            Con_Close
            status.Text = "�н����� ����! �Ϸ�"
            
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
    '�۾�: 'mnuViewWebBrowser_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewWebBrowser_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuViewOptions_Click()
    '�۾�: 'mnuViewOptions_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewOptions_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuViewRefresh_Click()
    '�۾�: 'mnuViewRefresh_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewRefresh_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub



Private Sub mnuVAIByDate_Click()
    '�۾�: 'mnuVAIByDate_Click' �ڵ带 �߰��մϴ�.
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    '�۾�: 'mnuVAIByName_Click' �ڵ带 �߰��մϴ�.
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    '�۾�: 'mnuVAIBySize_Click' �ڵ带 �߰��մϴ�.
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    '�۾�: 'mnuVAIByType_Click' �ڵ带 �߰��մϴ�.
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
    '�۾�: 'mnuEditInvertSelection_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditInvertSelection_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditSelectAll_Click()
    '�۾�: 'mnuEditSelectAll_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditSelectAll_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditPasteSpecial_Click()
    '�۾�: 'mnuEditPasteSpecial_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditPasteSpecial_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditPaste_Click()
    '�۾�: 'mnuEditPaste_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditPaste_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditCopy_Click()
    '�۾�: 'mnuEditCopy_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditCopy_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditCut_Click()
    '�۾�: 'mnuEditCut_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditCut_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuEditUndo_Click()
    '�۾�: 'mnuEditUndo_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuEditUndo_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileClose_Click()
    '���� ��ε��մϴ�.
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    '�۾�: 'mnuFileProperties_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileProperties_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileRename_Click()
    '�۾�: 'mnuFileRename_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileRename_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileDelete_Click()
    '�۾�: 'mnuFileDelete_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileDelete_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileNew_Click()
    '�۾�: 'mnuFileNew_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileNew_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileSendTo_Click()
    '�۾�: 'mnuFileSendTo_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSendTo_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileFind_Click()
    '�۾�: 'mnuFileFind_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileFind_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    'With dlgCommonDialog
        dlgCommonDialog.DialogTitle = "����"
        dlgCommonDialog.CancelError = False
        '�۾�: Common Dialog ��Ʈ���� �÷��׿� Ư���� �����մϴ�.
        dlgCommonDialog.Filter = "��� ����(*.*)|*.*"
        dlgCommonDialog.ShowOpen
        If Len(dlgCommonDialog.FileName) = 0 Then
            Exit Sub
        End If
        sFile = dlgCommonDialog.FileName
    'End With
    '�۾�: �ڵ带 �߰��Ͽ� ���� �ִ� ������ ó���մϴ�.

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
    '������ �̺�Ʈ�� �߻����� ����
    SecAll = SecAll + 1
    reSecAll = reSecAll + 1
Else
    '�ʱ�ȭ
    Sec60 = 0
    SecAll = 0
    reSecAll = 0
End If
    
    If reSecAll > 3600 Then
    
        Set frmMsgBox = New frmMsgBox60sec
        Load frmMsgBox
        
        reSecAll = 0
        frmMsgBox.setMsg "���α׷��� �ʹ� �������� �����Ǿ� ����˴ϴ�."
        frmMsgBox.cnt = 10
        
        TmrAfterTTS_exit = True
        frmMsgBox.Show vbModal, frmMain
        TmrAfterTTS_exit = False
        
        If frmMsgBox.btn_flag = 0 Then
            Unload Me
            End
        Else
            reSecAll = 3600 / 2 '���������� 30������ ������.
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
    If str <> "00_" Then '00_�� �ֻ��� Ű��
        If keycode = 13 Then
            Call tvTreeView_DblClick
        ElseIf keycode = vbKeyDelete Then
            If MsgBox("[" + tvTreeView.SelectedItem.FullPath + "]��������" + vbNewLine + "�����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbYes Then
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
    If Index >= 0 Then 'ddd -1�� �Ǿ� ���� ����. ��ġ��.
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
    getPyeong = " [�����մϴ�!.�������� �����ص� �������ϴ�.]"
ElseIf R >= 80 Then
    getPyeong = " [���� ���߱���!]"
ElseIf R >= 70 Then
    getPyeong = " [�� ���߽��ϴ�.]"
ElseIf R >= 60 Then
    getPyeong = " [���� ���Դϴ�.]"
ElseIf R >= 50 Then
    getPyeong = " [������������ ���� �� ����� �ʿ��մϴ�. ������ ������ �߿��Ѱ� �ƽ���?]"
ElseIf R >= 40 Then
    getPyeong = " [������ �������� ���� �� ����� ���ñ� �ٶ��ϴ�.]"
ElseIf R >= 30 Then
    getPyeong = " [������ ���� ���躸�⸦ �Ͻñ� �����մϴ�.]"
ElseIf R >= 20 Then
    getPyeong = " [�Ƿ��� ������ ���Դϴ�. ����� �ʿ��Ѱ� ������.]"
ElseIf R >= 10 Then
    getPyeong = " [����� ���� �ʿ��մϴ�.(��������)]"
Else
    getPyeong = " [�������� ������. �� �� �� ���� �״ϱ��.]"
End If

getPyeong = getPyeong & vbNewLine & "����: " & Fix(R) & "% (" & i_pass & "/" & i_total & ")"

End Function



Private Sub wb2_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'status.Text = "Navigate..."
End Sub

#If MYAGENTUSE_ON Then
Private Sub wb2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'status.Text = "�Ϸ�"

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
