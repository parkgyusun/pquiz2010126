VERSION 5.00
Begin VB.Form frmBegin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  '≈©±‚ ∞Ì¡§ ¥Î»≠ ªÛ¿⁄
   Caption         =   "#"
   ClientHeight    =   3540
   ClientLeft      =   1740
   ClientTop       =   1410
   ClientWidth     =   7545
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "±º∏≤"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "begin.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInstall 
      Cancel          =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4995
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   3030
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'æ¯¿Ω
      Height          =   915
      Left            =   405
      Picture         =   "begin.frx":0442
      ScaleHeight     =   915
      ScaleWidth      =   960
      TabIndex        =   6
      Top             =   720
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   1080
      Left            =   315
      TabIndex        =   7
      Top             =   630
      Width           =   1170
   End
   Begin VB.Frame fraDir 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   135
      TabIndex        =   4
      Top             =   2010
      Width           =   7296
      Begin VB.CommandButton cmdChDir 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4890
         MaskColor       =   &H00000000&
         TabIndex        =   0
         Top             =   195
         Width           =   2310
      End
      Begin VB.Label lblDestDir 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   300
         Width           =   4440
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2610
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   3030
      Width           =   2205
   End
   Begin VB.Line linTopOfExitButtonIfNoDestDir 
      Visible         =   0   'False
      X1              =   2670
      X2              =   4725
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblInstallMsg 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1725
      TabIndex        =   3
      Top             =   915
      Width           =   5565
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBegin 
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   285
      TabIndex        =   2
      Top             =   135
      Width           =   6450
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChDir_Click()
    ShowPathDialog

    If gintRetVal = gintRET_CONT Then
        lblDestDir.Caption = gstrDestDir
        cmdInstall.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub

Private Sub cmdInstall_Click()
    If IsValidDestDir(gstrDestDir) Then
        Unload Me
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    Dim intRes As Integer
    Dim yAdjust As Integer

    SetFormFont Me
    fraDir.Caption = ResolveResString(resFRMDIRECTORY)
    cmdChDir.Caption = ResolveResString(resBTNCHGDIR)
    cmdExit.Caption = ResolveResString(resBTNEXIT)
    lblBegin.Caption = ResolveResString(resLBLBEGIN)
    cmdInstall.Caption = ResolveResString(resSETUP%, gstrPIPE1, "")
    cmdInstall.ToolTipText = ResolveResString(resBTNTOOLTIPBEGIN)
    
    Caption = gstrTitle
    If gfForceUseDefDest Then
        intRes = resSPECNODEST
    Else
        intRes = resSPECDEST
    End If
    lblInstallMsg.Caption = ResolveResString(intRes, gstrPIPE1, gstrAppName)
    lblDestDir.Caption = gstrDestDir

    If gfForceUseDefDest Then
        'We are forced to use the default destination directory, so the user
        '  will not be able to change it.
        fraDir.Visible = False

        'Close in the blank space on the form by moving the Exit button to where this frame
        'currently is, and adjusting the size of the form respectively
        yAdjust = cmdExit.Top - linTopOfExitButtonIfNoDestDir.Y1
        cmdExit.Top = cmdExit.Top - yAdjust
        Height = Height - yAdjust
    End If
    EtchedLine Me, fraDir.Left, cmdExit.Top - cmdExit.Height \ 2, fraDir.Width

    CenterForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HandleFormQueryUnload UnloadMode, Cancel, Me
End Sub

Private Sub Picture1_Click()
    cmdInstall_Click
End Sub
