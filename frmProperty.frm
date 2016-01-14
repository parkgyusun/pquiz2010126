VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmProperty 
   BorderStyle     =   1  '단일 고정
   Caption         =   "속성"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmProperty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "적용 후 닫기(&O)"
      Height          =   375
      Left            =   1890
      TabIndex        =   45
      Top             =   5475
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기(&X)"
      Height          =   375
      Left            =   5565
      TabIndex        =   3
      Top             =   5475
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "복원(&C)"
      Height          =   375
      Left            =   4485
      TabIndex        =   2
      Top             =   5475
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "적용(&A)"
      Height          =   375
      Left            =   3405
      TabIndex        =   1
      Top             =   5475
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   9022
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "기본"
      TabPicture(0)   =   "frmProperty.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "이벤트 알림"
      TabPicture(1)   =   "frmProperty.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "고급"
      TabPicture(2)   =   "frmProperty.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "chkPoP1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkPoP2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkCompactDB"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chkWhenResize"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkShowWhitSpot"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "chkPoP3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "chkTTSuse"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkTTSuse1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkTTSuse2"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "chkTTSuse3"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "chkTTSuse4"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.CheckBox chkTTSuse4 
         Caption         =   "복습시 소리사용"
         Height          =   285
         Left            =   780
         TabIndex        =   50
         Top             =   4680
         Width           =   4770
      End
      Begin VB.CheckBox chkTTSuse3 
         Caption         =   "복습필요 문항 소리사용(살짝권장)"
         Height          =   285
         Left            =   780
         TabIndex        =   49
         Top             =   4320
         Width           =   4770
      End
      Begin VB.CheckBox chkTTSuse2 
         Caption         =   "불합격 문항 소리사용(권장)"
         Height          =   285
         Left            =   780
         TabIndex        =   48
         Top             =   3990
         Width           =   4770
      End
      Begin VB.CheckBox chkTTSuse1 
         Caption         =   "신규문항 소리사용(권장)"
         Height          =   285
         Left            =   780
         TabIndex        =   47
         Top             =   3660
         Width           =   4770
      End
      Begin VB.CheckBox chkTTSuse 
         Caption         =   "문제읽기 소리사용(TTS기능)"
         Height          =   285
         Left            =   540
         TabIndex        =   46
         Top             =   3330
         Width           =   4770
      End
      Begin VB.CheckBox chkPoP3 
         Caption         =   "닫을 때 질문"
         Height          =   285
         Left            =   540
         TabIndex        =   41
         Top             =   1560
         Width           =   3030
      End
      Begin VB.CheckBox chkShowWhitSpot 
         Caption         =   "창크기 변경시 답 힌트(화면 주위 하얀점)"
         Height          =   285
         Left            =   540
         TabIndex        =   40
         Top             =   2490
         Width           =   4470
      End
      Begin VB.CheckBox chkWhenResize 
         Caption         =   "창크기 변경시 체크문제로 설정(위와 같은 이유로)"
         Height          =   285
         Left            =   540
         TabIndex        =   39
         Top             =   2910
         Width           =   4470
      End
      Begin VB.CheckBox chkCompactDB 
         Caption         =   "종료시 DB 압축"
         Height          =   285
         Left            =   540
         TabIndex        =   38
         Top             =   2010
         Width           =   4470
      End
      Begin VB.CheckBox chkPoP2 
         Caption         =   "경미한 에러시 팝업알림"
         Height          =   285
         Left            =   540
         TabIndex        =   37
         Top             =   1095
         Width           =   4470
      End
      Begin VB.CheckBox chkPoP1 
         Caption         =   "삭제시 질문"
         Height          =   285
         Left            =   540
         TabIndex        =   36
         Top             =   675
         Width           =   4470
      End
      Begin VB.Frame Frame2 
         Height          =   3330
         Left            =   -74775
         TabIndex        =   17
         Top             =   495
         Width           =   5955
         Begin VB.TextBox txtAlm1 
            Height          =   330
            Index           =   2
            Left            =   1080
            TabIndex        =   25
            Top             =   2790
            Width           =   780
         End
         Begin VB.CheckBox chkAlm1 
            Caption         =   "체크된 문제(문제풀이중 H키를 눌러 힌트를 본 문제들)"
            Height          =   285
            Index           =   2
            Left            =   315
            TabIndex        =   24
            Top             =   2295
            Width           =   5010
         End
         Begin VB.TextBox txtAlm1 
            Height          =   330
            Index           =   1
            Left            =   1080
            TabIndex        =   22
            Top             =   1800
            Width           =   780
         End
         Begin VB.CheckBox chkAlm1 
            Caption         =   "잘 모르는 문제 알림"
            Height          =   285
            Index           =   1
            Left            =   315
            TabIndex        =   21
            Top             =   1305
            Width           =   2760
         End
         Begin VB.TextBox txtAlm1 
            Height          =   330
            Index           =   0
            Left            =   1080
            TabIndex        =   19
            Top             =   855
            Width           =   780
         End
         Begin VB.CheckBox chkAlm1 
            Caption         =   "모르는 문제 알림"
            Height          =   285
            Index           =   0
            Left            =   315
            TabIndex        =   18
            Top             =   360
            Width           =   2760
         End
         Begin VB.Label Label5 
            Caption         =   "개마다 "
            Height          =   330
            Index           =   2
            Left            =   1935
            TabIndex        =   26
            Top             =   2880
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "개마다 "
            Height          =   330
            Index           =   1
            Left            =   1935
            TabIndex        =   23
            Top             =   1890
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "개마다 "
            Height          =   330
            Index           =   0
            Left            =   1935
            TabIndex        =   20
            Top             =   945
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3510
         Left            =   -74685
         TabIndex        =   4
         Top             =   450
         Width           =   5850
         Begin VB.ComboBox cboFont 
            Height          =   300
            ItemData        =   "frmProperty.frx":0496
            Left            =   3720
            List            =   "frmProperty.frx":0498
            TabIndex        =   44
            Text            =   "Combo1"
            Top             =   570
            Width           =   1695
         End
         Begin VB.TextBox txtNansu 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   3510
            MaxLength       =   1
            TabIndex        =   43
            Top             =   2835
            Width           =   645
         End
         Begin VB.TextBox txtDelay2 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   3510
            MaxLength       =   1
            TabIndex        =   33
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox txtDelay1 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   285
            Left            =   3510
            MaxLength       =   4
            TabIndex        =   30
            Top             =   1260
            Width           =   600
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   2895
            Left            =   2880
            TabIndex        =   28
            Top             =   360
            Width           =   15
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "폰트 Bold"
            Height          =   240
            Left            =   3105
            TabIndex        =   27
            Top             =   315
            Width           =   1185
         End
         Begin VB.TextBox txtTimeOut 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   285
            Left            =   1215
            TabIndex        =   15
            Top             =   1260
            Width           =   600
         End
         Begin VB.CheckBox chkTimeOut 
            Caption         =   "문제풀때 시간 흐름"
            Height          =   195
            Left            =   225
            TabIndex        =   14
            Top             =   945
            Width           =   2175
         End
         Begin VB.TextBox txtStreatX 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1215
            TabIndex        =   13
            Text            =   "1"
            Top             =   2970
            Width           =   600
         End
         Begin VB.CheckBox chkStreatX 
            BackColor       =   &H8000000A&
            Caption         =   "연속틀림에서 자동 멈춤"
            Enabled         =   0   'False
            Height          =   330
            Left            =   225
            TabIndex        =   12
            Top             =   2610
            Value           =   1  '확인
            Width           =   2445
         End
         Begin VB.TextBox txtSetTimeOutS 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   285
            Left            =   1215
            TabIndex        =   9
            Top             =   2160
            Width           =   600
         End
         Begin VB.CheckBox chkSetTimeOutS 
            Caption         =   "학습시 시간 흐름"
            Height          =   330
            Left            =   225
            TabIndex        =   8
            Top             =   1800
            Width           =   1770
         End
         Begin VB.TextBox txtFontSize 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   285
            Left            =   1215
            TabIndex        =   6
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label9 
            Caption         =   "◎ 폰트"
            Height          =   195
            Left            =   3105
            TabIndex        =   42
            Top             =   630
            Width           =   1320
         End
         Begin VB.Label Label8 
            Caption         =   "◎ 난수 발생정도"
            Height          =   240
            Left            =   3105
            TabIndex        =   35
            Top             =   2520
            Width           =   1950
         End
         Begin VB.Label Label6 
            Caption         =   "Sec"
            Height          =   240
            Index           =   1
            Left            =   4275
            TabIndex        =   34
            Top             =   2115
            Width           =   555
         End
         Begin VB.Label Label7 
            Caption         =   "◎ 새로운 문제알림 "
            Height          =   285
            Left            =   3105
            TabIndex        =   32
            Top             =   1710
            Width           =   1950
         End
         Begin VB.Label Label6 
            Caption         =   "mSec"
            Height          =   240
            Index           =   0
            Left            =   4185
            TabIndex        =   31
            Top             =   1305
            Width           =   555
         End
         Begin VB.Label lblLockTime 
            Caption         =   "◎ 다음진행키 락시간"
            Height          =   285
            Left            =   3105
            TabIndex        =   29
            Top             =   945
            Width           =   1950
         End
         Begin VB.Label Label3 
            Caption         =   "초"
            Height          =   240
            Index           =   1
            Left            =   1890
            TabIndex        =   16
            Top             =   1305
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "회"
            Height          =   240
            Left            =   1890
            TabIndex        =   11
            Top             =   3060
            Width           =   330
         End
         Begin VB.Label Label3 
            Caption         =   "초"
            Height          =   240
            Index           =   0
            Left            =   1890
            TabIndex        =   10
            Top             =   2250
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "pt(8 ~ 99)"
            Height          =   195
            Left            =   1890
            TabIndex        =   7
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "폰트 SIZE"
            Height          =   285
            Left            =   180
            TabIndex        =   5
            Top             =   315
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frmProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public frmMain As frmMain
Public frmQuiz As frmQuiz

Private Sub chkSetTimeOutS_Click()
If chkSetTimeOutS.Value = vbUnchecked Then
    If chkStreatX.Value = vbChecked Then
        chkStreatX.Value = vbGrayed
    End If
Else
    If chkStreatX.Value = vbGrayed Then
        chkStreatX.Value = vbChecked
    End If
End If

End Sub

Private Sub chkShowWhitSpot_Click()
If chkShowWhitSpot.Value = vbUnchecked And chkWhenResize.Value = vbChecked Then
    chkWhenResize.Value = vbGrayed
Else
    chkWhenResize.Value = IIf(CBool(chkWhenResize) = False, vbUnchecked, vbChecked)
    
End If
End Sub

Private Sub chkTTSuse_Click()
Call calChkTTS
End Sub

Private Sub chkTTSuse1_Click()
    Call calChkTTS
End Sub
Private Sub calChkTTS()
    If chkTTSuse.Value = vbUnchecked Then
        If chkTTSuse1.Value Then
            chkTTSuse1.Value = vbGrayed
        End If
        If chkTTSuse2.Value Then
            chkTTSuse2.Value = vbGrayed
        End If
        If chkTTSuse3.Value Then
            chkTTSuse3.Value = vbGrayed
        End If
        If chkTTSuse4.Value Then
            chkTTSuse4.Value = vbGrayed
        End If
    ElseIf chkTTSuse.Value = vbChecked Then
        If chkTTSuse1.Value = vbUnchecked And chkTTSuse2.Value = vbUnchecked And chkTTSuse3.Value = vbUnchecked And chkTTSuse4.Value = vbUnchecked Then
            chkTTSuse.Value = vbGrayed
        Else
            If chkTTSuse1 Then
                chkTTSuse1.Value = vbChecked
            End If
            If chkTTSuse2 Then
                chkTTSuse2.Value = vbChecked
            End If
            If chkTTSuse3 Then
                chkTTSuse3.Value = vbChecked
            End If
            If chkTTSuse4 Then
                chkTTSuse4.Value = vbChecked
            End If
            
        End If
        
    Else
        If Not (chkTTSuse1.Value = vbUnchecked And chkTTSuse2.Value = vbUnchecked And chkTTSuse3.Value = vbUnchecked And chkTTSuse4.Value = vbUnchecked) Then
            chkTTSuse.Value = vbChecked
        End If
        
    End If
    
    'If chkTTSuse1.Value = vbUnchecked and Then
    
End Sub

Private Sub chkTTSuse2_Click()
Call calChkTTS
End Sub

Private Sub chkTTSuse3_Click()
Call calChkTTS
End Sub

Private Sub chkTTSuse4_Click()
Call calChkTTS
End Sub

Private Sub chkWhenResize_Click()
If chkWhenResize.Value = vbChecked And chkShowWhitSpot.Value = vbUnchecked Then
    chkWhenResize.Value = vbGrayed
End If
End Sub

Private Sub cmdCancel_Click()
    txtFontSize.Text = cFTP.Profile.FontSize
    chkStreatX.Value = IIf(cFTP.Profile.bStreateOutCntCheck, IIf(cFTP.Profile.bSetTimeOutStudy, vbChecked, vbGrayed), vbUnchecked)
    chkSetTimeOutS.Value = IIf(cFTP.Profile.bSetTimeOutStudy, vbChecked, vbUnchecked)
    
    txtSetTimeOutS = cFTP.Profile.TimeOutSecStudy
    txtStreatX = cFTP.Profile.CntOfStreatOutSetting

    txtTimeOut = cFTP.Profile.TimeOutSec
    chkTimeOut = IIf(cFTP.Profile.bSetTimeOut, vbChecked, vbUnchecked)
    
    'With cFTP.Profile
        chkAlm1(0) = IIf(cFTP.Profile.bAlarm1, vbChecked, vbUnchecked)
        chkAlm1(1) = IIf(cFTP.Profile.bAlarm2, vbChecked, vbUnchecked)
        chkAlm1(2) = IIf(cFTP.Profile.bAlarm5, vbChecked, vbUnchecked)
        
'        chkTTSbeta = IIf(cFTP.Profile.bTTSbeta, vbChecked, vbUnchecked)
        chkTTSuse = IIf(cFTP.Profile.bChkTTSuse, vbChecked, vbUnchecked)
        chkTTSuse1 = IIf(cFTP.Profile.bChkTTSuse1, vbChecked, vbUnchecked)
        chkTTSuse2 = IIf(cFTP.Profile.bChkTTSuse2, vbChecked, vbUnchecked)
        chkTTSuse3 = IIf(cFTP.Profile.bChkTTSuse3, vbChecked, vbUnchecked)
        chkTTSuse4 = IIf(cFTP.Profile.bChkTTSuse4, vbChecked, vbUnchecked)
        
        Call calChkTTS
        
        txtAlm1(0) = cFTP.Profile.CntOfAlarm1
        txtAlm1(1) = cFTP.Profile.CntOfAlarm2
        txtAlm1(2) = cFTP.Profile.CntOfAlarm5
        chkBold = IIf(cFTP.Profile.FontBold, vbChecked, vbUnchecked)
        txtDelay1 = cFTP.Profile.DelayTime1
        txtDelay2 = cFTP.Profile.DelayTime2
        txtNansu = cFTP.Profile.nansu
        
        chkPoP1 = IIf(cFTP.Profile.setPoP1, vbChecked, vbUnchecked)
        chkPoP2 = IIf(cFTP.Profile.setPoP2, vbChecked, vbUnchecked)
        chkPoP3 = IIf(cFTP.Profile.setPoP3, vbChecked, vbUnchecked)
        chkCompactDB = IIf(cFTP.Profile.compactDB, vbChecked, vbUnchecked)
        
        chkShowWhitSpot = IIf(cFTP.Profile.ShowWhiteSpot, vbChecked, vbUnchecked)
        chkWhenResize = IIf(cFTP.Profile.checkWhenResize, vbChecked, vbUnchecked)
        
    
    
    'End With
    
End Sub

Private Sub cmdClose_Click()

If frmMain.mnuEdit2.Visible Then


    frmMain.setfontsize_menu_bysize cFTP.Profile.FontSize

    frmQuiz.quizDisp False
    
End If
Unload Me
End Sub

Private Sub cmdOK_Click()
If Not IDEMODE Then
    On Error Resume Next
End If
    'With cFTP.Profile
        cFTP.Profile.FontName = cboFont.Text
        cFTP.Profile.FontSize = txtFontSize
        cFTP.Profile.bSetTimeOut = CBool(chkTimeOut)
        cFTP.Profile.TimeOutSec = txtTimeOut
        cFTP.Profile.bSetTimeOutStudy = CBool(chkSetTimeOutS)
        cFTP.Profile.TimeOutSecStudy = txtSetTimeOutS
        cFTP.Profile.bStreateOutCntCheck = CBool(chkStreatX)
        cFTP.Profile.CntOfStreatOutSetting = txtStreatX
        
        cFTP.Profile.bAlarm1 = chkAlm1(0)
        cFTP.Profile.bAlarm2 = chkAlm1(1)
        cFTP.Profile.bAlarm5 = chkAlm1(2)
        
        cFTP.Profile.CntOfAlarm1 = txtAlm1(0)
        cFTP.Profile.CntOfAlarm2 = txtAlm1(1)
        cFTP.Profile.CntOfAlarm5 = txtAlm1(2)
        cFTP.Profile.FontBold = CBool(chkBold)
        
        cFTP.Profile.DelayTime1 = txtDelay1
        cFTP.Profile.DelayTime2 = txtDelay2
        cFTP.Profile.nansu = txtNansu
        
        cFTP.Profile.setPoP1 = CBool(chkPoP1) '삭제시 질문
        cFTP.Profile.setPoP2 = CBool(chkPoP2)
        cFTP.Profile.setPoP3 = CBool(chkPoP3)
        
        cFTP.Profile.compactDB = CBool(chkCompactDB)
        
        
        cFTP.Profile.ShowWhiteSpot = CBool(chkShowWhitSpot)
        cFTP.Profile.checkWhenResize = CBool(chkWhenResize)
        cFTP.Profile.FontName = cboFont.Text
'        cFTP.Profile.bTTSbeta = CBool(chkTTSbeta)
        cFTP.Profile.bChkTTSuse = CBool(chkTTSuse)
        cFTP.Profile.bChkTTSuse1 = CBool(chkTTSuse1)
        cFTP.Profile.bChkTTSuse2 = CBool(chkTTSuse2)
        cFTP.Profile.bChkTTSuse3 = CBool(chkTTSuse3)
        cFTP.Profile.bChkTTSuse4 = CBool(chkTTSuse4)
        cFTP.Profile.save
        
        
    'End With
     
End Sub

Private Sub Command1_Click()
cmdOK_Click
cmdClose_Click
End Sub

Private Sub Form_Load()
Call FillListWithFonts(cboFont)
    SSTab1.Tab = 0
    cmdCancel_Click
End Sub

Private Sub txtAlm1_DblClick(Index As Integer)
selall txtAlm1(Index)
End Sub

Private Sub txtAlm1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)
End Sub

Private Sub txtAlm1_LostFocus(Index As Integer)
SELFLAG = False
End Sub

Private Sub txtAlm1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If SELFLAG = False Then
    selall txtAlm1(Index)
    SELFLAG = True
End If

End Sub

Private Sub txtDelay1_DblClick()
selall txtDelay1
End Sub

Private Sub txtDelay1_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)
End Sub

Private Sub txtDelay1_LostFocus()
SELFLAG = False
End Sub

Private Sub txtDelay1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If SELFLAG = False Then
    selall txtDelay1
    SELFLAG = True
End If

End Sub

Private Sub txtDelay2_DblClick()
selall txtDelay2
End Sub

Private Sub txtDelay2_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)
End Sub

Private Sub txtDelay2_LostFocus()
SELFLAG = False
End Sub

Private Sub txtDelay2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If SELFLAG = False Then
    selall txtDelay2
    SELFLAG = True
End If
End Sub

Private Sub txtFontSize_DblClick()
selall txtFontSize
End Sub

Private Sub txtFontSize_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)

End Sub

Private Sub txtFontSize_LostFocus()
SELFLAG = False
End Sub

Private Sub txtFontSize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If SELFLAG = False Then
    selall txtFontSize
    SELFLAG = True
End If

End Sub

'Private Sub txtNansu_DblClick()
'SELALL txtNansu
'End Sub
'
'Private Sub txtNansu_KeyPress(KeyAscii As Integer)
'KeyAscii = ONLYNUM(KeyAscii)
'End Sub
'
'Private Sub txtNansu_LostFocus()
'SELFLAG = False
'End Sub

'Private Sub txtNansu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If SELFLAG = False Then
'    SELALL txtNansu
'    SELFLAG = True
'End If
'End Sub

Private Sub txtSetTimeOutS_DblClick()
selall txtSetTimeOutS
End Sub

Private Sub txtSetTimeOutS_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)

End Sub

Private Sub txtSetTimeOutS_LostFocus()
SELFLAG = False
End Sub

Private Sub txtSetTimeOutS_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If SELFLAG = False Then
    selall txtSetTimeOutS
    SELFLAG = True
End If

End Sub

Private Sub txtStreatX_DblClick()
selall txtStreatX
End Sub

Private Sub txtStreatX_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)

End Sub

Private Sub txtStreatX_LostFocus()
SELFLAG = False
End Sub

Private Sub txtStreatX_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If SELFLAG = False Then
    selall txtStreatX
    SELFLAG = True
End If

End Sub

Private Sub txtTimeOut_DblClick()
selall txtTimeOut
End Sub

Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)

End Sub

Private Sub txtTimeOut_LostFocus()
SELFLAG = False
End Sub

Private Sub txtTimeOut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If SELFLAG = False Then
    selall txtTimeOut
    SELFLAG = True
End If

End Sub
