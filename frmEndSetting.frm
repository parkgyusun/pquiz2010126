VERSION 5.00
Begin VB.Form frmEndSetting 
   Caption         =   "종료설정"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   Icon            =   "frmEndSetting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   180
      TabIndex        =   19
      Top             =   4080
      Width           =   4695
      Begin VB.OptionButton optUseOrNot 
         Caption         =   "설정"
         Height          =   495
         Index           =   0
         Left            =   960
         TabIndex        =   21
         ToolTipText     =   "사용으로 설정합니다."
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optUseOrNot 
         Caption         =   "해제"
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   20
         ToolTipText     =   "사용을 해제합니다"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdApp2 
      Caption         =   "해제로 적용 후 닫기"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdApp1 
      Caption         =   "설정으로 적용 후 닫기"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소(&C)"
      Height          =   375
      Left            =   2920
      TabIndex        =   4
      Top             =   5640
      Width           =   915
   End
   Begin VB.Timer tmrEnd 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2460
      Top             =   2100
   End
   Begin VB.CheckBox chkQuestion 
      Caption         =   "종료시 질문"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      ToolTipText     =   "설정을 해제하면 강제 종료됩니다."
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtEndMin 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "30"
      Top             =   2820
      Width           =   735
   End
   Begin VB.TextBox txtQuiz 
      Height          =   315
      Left            =   4200
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "30"
      Top             =   1380
      Width           =   495
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      Left            =   4200
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "60"
      Top             =   900
      Width           =   495
   End
   Begin VB.CheckBox chkQuiz 
      Caption         =   "퀴즈시작시 종료시간 설정 보이기"
      Height          =   375
      Left            =   300
      TabIndex        =   8
      Top             =   1320
      Width           =   3195
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "적용 후 닫기(&O)"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "닫기(&X)"
      Height          =   375
      Left            =   4020
      TabIndex        =   5
      Top             =   5640
      Width           =   915
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "적용(&A)"
      Height          =   375
      Left            =   1820
      TabIndex        =   3
      Top             =   5640
      Width           =   915
   End
   Begin VB.ListBox lstEndMin 
      Height          =   1620
      ItemData        =   "frmEndSetting.frx":000C
      Left            =   240
      List            =   "frmEndSetting.frx":002B
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox chkLogin 
      Caption         =   "로그인시 종료 시간 설정 보이기"
      Height          =   435
      Left            =   300
      TabIndex        =   6
      Top             =   840
      Width           =   3195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   5160
      Y1              =   5050
      Y2              =   5050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   -120
      X2              =   5040
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label3 
      Caption         =   "분"
      Height          =   255
      Index           =   3
      Left            =   4740
      TabIndex        =   18
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "기본"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   17
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "기본"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   1380
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "분"
      Height          =   255
      Index           =   0
      Left            =   4740
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lbEndDesc 
      Alignment       =   2  'Center
      Caption         =   "00분 00초 후 종료 예정"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   180
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "분 후 종료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   2880
      Width           =   1635
   End
End
Attribute VB_Name = "frmEndSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dateEndTime As Date
Public isForce As Boolean
Private bTimerEnabled As Boolean
Public parent As frmMain


Private Sub Check1_Click()

End Sub

Private Sub cmdApp1_Click()
'optUseOrNot(0).Value = True
cmdOK_Click

End Sub

Private Sub cmdApp2_Click()
'optUseOrNot(1).Value = True
cmdOK_Click

End Sub

Private Sub cmdApply_Click()

'tmrEnd.Enabled = False 'optUseOrNot(0).Value

If chkLogin.Value = vbChecked Then
    cFTP.Profile.endSettingLogin = "1"
Else
    cFTP.Profile.endSettingLogin = "2"
End If

If chkQuiz.Value = vbChecked Then
    cFTP.Profile.endSettingQuiz = "1"
Else
    cFTP.Profile.endSettingQuiz = "2"
End If

If IsNumeric(txtLogin.Text) Then
    cFTP.Profile.endSettingLoginDefaultMin = txtLogin.Text
End If

If IsNumeric(txtQuiz.Text) Then
    cFTP.Profile.endSettingQuizDefaultMin = txtQuiz.Text
End If

cFTP.Profile.endSettingSetMin = txtEndMin.Text
If chkQuestion.Value = vbChecked Then
    cFTP.Profile.endSettingQuestion = "True"
Else
    cFTP.Profile.endSettingQuestion = "False"
End If

cFTP.Profile.save
    
    While Not IsNumeric(txtEndMin.Text)
        MsgBox "종료분을 숫자형으로 입력하세요.", vbOKOnly + vbExclamation
        txtEndMin.SetFocus
    Wend

Dim ssVal As Long, mmVal As Long

If optUseOrNot(0).Value Then
    dateEndTime = DateAdd("n", CInt(txtEndMin.Text), Now())
    lbEndDesc.Visible = True
    '종료시간 재설정
    
    ssVal = CLng(DateDiff("s", Now(), dateEndTime))
'        Me.Tag = ssVal
    mmVal = ssVal \ 60
    ssVal = ssVal - mmVal * 60
    lbEndDesc.Caption = "" & mmVal & "분 " & ssVal & "초 후 종료"
'    tmrEnd.Interval = 1000
Else
    lbEndDesc.Visible = False
'    tmrEnd.Interval = 1000
End If

tmrEnd.Enabled = optUseOrNot(0).Value

End Sub

Private Sub cmdCancel_Click()

If bTimerEnabled Then
    optUseOrNot(0).Value = True
Else
    optUseOrNot(1).Value = True
End If
lbEndDesc.Visible = optUseOrNot(0).Value
tmrEnd.Enabled = optUseOrNot(0).Value

If cFTP.Profile.endSettingQuiz = "1" Then
    chkQuiz.Value = vbChecked
Else
    chkQuiz.Value = vbUnchecked
End If

If cFTP.Profile.endSettingLogin = "1" Then
    chkLogin.Value = vbChecked
Else
    chkLogin.Value = vbUnchecked
End If

txtLogin.Text = cFTP.Profile.endSettingLoginDefaultMin
txtQuiz.Text = cFTP.Profile.endSettingQuizDefaultMin

txtEndMin.Text = cFTP.Profile.endSettingSetMin

If cFTP.Profile.endSettingQuestion Then
    chkQuestion.Value = vbChecked
Else
    chkQuestion.Value = vbUnchecked
End If



'tmrEnd.Interval = 3000 '3초마다
'Me.Hide

End Sub

Private Sub cmdClose_Click()
    If lbEndDesc.Visible Then
        tmrEnd.Enabled = True
        optUseOrNot(0).Value = True
    Else
        tmrEnd.Enabled = False
        optUseOrNot(1).Value = True
    End If

txtEndMin.Text = cFTP.Profile.endSettingSetMin

If tmrEnd.Enabled = False Then
    parent.sbStatusBar.Panels(3).Style = sbrTime
End If

'tmrEnd.Interval = 10000 '10초마다

Me.Hide

End Sub

Private Sub cmdOK_Click()
cmdApply_Click
cmdClose_Click
End Sub

Private Sub Form_Activate()

isForce = False

    Dim ssVal As Double, mmVal As Double
    
    bTimerEnabled = tmrEnd.Enabled

    If tmrEnd.Enabled Then
        cmdClose.Default = True
        cmdClose.SetFocus
        optUseOrNot(0).Value = True
        lbEndDesc.Visible = True
        ssVal = CLng(DateDiff("s", Now(), dateEndTime))
'        Me.Tag = ssVal
        mmVal = ssVal \ 60
        ssVal = ssVal - mmVal * 60
        lbEndDesc.Caption = "" & mmVal & "분 " & ssVal & "초 후 종료"
'        tmrEnd.Interval = 1000 '1초 마다
        tmrEnd.Enabled = False 'true 전에 false로 해서 타이머 예약을 바꾼다.
        tmrEnd.Enabled = True '10초 후 예약을 지우고 1초후 예약으로 재설정 하기 위해 쓰잘데기 없어 보이는 작업을 함
'        txtEndMin.Text = cFTP.Profile.endSettingSetMin
    Else
        lbEndDesc.Visible = False
        If lbEndDesc.Visible Then
            optUseOrNot(1).Value = True
        End If
'        txtEndMin.Text = cFTP.Profile.endSettingSetMin
        
    End If
    
    If (cFTP.Profile.endSettingQuestion) Then
        chkQuestion.Value = vbChecked
    Else
        chkQuestion.Value = vbUnchecked
    End If
    
    
    If cFTP.Profile.endSettingLogin = "1" Then
        chkLogin.Value = vbChecked
    Else
        chkLogin.Value = vbUnchecked
    End If
    
    If cFTP.Profile.endSettingQuiz = "1" Then
        chkQuiz.Value = vbChecked
    Else
        chkQuiz.Value = vbUnchecked
    End If
    
    txtLogin.Text = cFTP.Profile.endSettingLoginDefaultMin
    
    txtQuiz.Text = cFTP.Profile.endSettingQuizDefaultMin
End Sub


Private Sub List1_Click()

End Sub



Private Sub Form_Unload(Cancel As Integer)
  If isForce = False Then
    Cancel = 1
    Me.Hide
  End If
End Sub

Private Sub lstEndMin_Click()
Dim LC As Integer
txtEndMin.Text = lstEndMin.Text

End Sub

Private Sub optUseOrNot_Click(Index As Integer)

If Index = 0 Then
    cmdApp1.Default = True
Else
    cmdApp2.Default = True
End If
    
    
End Sub

Private Sub optUseOrNot_DblClick(Index As Integer)
If Index = 0 Then
    cmdApp1_Click
Else
    cmdApp2_Click
End If
End Sub

Private Sub tmrEnd_Timer()

If Now() > dateEndTime Then
    Dim frmMsgBox As frmMsgBox60sec
    Set frmMsgBox = New frmMsgBox60sec
    Load frmMsgBox

    If chkQuestion.Value = vbChecked Then
        
        frmMsgBox.setMsg "종료시간 설정에 의해 종료됩니다."
        TmrAfterTTS_exit = True
        frmMsgBox.Show vbModal, parent
        TmrAfterTTS_exit = False
        
        If frmMsgBox.btn_flag = 0 Then
            If Not (fMainForm Is Nothing) Then
                Unload fMainForm
            End If
            End
        Else
            tmrEnd.Enabled = False
        End If

                
        'If MsgBox("종료시간 설정에 의해 종료됩니다.", vbOKCancel + vbExclamation) = vbOK Then
        '    End
        'Else
        '    tmrEnd.Enabled = False
        'End If
    Else
    
        frmMsgBox.btn_type = vbOKOnly
        frmMsgBox.setMsg "수고하셨습니다." & vbCrLf & vbCrLf & "[" & txtEndMin & "]분 동안 학습하셨습니다."
        TmrAfterTTS_exit = True
        frmMsgBox.Show vbModal, parent
        End
        
        'MsgBox "수고하셨습니다." & vbCrLf & vbCrLf & "[" & txtEndMin & "]분 동안 학습하셨습니다.", vbOKOnly + vbExclamation
        'End
    End If
    
    Unload frmMsgBox
Else

Dim ssVal As Long, mmVal As Long

ssVal = CLng(DateDiff("s", Now(), dateEndTime))
'        Me.Tag = ssVal
mmVal = ssVal \ 60
ssVal = ssVal - mmVal * 60
lbEndDesc.Caption = "" & mmVal & "분 " & ssVal & "초 후 종료"

'    If Me.Visible Then
        'tmrEnd.Interval = 1000 '1초마다
        'parent.sbStatusBar.Panels(3).Style = sbrTime
 '   Else
        parent.sbStatusBar.Panels(3).Style = sbrText
        'parent.sbStatusBar.Panels(3).Picture.
        parent.sbStatusBar.Panels(3).Text = lbEndDesc.Caption
        
  '  End If
End If

End Sub



Private Sub txtEndMin_Change()
    optUseOrNot(0).Value = True
End Sub

Private Sub txtEndMin_GotFocus()
txtEndMin.SelStart = 0
txtEndMin.SelLength = Len(txtEndMin)
End Sub
