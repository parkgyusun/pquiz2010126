VERSION 5.00
Begin VB.Form frmMemo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "메모"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8685
   Icon            =   "frmMemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   8685
      TabIndex        =   1
      Top             =   4170
      Width           =   8685
      Begin VB.CommandButton cmdTTS_cn 
         Caption         =   "중국어 읽기"
         Height          =   420
         Left            =   4410
         TabIndex        =   5
         Top             =   45
         Width           =   1455
      End
      Begin VB.CommandButton cmdTTS_ja 
         Caption         =   "일본어 읽기"
         Height          =   420
         Left            =   2955
         TabIndex        =   4
         Top             =   45
         Width           =   1455
      End
      Begin VB.CommandButton cmdTTS_kr 
         Caption         =   "한국어 읽기"
         Height          =   420
         Left            =   1500
         TabIndex        =   3
         Top             =   45
         Width           =   1455
      End
      Begin VB.CommandButton cmdTTS_en 
         Caption         =   "영어 읽기"
         Height          =   420
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   1455
      End
   End
   Begin VB.TextBox rtxt1 
      Height          =   1770
      IMEMode         =   10  'SBCS HANGUL
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   135
      Width           =   1545
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private refCquiz As clsQuiz
Public parent As frmQuiz

Private mvaruserid As String
Private mvarSubj As String
Private mvarSeq As Long
Private mvarMemo As String

Public frmQuiz_obj As frmQuiz

Private Sub cmdTTS_cn_Click()
frmQuiz_obj.clipText = rtxt1.Text
frmQuiz_obj.gSetLang = "zh-CN"
Call frmQuiz_obj.mnuTTS0_Click
frmQuiz_obj.gSetLang = ""

End Sub

Private Sub cmdTTS_cnClipboard_Click()

frmQuiz_obj.clipText = Clipboard.GetText
frmQuiz_obj.gSetLang = "zh-CN"
Call frmQuiz_obj.mnuTTS0_Click
frmQuiz_obj.gSetLang = ""


End Sub

Private Sub cmdTTS_en_Click()
frmQuiz_obj.clipText = rtxt1.Text
frmQuiz_obj.gSetLang = "en-US"
Call frmQuiz_obj.mnuTTS0_Click
frmQuiz_obj.gSetLang = ""

End Sub

Private Sub cmdTTS_ja_Click()
frmQuiz_obj.clipText = rtxt1.Text
frmQuiz_obj.gSetLang = "ja-JP"
Call frmQuiz_obj.mnuTTS0_Click
frmQuiz_obj.gSetLang = ""

End Sub

Private Sub cmdTTS_kr_Click()
frmQuiz_obj.clipText = rtxt1.Text
frmQuiz_obj.gSetLang = "ko-KR"
Call frmQuiz_obj.mnuTTS0_Click
frmQuiz_obj.gSetLang = ""

End Sub

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MemoLeft", 0)
    Me.Top = GetSetting(App.Title, "Settings", "MemoTop", 0)
    Me.Width = GetSetting(App.Title, "Settings", "MemoWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MemoHeight", 6500)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mvarSubj = refCquiz.subj And mvarSeq = refCquiz.seq Then
save
End If
'Debug.Assert False
End Sub

Private Sub Form_Resize()

Call rtxt1.Move(0, 0, Me.Width - 110, Me.Height - 392 - picBottom.Height)
If Me.Left < 0 Then
    Me.Left = 0
End If
End Sub

Public Sub setVal(userid, subj, seq, memo, ByRef clsMy As clsFTP, ByRef pcQuiz As clsQuiz)

If mvaruserid = userid And mvarSubj = subj And mvarSeq = seq Then
    'rtxt1.Text =
Else
    mvaruserid = userid
    mvarSubj = subj
    mvarSeq = seq
    rtxt1.Text = memo
    mvarMemo = memo
End If

rtxt1.Font.Size = clsMy.Profile.FontSize

rtxt1.Font.Bold = clsMy.Profile.FontBold
If clsMy.Profile.FontName <> "" Then
    rtxt1.Font.Name = clsMy.Profile.FontName
End If

Set refCquiz = pcQuiz

End Sub

Public Sub save()

    Dim affected As Long
    
    If rtxt1.Text = mvarMemo Then Exit Sub
    If Len(Trim(rtxt1.Text)) = 0 Then
        If Len(mvarMemo) > 0 Then
            If MsgBox("메모내용을 정말 지우시겠습니까?", vbYesNo + vbDefaultButton2 + vbQuestion, "질문") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    sSql = "update th01 set hint='" & MakeDbTxt(rtxt1.Text) & "' where userid='" & mvaruserid & "' and subj = '" & mvarSubj & "' and seq=" & mvarSeq & ""
    
    affected = Fn_SQLExec(sSql, , True).nrow
    
    If affected = 0 Then
        
        refCquiz.memo = rtxt1.Text
        
        sSql = "insert into th01(userid,subj,seq,hint,bshare) values('" & mvaruserid & "','" & mvarSubj & "'," & mvarSeq & ",'" & MakeDbTxt(rtxt1.Text) & "','1')"
        
        affected = Fn_SQLExec(sSql).nrow
        
        Debug.Assert affected > 0
        
    End If
    
'    refCquiz.memo = rtxt1.Text
    
'    If rtxt1.Text <> "" Then
'        Call parent.oTooltip.setTitle("", 0)
        
'        parent.oTooltip.ToolText(parent.cmdMemo) = Replace(rtxt1.Text, vbTab, Space(4))
'    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        SaveSetting App.Title, "Settings", "MemoLeft", Me.Left
        SaveSetting App.Title, "Settings", "MemoTop", Me.Top
        SaveSetting App.Title, "Settings", "MemoWidth", Me.Width
        SaveSetting App.Title, "Settings", "MemoHeight", Me.Height

End Sub



