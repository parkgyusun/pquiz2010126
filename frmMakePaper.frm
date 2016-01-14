VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMakePaper 
   Caption         =   "시험지만들기"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   Icon            =   "frmMakePaper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame fra5 
      Height          =   4020
      Left            =   1170
      TabIndex        =   36
      Top             =   1260
      Visible         =   0   'False
      Width           =   4605
      Begin VB.TextBox txt5_1 
         Height          =   330
         Left            =   990
         MaxLength       =   15
         TabIndex        =   42
         Top             =   1845
         Width           =   2130
      End
      Begin VB.ListBox lst5_2 
         Height          =   1500
         Left            =   180
         TabIndex        =   41
         Top             =   2340
         Width           =   2265
      End
      Begin VB.CommandButton cmdMake5 
         Caption         =   "만들기"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3195
         TabIndex        =   40
         Top             =   3285
         Width           =   1230
      End
      Begin VB.TextBox txt5_2 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   330
         Left            =   3465
         TabIndex        =   39
         Text            =   "100"
         Top             =   2475
         Width           =   690
      End
      Begin VB.ListBox lstSubj5 
         Height          =   780
         Left            =   990
         MultiSelect     =   2  '확장형
         TabIndex        =   38
         Top             =   585
         Width           =   2130
      End
      Begin VB.CommandButton cmdSearch5 
         Caption         =   "조회"
         Height          =   780
         Left            =   3195
         TabIndex        =   37
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "(※ 다중선택 가능) "
         Height          =   285
         Index           =   11
         Left            =   990
         TabIndex        =   49
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "문제수"
         Height          =   285
         Index           =   9
         Left            =   2655
         TabIndex        =   47
         Top             =   2475
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "시험지명"
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   46
         Top             =   1845
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "개"
         Height          =   285
         Index           =   1
         Left            =   4275
         TabIndex        =   45
         Top             =   2475
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "오늘의 필수 복습문제를 만듭니다."
         Height          =   330
         Left            =   225
         TabIndex        =   44
         Top             =   315
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "과목선택"
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   43
         Top             =   675
         Width           =   780
      End
   End
   Begin VB.Frame fra2 
      Height          =   3525
      Left            =   1170
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   4605
      Begin VB.TextBox txtPn 
         Height          =   330
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   6
         Top             =   810
         Width           =   2175
      End
      Begin VB.ListBox lstDate 
         Height          =   1500
         Left            =   180
         TabIndex        =   5
         Top             =   1890
         Width           =   2265
      End
      Begin VB.CommandButton cmdMake2 
         Caption         =   "만들기"
         Height          =   330
         Left            =   3195
         TabIndex        =   3
         Top             =   2835
         Width           =   1230
      End
      Begin VB.TextBox txtOld 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   330
         Left            =   1800
         TabIndex        =   2
         Text            =   "100"
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "문제수"
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   8
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "시험지명"
         Height          =   285
         Index           =   0
         Left            =   765
         TabIndex        =   7
         Top             =   2025
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "개"
         Height          =   285
         Left            =   2610
         TabIndex        =   4
         Top             =   1395
         Width           =   240
      End
      Begin VB.Label lblOld 
         Caption         =   "가장 오래전에 풀었던 문제"
         Height          =   330
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame fra3 
      Height          =   3525
      Left            =   1170
      TabIndex        =   15
      Top             =   1305
      Visible         =   0   'False
      Width           =   4605
      Begin VB.TextBox txt3_1 
         Height          =   330
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   18
         Top             =   765
         Width           =   2175
      End
      Begin VB.CommandButton cmdMake3 
         Caption         =   "만들기"
         Height          =   330
         Left            =   3150
         TabIndex        =   17
         Top             =   2745
         Width           =   1140
      End
      Begin VB.ListBox lstOX 
         Height          =   960
         Left            =   315
         TabIndex        =   16
         Top             =   2070
         Width           =   2175
      End
      Begin VB.TextBox txt3_2 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   330
         Left            =   1800
         TabIndex        =   19
         Text            =   "100"
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "맞은수/틀린수/문제수"
         Height          =   285
         Left            =   315
         TabIndex        =   23
         Top             =   1800
         Width           =   2220
      End
      Begin VB.Label Label1 
         Caption         =   "시험지명"
         Height          =   285
         Index           =   3
         Left            =   225
         TabIndex        =   21
         Top             =   810
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "문제수"
         Height          =   285
         Index           =   2
         Left            =   315
         TabIndex        =   20
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "맞은수 틀린수별 문제지를 만듭니다"
         Height          =   330
         Left            =   225
         TabIndex        =   22
         Top             =   315
         Width           =   3975
      End
   End
   Begin VB.Frame fra1 
      Height          =   3480
      Left            =   1170
      TabIndex        =   10
      Top             =   1350
      Visible         =   0   'False
      Width           =   4605
      Begin VB.CommandButton cmdPaper 
         Caption         =   "만들기"
         Height          =   330
         Left            =   3195
         TabIndex        =   12
         Top             =   1620
         Width           =   1230
      End
      Begin VB.ComboBox cboList 
         Height          =   300
         Left            =   225
         Style           =   2  '드롭다운 목록
         TabIndex        =   11
         Top             =   1620
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "시험 종류를 선택하세요"
         Height          =   240
         Left            =   225
         TabIndex        =   13
         Top             =   1260
         Width           =   2805
      End
   End
   Begin VB.Frame fra4 
      Height          =   4020
      Left            =   1170
      TabIndex        =   24
      Top             =   1260
      Visible         =   0   'False
      Width           =   4605
      Begin VB.CommandButton cmdSearch4 
         Caption         =   "조회"
         Height          =   780
         Left            =   3195
         TabIndex        =   35
         Top             =   630
         Width           =   1005
      End
      Begin VB.ListBox lstSubj 
         Height          =   780
         Left            =   990
         MultiSelect     =   2  '확장형
         TabIndex        =   34
         Top             =   585
         Width           =   2130
      End
      Begin VB.TextBox txt4_2 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   330
         Left            =   3465
         TabIndex        =   28
         Text            =   "100"
         Top             =   2475
         Width           =   690
      End
      Begin VB.CommandButton cmdMake4 
         Caption         =   "만들기"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3195
         TabIndex        =   27
         Top             =   3285
         Width           =   1230
      End
      Begin VB.ListBox lstRo 
         Height          =   1500
         Left            =   180
         TabIndex        =   26
         Top             =   2340
         Width           =   2265
      End
      Begin VB.TextBox txt4_1 
         Height          =   330
         Left            =   990
         MaxLength       =   15
         TabIndex        =   25
         Top             =   1845
         Width           =   2130
      End
      Begin VB.Label Label1 
         Caption         =   "(※ 다중선택 가능) "
         Height          =   285
         Index           =   10
         Left            =   990
         TabIndex        =   48
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "과목선택"
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   33
         Top             =   675
         Width           =   780
      End
      Begin VB.Label lbl4 
         Caption         =   "틀린 비중별로 문제를 만듭니다."
         Height          =   330
         Left            =   225
         TabIndex        =   32
         Top             =   315
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "개"
         Height          =   285
         Index           =   0
         Left            =   4275
         TabIndex        =   31
         Top             =   2475
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "시험지명"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   30
         Top             =   1845
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "문제수"
         Height          =   285
         Index           =   4
         Left            =   2655
         TabIndex        =   29
         Top             =   2475
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   330
      Left            =   5355
      TabIndex        =   14
      Top             =   5580
      Width           =   915
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4965
      Left            =   495
      TabIndex        =   9
      Top             =   495
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8758
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "1. 기본형 문제생성"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2. 확장형 문제생성"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "3. 맞은수틀린수별문제생성"
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "4. 틀린비중"
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "5. 필수점검문제"
            Object.Tag             =   "5"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMakePaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public parent As frmMain

Dim b1 As Boolean
Dim b2 As Boolean
Dim b3 As Boolean
Dim b45 As Boolean
Dim b6 As Boolean


Private Sub cmdClose_Click()
'b1 = False
b2 = False
b3 = False
b45 = False
b6 = False

Unload Me
End Sub

Private Sub cmdMake2_Click()
Dim succ As Boolean
Dim Cnt As Long
Cnt = CLng(txtOld.Text)
If txtPn.Text = "" Then
    MsgBox "시험지명을 입력하세요.", vbExclamation
    txtPn.SetFocus
    Exit Sub
End If

If Con.State = 1 Then
    succ = insert_pocketinfo2(txtPn.Text, Cnt)
Else
    Con_Open
    succ = insert_pocketinfo2(txtPn.Text, Cnt)
    Con_Close
End If
If succ Then
    parent.mnuRefresh_Click
End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub cmdMake3_Click()

Dim succ As Boolean
Dim Cnt As Long
Dim STRO As String
Dim STRX As String

Cnt = CLng(txt3_2.Text)
Dim vData
vData = Split(lstOX, vbTab)
STRO = vData(1)
STRX = vData(2)
If Trim(txt3_1.Text) = "" Then
    MsgBox "시험지명을 입력하세요.", vbExclamation
    txt3_1.SetFocus
    Exit Sub
End If

If Con.State = 1 Then
    succ = insert_pocketinfo3(txt3_1.Text, Cnt, STRO, STRX)
Else
    Con_Open
    succ = insert_pocketinfo3(txt3_1.Text, Cnt, STRO, STRX)
    Con_Close
End If
If succ Then
    parent.mnuRefresh_Click
End If

End Sub

Private Sub cmdMake4_Click()

    Dim succ As Boolean
    Dim Cnt As Long
    Dim STRO As String
    Dim STRX As String
    
    Cnt = CLng(txt4_2.Text)
    Dim vData
    vData = Split(lstOX, vbTab)
    If Trim(txt4_1.Text) = "" Then
        MsgBox "시험지명을 입력하세요.", vbExclamation
        txt4_1.SetFocus
        Exit Sub
    End If
    
    If Con.State = 1 Then
        succ = insert_pocketinfo4(txt4_1.Text, Cnt, lstSubj)
    Else
        Con_Open
        succ = insert_pocketinfo4(txt4_1.Text, Cnt, lstSubj)
        Con_Close
    End If
    If succ Then
        parent.mnuRefresh_Click
    End If

End Sub

Private Sub cmdMake5_Click()

    Dim succ As Boolean
    Dim Cnt As Long
    Dim STRO As String
    Dim STRX As String
    
    Cnt = CLng(txt5_2.Text)
    
    If Trim(txt5_1.Text) = "" Then
        MsgBox "시험지명을 입력하세요.", vbExclamation
        txt5_1.SetFocus
        Exit Sub
    End If
    
    If Con.State = 1 Then
        succ = insert_pocketinfo6(txt5_1.Text, Cnt, lstSubj5)
    Else
        Con_Open
        succ = insert_pocketinfo6(txt5_1.Text, Cnt, lstSubj5)
        Con_Close
    End If
    If succ Then
        parent.mnuRefresh_Click
    End If

End Sub

Private Sub cmdSearch4_Click()

Dim lRs As ADODB.Recordset

'##1
Dim conbef As Integer
If Con.State = 1 Then
    conbef = 1
Else
    Con_Open
    conbef = 0
End If

sSql = "select x/(o+x), count(*) from tu02 where userid='" & gUserid & "' and o+x >1  "

If Not lstSubj.Selected(0) Then
    sSql = sSql & vbNewLine & "and subj in (" & selectSeries(lstSubj, vbTab) & ")"
End If

sSql = sSql & vbNewLine & " group by x/(o+x) order by 1 desc"

Set lRs = Fn_SQLExec(sSql).rs

lstRo.Clear

If lRs.RecordCount = 0 Then

    status.Text = "조회결과 없음"
Else

    Do Until lRs.EOF
        lstRo.AddItem Format(lRs(0) * 100#, "00.00") & "    (" & lRs(1) & ")" & Space(200) & vbTab & lRs(1)
        lRs.MoveNext
    Loop
    status.Text = "조회 완료"
End If

'##2
If conbef = 1 Then
Else
    Con_Close
End If

End Sub

Private Sub cmdSearch5_Click()

Dim lRs As ADODB.Recordset

Dim ymd As String

ymd = GETYMD()

'●＃＆＊＠§※☆★○◎◇◆□■△▲▽▼
'→←↑↓↔〓◁◀▷▶♤♠♡♥♧♣⊙◈▣
'◐◑▒▤▥▨▧▦▩♨☏☎☜☞¶†↕↗↙
'↖↘♭♩♪♬㉿㈜№㏇™㏂㏘℡®ªº


'##1
Dim conbef As Integer
If Con.State = 1 Then
    conbef = 1
Else
    Con_Open
    conbef = 0
End If

sSql = "select reserve_ymd, count(*) from tu02 where userid='" & gUserid & "' and reserve_ymd<'" & ymd & "'  "

If Not lstSubj5.Selected(0) Then
    sSql = sSql & vbNewLine & "and subj in (" & selectSeries(lstSubj5, vbTab) & ")"
End If

sSql = sSql & vbNewLine & " group by reserve_ymd order by 1 asc"

Set lRs = Fn_SQLExec(sSql).rs

lst5_2.Clear

If lRs.RecordCount = 0 Then
    status.Text = "조회결과 없음" 'arrStatus(3)'조회결과 없음
Else
    Do Until lRs.EOF
        lst5_2.AddItem "점검" & lRs(0) & " (" & lRs(1) & ")" & Space(200) & vbTab & lRs(1)
        lRs.MoveNext
    Loop
    status.Text = "조회 완료"
End If


'##2
If conbef = 1 Then
Else
    Con_Close
End If
End Sub

Private Sub Form_Activate()
TabStrip1_Click
End Sub

Private Sub lst5_2_Click()
Dim vData As Variant
vData = Split(lst5_2, vbTab)

txt5_1 = Trim(vData(0))

txt5_2.Text = uppersum(lst5_2, lst5_2.ListIndex) ' GETLASTTXT(lstOX, vbTab)

If txt5_2.Text > 0 Then
    cmdMake5.Enabled = True
End If
End Sub

Private Sub lstDate_Click()
txtPn = Left(lstDate, 8)
txtOld.Text = uppersum(lstDate, lstDate.ListIndex) 'GETLASTTXT(lstDate, vbTab)
End Sub

Private Sub lstOX_Click()
Dim vData As Variant
vData = Split(lstOX, vbTab)

txt3_1 = Trim(vData(0))

txt3_2.Text = uppersum(lstOX, lstOX.ListIndex) ' GETLASTTXT(lstOX, vbTab)

End Sub

Private Function uppersum(lst As ListBox, idx As Integer) As Long
Dim sm As Long
Dim i As Integer
For i = 0 To idx
    sm = sm + GETLASTTXT(lst.List(i), vbTab)
Next
uppersum = sm
End Function

Private Sub lstRo_Click()
Dim vData As Variant
vData = Split(lstRo, vbTab)

txt4_1 = Trim(vData(0))

txt4_2.Text = uppersum(lstRo, lstRo.ListIndex) ' GETLASTTXT(lstOX, vbTab)

If txt4_2.Text > 0 Then
    cmdMake4.Enabled = True
End If

End Sub

Private Sub lstSubj_Click()
cmdMake4.Enabled = False
lstRo.Clear
End Sub

Private Sub TabStrip1_Click()
fra1.Visible = False
fra2.Visible = False
fra3.Visible = False
fra4.Visible = False
fra5.Visible = False

Dim txt As String
Dim gunsu As Long
Dim maxcnt As Long




If TabStrip1.SelectedItem.Tag = "1" Then
    fra1.Visible = True
    
    If b1 = False Then
        sSql = "select pocketnm from tp01 order by seq"
        Set rs = Fn_SQLExec(sSql).rs
        Do Until rs.EOF
            txt = rs(0)
        '    vData = Split(txt, "|")
            gunsu = getgunsu(rs(0))
            txt = txt & " (" & gunsu & " 문제)" & Space(200) & vbTab & rs(0)
        
        
            cboList.AddItem txt
            rs.MoveNext
        Loop
        b1 = True
    End If
    
ElseIf TabStrip1.SelectedItem.Tag = "2" Then
    fra2.Visible = True
    
    If b2 = False Then
        sSql = "select count(*) from tu02 where o+x>0 and userid='" & gUserid & "'"
        maxcnt = Fn_SQLExec(sSql).rs(0)
        If maxcnt = 0 Then
            cmdMake2.Enabled = False
        Else
            cmdMake2.Enabled = True
        End If
        
        txtOld.Text = maxcnt
        txtOld.MaxLength = Len(CStr(maxcnt))
        lblOld.Caption = "가장 오래전에 풀었던 문제(최대:" & Format(maxcnt, "##,##0") & " 개)"
        
        sSql = "select update_ymd,count(*) from tu02 where (o > 0 or x>0) and userid='" & gUserid & "' group by update_ymd "
        Dim rs1 As New ADODB.Recordset
        rs1.Open sSql, STRCON, 1, 3
        lstDate.Clear
        Do Until rs1.EOF
            lstDate.AddItem rs1(0) & " (" & rs1(1) & "건)" & Space(200) & vbTab & rs1(1)
            rs1.MoveNext
        Loop
        rs1.Close
        
        
        b2 = True
    End If
    
ElseIf TabStrip1.SelectedItem.Tag = "3" Then
    fra3.Visible = True
    
    If b3 = False Then
        sSql = "SELECT o, x, count(*) From tu02 WHERE userid='" & gUserid & "' and (o>0 or x>0) GROUP BY o, x ORDER BY x DESC , o"
        rs1.Open sSql, STRCON, 1, 3
        lstOX.Clear
        Do Until rs1.EOF
            lstOX.AddItem rs1(0) & "/" & rs1(1) & "/" & rs1(2) & Space(200) & vbTab & rs1(0) & vbTab & rs1(1) & vbTab & rs1(2)
            rs1.MoveNext
        Loop
        rs1.Close
        
        b3 = True
    End If
    
ElseIf TabStrip1.SelectedItem.Tag = "4" Then
    fra4.Visible = True
        
    If b45 = False Then
        lstSubj.Clear
        lstSubj5.Clear
        sSql = "select subj,subjnm from ts01" ' where userid='" & gUserid & "' and o+x>0"
        Set rs1 = Fn_SQLExec(sSql).rs
        lstSubj.AddItem "<전체>"
        lstSubj5.AddItem "<전체>"
        Do Until rs1.EOF
            lstSubj.AddItem rs1("subjnm") & Space(200) & vbTab & rs1("subj")
            lstSubj5.AddItem rs1("subjnm") & Space(200) & vbTab & rs1("subj")
            rs1.MoveNext
        Loop
        
        lstSubj.Selected(0) = True
        lstSubj5.Selected(0) = True
        
        b45 = True
    End If

    
    
ElseIf TabStrip1.SelectedItem.Tag = "5" Then
    fra5.Visible = True
    
    If b45 = False Then
        lstSubj.Clear
        lstSubj5.Clear
'        sSql = "select subj,subjnm from ts01 where cond not like '%select%'" ' userid='" & gUserid & "' and o+x>0"
        sSql = " select a.subj,a.subjnm ,b.cond from ts01 a ,tp01 b where a.subj=b.pocketnm and b.cond not like '%select%' "
        Set rs1 = Fn_SQLExec(sSql).rs
        lstSubj.AddItem "<전체>"
        lstSubj5.AddItem "<전체>"
        Do Until rs1.EOF
            lstSubj.AddItem rs1("subjnm") & Space(200) & vbTab & rs1("subj")
            lstSubj5.AddItem rs1("subjnm") & Space(200) & vbTab & rs1("subj")
            rs1.MoveNext
        Loop
        
        lstSubj.Selected(0) = True
        lstSubj5.Selected(0) = True
        
        b45 = True
    End If
Else
    Debug.Assert False
End If
End Sub

Private Sub txt3_2_DblClick()
selall txt3_2
End Sub

Private Sub txt3_2_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)

End Sub

Private Sub txt3_2_LostFocus()
SELFLAG = False
End Sub

Private Sub txt3_2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SELFLAG = False Then
    selall txt3_2
    SELFLAG = True
End If

End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
KeyAscii = ONLYNUM(KeyAscii)
End Sub

Private Sub txtOld_LostFocus()
    SELFLAG = False

End Sub

Private Sub txtOld_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SELFLAG = False Then
    selall txtOld
    SELFLAG = True
End If

End Sub

Private Sub cmdPaper_Click()
Dim succ As Boolean
If cboList.ListIndex = -1 Then
    MsgBox "시험 유형을 선택하세요.", vbExclamation
    Exit Sub
End If
If Con.State = 1 Then
    succ = insert_pocketinfo(GETLASTTXT(cboList.Text, " " & vbTab))
Else
    Con_Open
    succ = insert_pocketinfo(GETLASTTXT(cboList.Text, " " & vbTab))
    Con_Close
End If
If succ Then
    parent.mnuRefresh_Click
End If
End Sub

Private Sub Form_Load()
Dim txt As String
Dim maxcnt As Long
Dim vData As Variant

Screen.MousePointer = vbHourglass
fra2.ZOrder 0
fra3.ZOrder 0
fra4.ZOrder 0
fra5.ZOrder 0
fra1.ZOrder 0

fra5.ZOrder 0
fra5.Visible = True

'TabStrip1.Tabs.Item(1).Selected = True
TabStrip1.Tabs.Item(5).Selected = True

'fra1.Visible = True

cboList.Clear

Screen.MousePointer = vbDefault
End Sub

Public Function getgunsu(pn As String) As Long
    Dim cond As String
    Dim vData As Variant
    Dim rs1 As ADODB.Recordset

    cond = getcond(pn)
    vData = Split(cond, "|")
    Dim tName As String
    
    If UBound(vData) = 1 Then
        tName = vData(1)
        If InStr(tName, vData(0)) > 0 Then
            sSql = "select count(*) from " & tName & " as a"
        Else
            sSql = "select count(*) from " & tName & " where " & vData(0)
        End If
    Else
        tName = "tu02"
        sSql = "select count(*) from " & tName & " where userid='" & gUserid & "' and " & vData(0)
    End If
'    Set rs1 = New ADODB.Recordset
'    rs1.Open sSql, con, 1, 3
    
'    sSql = "select count(*) from " & tName & " where userid='" & gUserid & "' and " & vData(0)
On Error GoTo ErrTrap
    getgunsu = Fn_SQLExec(sSql).rs(0)
Exit Function
ErrTrap:
getgunsu = -1

End Function
Public Function getcond(pn As String) As String
    sSql = "select cond from tp01 where pocketnm='" & pn & "'"
    getcond = Fn_SQLExec(sSql).rs(0)
End Function

Private Sub txtOld_DblClick()
selall txtOld
End Sub
