VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApproach 
   BorderStyle     =   0  'None
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12735
   ControlBox      =   0   'False
   Icon            =   "frmApproach.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "FORM"
   Begin VB.CommandButton cboReword 
      Caption         =   "보상처리"
      Height          =   735
      Left            =   12510
      TabIndex        =   25
      Top             =   405
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog dlgChart 
      Left            =   2250
      Top             =   1395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3330
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   18
      Top             =   1170
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSChart20Lib.MSChart mchart 
      Height          =   4830
      Left            =   2205
      OleObjectBlob   =   "frmApproach.frx":000C
      TabIndex        =   1
      Top             =   1350
      Width           =   10320
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   12390
      Begin VB.CommandButton cmd1stSeries 
         Caption         =   "계열 쌓기"
         Height          =   330
         Left            =   3195
         TabIndex        =   19
         Top             =   720
         Width           =   1050
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "사용자별"
         Height          =   330
         Left            =   2115
         TabIndex        =   22
         Top             =   720
         Width           =   1050
      End
      Begin POCKETQUIZ.numText txtTo 
         Height          =   285
         Left            =   2295
         TabIndex        =   21
         Top             =   270
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatMask      =   ""
         FontSize        =   9
         FontName        =   "굴림"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      End
      Begin POCKETQUIZ.numText txtFrom 
         Height          =   285
         Left            =   990
         TabIndex        =   20
         Top             =   270
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatMask      =   ""
         FontSize        =   9
         FontName        =   "굴림"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      End
      Begin VB.ComboBox cbo3 
         Height          =   300
         ItemData        =   "frmApproach.frx":2359
         Left            =   990
         List            =   "frmApproach.frx":235B
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "닫기"
         Height          =   375
         Left            =   11385
         TabIndex        =   15
         Top             =   720
         Width           =   870
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "출력"
         Height          =   375
         Left            =   10485
         TabIndex        =   14
         Top             =   720
         Width           =   870
      End
      Begin VB.ComboBox cbo2 
         Height          =   300
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   765
         Width           =   1455
      End
      Begin VB.ComboBox cbo1 
         Height          =   300
         ItemData        =   "frmApproach.frx":235D
         Left            =   9000
         List            =   "frmApproach.frx":235F
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.ListBox lst2 
         Height          =   645
         Left            =   6525
         MultiSelect     =   2  'Extended
         TabIndex        =   9
         Top             =   315
         Width           =   1455
      End
      Begin VB.ListBox lst1 
         Height          =   645
         Left            =   4365
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   315
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         Default         =   -1  'True
         Height          =   375
         Left            =   10485
         TabIndex        =   4
         Top             =   315
         Width           =   1770
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Top             =   270
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         _Version        =   393216
         Format          =   87883777
         CurrentDate     =   38188
      End
      Begin MSComCtl2.DTPicker dtp2 
         Height          =   285
         Left            =   3105
         TabIndex        =   24
         Top             =   270
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         _Version        =   393216
         Format          =   87883777
         CurrentDate     =   38188
      End
      Begin VB.Label Label7 
         Caption         =   "조회조건"
         Height          =   285
         Left            =   135
         TabIndex        =   17
         Top             =   765
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "차트형식"
         Height          =   195
         Left            =   8190
         TabIndex        =   11
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "출력항목"
         Height          =   240
         Left            =   8190
         TabIndex        =   10
         Top             =   405
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "과목"
         Height          =   240
         Left            =   6075
         TabIndex        =   8
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "사용자"
         Height          =   240
         Left            =   3735
         TabIndex        =   7
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "~"
         Height          =   195
         Left            =   2115
         TabIndex        =   5
         Top             =   315
         Width           =   150
      End
      Begin VB.Label Label1 
         Caption         =   "조회일자"
         Height          =   285
         Left            =   135
         TabIndex        =   3
         Top             =   360
         Width           =   825
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4830
      Left            =   90
      TabIndex        =   0
      Top             =   1350
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   8520
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2025
      MousePointer    =   9  'Size W E
      Top             =   1350
      Width           =   150
   End
   Begin VB.Menu mnuChart 
      Caption         =   "챠트"
      Begin VB.Menu mnuColor 
         Caption         =   "채우기 색상"
      End
      Begin VB.Menu mnuPattern 
         Caption         =   "패턴 타입"
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 0"
            Index           =   0
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 1"
            Index           =   1
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 2"
            Index           =   2
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 3"
            Index           =   3
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 4"
            Index           =   4
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 5"
            Index           =   5
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 6"
            Index           =   6
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 7"
            Index           =   7
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 8"
            Index           =   8
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 9"
            Index           =   9
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 10"
            Index           =   10
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 11"
            Index           =   11
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 12"
            Index           =   12
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 13"
            Index           =   13
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 14"
            Index           =   14
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 15"
            Index           =   15
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 16"
            Index           =   16
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 17"
            Index           =   17
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 18"
            Index           =   18
         End
         Begin VB.Menu mnuType 
            Caption         =   "TYPE 19"
            Index           =   19
         End
      End
      Begin VB.Menu mnuP_Color 
         Caption         =   "패턴 색상"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuForPrint 
         Caption         =   "출력용 챠트"
      End
      Begin VB.Menu mnuForGeneral 
         Caption         =   "일반 챠트"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavePicture 
         Caption         =   "그림으로 저장"
      End
      Begin VB.Menu mnuFootNote 
         Caption         =   "각주 보이기"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmApproach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'기억량 조회 쿼리 예
'기억량-일:select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct reserve_ymd from tu02 where userid='영어' and reserve_ymd != '99999999' and reserve_ymd between '20051001' and '20051231')  b where a.userid='영어' and a.reserve_ymd != '99999999' and a.reserve_ymd>=b.reserve_ymd group by b.reserve_ymd,a.subj order by 1
'기억량이란. 시간이 지남에 따라 기억에 남은 수가 줄어들게 되는데 그 수치를 의미합니다.
'기억량-월:select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct concat(substring(reserve_ymd,1,6),'01') reserve_ymd from tu02 where userid='영어' and reserve_ymd != '99999999' and reserve_ymd between '20051001' and '20051231')  b where a.userid='영어' and a.reserve_ymd != '99999999' and a.reserve_ymd>=b.reserve_ymd group by b.reserve_ymd,a.subj order by 1
'[절 삭 율 ]이라는 명칭을 [기억량]으로 변경함 20060101


Private conflag As Boolean 'DB가 기존에 연결상태인가?
Private localrs As ADODB.Recordset

Dim strFootNote As String

Dim lastSelectIDX As Integer

Private mbMoving As Boolean
Const sglSplitLimit = 50

Private Sub cbo1_Click()
If cbo1.Text = "[기억량]" Then
    If cbo3.Text <> "[기억량-일]" And cbo3.Text <> "[기억량-월]" Then
        cbo3.Text = "[기억량-월]"
    End If
Else
    If cbo3.Text = "[기억량-일]" Or cbo3.Text = "[기억량-월]" Then
        cbo3.ListIndex = 0
    End If
End If
End Sub

Private Sub cbo2_Click()
Dim i As Long
mchart.chartType = CInt(GETLASTTXT(cbo2.Text, vbTab))
'With mchart
If mchart.chartType = VtChChartType2dCombination Then
    For i = 1 To mchart.ColumnCount
        mchart.Column = i
        mchart.SeriesColumn = 1
    Next
ElseIf mchart.chartType = VtChChartType3dCombination Then
    For i = 1 To mchart.ColumnCount
        mchart.Column = i
        mchart.SeriesColumn = 1
    Next

Else
    For i = 1 To mchart.ColumnCount
        mchart.Column = i
        mchart.SeriesColumn = i
    Next

End If
'End With

End Sub

Private Sub cbo3_Click()
If cbo3.Text = "[기억량-일]" Or cbo3.Text = "[기억량-월]" Then
    cbo1.Text = "[기억량]"
End If

End Sub

Private Sub cboReword_Click()

If MsgBox("보상처리 하시겠습니까?", vbQuestion + vbYesNo, "POCKETQUIZ") = vbYes Then

Dim str As String
Dim sid As Long ' 원래는 세션아이디로 해야하나

Randomize
sid = CLng(Rnd() * 100000)
str = "DROP TEMPORARY TABLE IF EXISTS tmp_tg01_" & sid

Call Fn_SQLExec(str)

str = "create TEMPORARY table tmp_tg01_" & sid & " as select gijunilja,userid,h,o_cnt-x_cnt reword_cnt,subj from tg01 where userid in (" & selectSeriesFrom0(lst1) & ") "
If lst2.ListIndex <> 0 Then
    str = str + " and subj in (" & selectSeriesFrom0(lst2) & ") "
End If
str = str + " and gijunilja between '" & txtFrom & "' and '" & txtTo & "' and o_cnt-x_cnt<>reword_cnt"

Call Fn_SQLExec(str)

Dim rs As ADODB.Recordset

str = "select * from tmp_tg01_" & sid
Set rs = Fn_SQLExec(str).rs

Dim Sum As Long

Sum = 0

Do Until rs.EOF
    str = "update tg01 set reword_cnt=" & rs(3) & " where userid='" & rs(1) & "' and gijunilja='" & rs(0) & "' and h=" & rs(2) & " and subj='" & rs(4) & "'"
    Sum = Sum + rs(3)
    Call Fn_SQLExec(str)
    rs.MoveNext
Loop

str = "DROP TEMPORARY TABLE IF EXISTS tmp_tg01_" & sid
Call Fn_SQLExec(str)

MsgBox "[" & Sum & "]개의 문항의 보상이 처리 되었습니다.", vbExclamation + vbOKOnly, "POCKETQUIZ"

End If

End Sub

Private Sub cmd1stSeries_Click()
Dim i As Long
'With mchart
    For i = 1 To mchart.ColumnCount
        mchart.Column = i
        mchart.SeriesColumn = 1
    Next
'End With

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo ErrTrap
Dim fname As String
Dim rval As Integer

rval = MsgBox("출력용 차트로 변환 후 출력하시겠습니까?", vbYesNoCancel + vbQuestion)

Select Case rval
Case vbYes
    mnuForPrint_Click
Case vbNo
    
Case vbCancel
    Exit Sub
End Select


mchart.EditCopy


GetTempFile "", "~rs", 0, fname

SavePicture Clipboard.GetData, fname

Printer.PaintPicture LoadPicture(fname), 1, 1
Printer.EndDoc

Kill fname

Exit Sub

ErrTrap:
Debug.Assert False
Debug.Print err.Description

End Sub

Private Sub cmdSearch_Click()
Dim userid_brace As String
Dim subj_brace As String

lastSelectIDX = 0

'항목을 미정산 금액으로 해야 하겠다.

If cbo1.Text = "[기억량]" And cbo3.Text <> "[기억량-일]" And cbo3.Text <> "[기억량-월]" Then
    MsgBox "기억량 조회시는 기억량-일,월 만 선택가능합니다.", vbExclamation
    cbo3.Text = "[기억량-일]"
    cbo3.SetFocus
    If gUserid <> "영어" Then
        Exit Sub
    End If
ElseIf cbo1.Text <> "[기억량]" And (cbo3.Text = "[기억량-일]" Or cbo3.Text = "[기억량-월]") Then
    MsgBox "기억량 일,월을 선택한 경우 기억량로만 선택가능합니다.", vbExclamation
    cbo1.Text = "[기억량]"
    cbo3.SetFocus
    If gUserid <> "영어" Then
        Exit Sub
    End If
End If
txtFrom.Text = Replace(txtFrom.Text, ",", "")
If Len(Replace(txtFrom.Text, ",", "")) <> 8 Then
    MsgBox "날짜형식(yyyyMMDD) 오류", vbExclamation
    txtFrom.SetFocus
    Exit Sub
End If

If Len(Replace(txtTo.Text, ",", "")) <> 8 Then
    MsgBox "날짜형식(yyyyMMDD) 오류", vbExclamation
    txtTo.SetFocus
    Exit Sub
End If

'If Left(txtFrom.Text, 4) <> Left(txtTo.Text, 4) Then
'    MsgBox "같은 년도에서만 조회가 가능합니다.", vbExclamation
'    txtFrom.SetFocus
'    Exit Sub
'End If

'review_cnt는 유효문항에 대한 보상 으로 책정한다.

If cbo3.ListIndex < 7 Then '0~6조건
    Select Case cbo3.ListIndex
    Case 0: '일
        If cbo1.ListIndex = 0 Then
            sSql = "select gijunilja , subj ,sum(totalsec) from tg01 "
        ElseIf cbo1.ListIndex = 1 Then
            sSql = "select gijunilja , subj ,sum(totalsec)/60 from tg01 "
        ElseIf cbo1.ListIndex = 2 Then
            sSql = "select gijunilja , subj ,sum(totalsec)/3600 from tg01 "
        ElseIf cbo1.ListIndex = 3 Then '
            sSql = "select gijunilja , subj ,sum(next_cnt-backNew_cnt-backReview_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 4 Then '
            sSql = "select gijunilja , subj ,sum(o_cnt-x_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 5 Then '
            'sSql = "select gijunilja , subj ,sum(new_cnt-backNew_cnt) from tg01 "
            sSql = "select gijunilja , subj ,sum(new_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 6 Then '
            'sSql = "select gijunilja , subj ,sum(review_cnt-backReview_cnt) from tg01 "
            sSql = "select gijunilja , subj ,sum(o_cnt-x_cnt-reword_cnt) from tg01 "
        End If
    Case 1: '주
        If cbo1.ListIndex = 0 Then
            sSql = "select ww , subj ,sum(totalsec) from tg01,yyyy "
        ElseIf cbo1.ListIndex = 1 Then
            sSql = "select ww , subj ,sum(totalsec)/60 ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 2 Then
            sSql = "select ww , subj ,sum(totalsec)/3600  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 3 Then '
            sSql = "select ww , subj ,sum(next_cnt-backNew_cnt-backReview_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 4 Then '
            sSql = "select ww , subj ,sum(o_cnt-x_cnt)  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 5 Then '
            'sSql = "select ww , subj ,sum(new_cnt-backNew_cnt)  ,yyyy from tg01 "
            sSql = "select ww , subj ,sum(new_cnt)  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 6 Then '
    '        sSql = "select ww , subj ,sum(review_cnt-backReview_cnt)  ,yyyy from tg01 "
             sSql = "select ww , subj ,sum(o_cnt-x_cnt-reword_cnt)  ,yyyy from tg01 "
        End If
    
    Case 2 'cbo3.AddItem "월"
        If cbo1.ListIndex = 0 Then
            sSql = "select (yyyy-2000)*100+m m , subj ,sum(totalsec) ,yyyy  from tg01 "
        ElseIf cbo1.ListIndex = 1 Then
            sSql = "select (yyyy-2000)*100+m m , subj ,sum(totalsec)/60  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 2 Then
            sSql = "select (yyyy-2000)*100+m m , subj ,sum(totalsec)/3600  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 3 Then '
            sSql = "select (yyyy-2000)*100+m m , subj ,sum(next_cnt-backNew_cnt-backReview_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 4 Then '
            sSql = "select (yyyy-2000)*100+m m , subj ,sum(o_cnt-x_cnt) ,yyyy  from tg01 "
        ElseIf cbo1.ListIndex = 5 Then '
            'sSql = "select m , subj ,sum(new_cnt-backNew_cnt)  ,yyyy from tg01 "
            sSql = "select (yyyy-2000)*100+m m , subj ,sum(new_cnt)  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 6 Then '
            'sSql = "select m , subj ,sum(review_cnt-backReview_cnt)  ,yyyy from tg01 "
            sSql = "select (yyyy-2000)*100+m m , subj ,sum(o_cnt-x_cnt-reword_cnt)  ,yyyy from tg01 "
        End If
    
    Case 3 'cbo3.AddItem "분기"
        If cbo1.ListIndex = 0 Then
            sSql = "select (yyyy-2000)*100+q , subj ,sum(totalsec)  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 1 Then
            sSql = "select (yyyy-2000)*100+q , subj ,sum(totalsec)/60  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 2 Then
            sSql = "select (yyyy-2000)*100+q , subj ,sum(totalsec)/3600  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 3 Then '
            sSql = "select (yyyy-2000)*100+q , subj ,sum(next_cnt-backNew_cnt-backReview_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 4 Then '
            sSql = "select (yyyy-2000)*100+q , subj ,sum(o_cnt-x_cnt)  ,yyyy from tg01 "
        ElseIf cbo1.ListIndex = 5 Then '
            'sSql = "select q , subj ,sum(new_cnt-backNew_cnt) ,yyyy  from tg01 "
            sSql = "select (yyyy-2000)*100+q , subj ,sum(new_cnt) ,yyyy  from tg01 "
        ElseIf cbo1.ListIndex = 6 Then '
            'sSql = "select q , subj ,sum(review_cnt-backReview_cnt)  ,yyyy from tg01 "
            sSql = "select (yyyy-2000)*100+q , subj ,sum(o_cnt-x_cnt-reword_cnt)  ,yyyy from tg01 "
        End If
    
    Case 4 'cbo3.AddItem "년"
        If cbo1.ListIndex = 0 Then
            sSql = "select yyyy , subj ,sum(totalsec) from tg01 "
        ElseIf cbo1.ListIndex = 1 Then
            sSql = "select yyyy , subj ,sum(totalsec)/60 from tg01 "
        ElseIf cbo1.ListIndex = 2 Then
            sSql = "select yyyy , subj ,sum(totalsec)/3600 from tg01 "
        ElseIf cbo1.ListIndex = 3 Then '
            sSql = "select yyyy , subj ,sum(next_cnt-backNew_cnt-backReview_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 4 Then '
            sSql = "select yyyy , subj ,sum(o_cnt-x_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 5 Then '
            'sSql = "select yyyy , subj ,sum(new_cnt-backNew_cnt) from tg01 "
            sSql = "select yyyy , subj ,sum(new_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 6 Then '
            'sSql = "select yyyy , subj ,sum(review_cnt-backReview_cnt) from tg01 "
            sSql = "select yyyy , subj ,sum(o_cnt-x_cnt-reword_cnt) from tg01 "
        End If
    
    Case 5 'cbo3.AddItem "요일별" '1일요일, 7:토요일
        If cbo1.ListIndex = 0 Then
            sSql = "select w , subj ,sum(totalsec) from tg01 "
        ElseIf cbo1.ListIndex = 1 Then
            sSql = "select w , subj ,sum(totalsec)/60 from tg01 "
        ElseIf cbo1.ListIndex = 2 Then
            sSql = "select w , subj ,sum(totalsec)/3600 from tg01 "
        ElseIf cbo1.ListIndex = 3 Then '
            sSql = "select w , subj ,sum(next_cnt-backNew_cnt-backReview_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 4 Then '
            sSql = "select w , subj ,sum(o_cnt-x_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 5 Then '
            'sSql = "select w , subj ,sum(new_cnt-backNew_cnt) from tg01 "
            sSql = "select w , subj ,sum(new_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 6 Then '
            'sSql = "select w , subj ,sum(review_cnt-backReview_cnt) from tg01 "
            sSql = "select w , subj ,sum(o_cnt-x_cnt-reword_cnt) from tg01 "
        End If
    
    Case 6 'cbo3.AddItem "시간대별" '24개로 보여줍니다.
        If cbo1.ListIndex = 0 Then
            sSql = "select h , subj ,sum(totalsec) from tg01 "
        ElseIf cbo1.ListIndex = 1 Then
            sSql = "select h , subj ,sum(totalsec)/60 from tg01 "
        ElseIf cbo1.ListIndex = 2 Then
            sSql = "select h , subj ,sum(totalsec)/3600 from tg01 "
        ElseIf cbo1.ListIndex = 3 Then '
            sSql = "select h , subj ,sum(next_cnt-backNew_cnt-backReview_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 4 Then '
            sSql = "select h , subj ,sum(o_cnt-x_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 5 Then '
            'sSql = "select h , subj ,sum(new_cnt-backNew_cnt) from tg01 "
            sSql = "select h , subj ,sum(new_cnt) from tg01 "
        ElseIf cbo1.ListIndex = 6 Then '
            'sSql = "select h , subj ,sum(review_cnt-backReview_cnt) from tg01 "
            sSql = "select h , subj ,sum(o_cnt-x_cnt-reword_cnt) from tg01 "
        End If
    
    End Select
    
    sSql = sSql & vbCrLf & "where gijunilja between '" & txtFrom.Text & "' and '" & txtTo.Text & "' "
    
    If lst1.List(0) <> "<전체>" Then
        sSql = sSql & vbCrLf & "and userid in (" & selectSeriesFrom0(lst1) & ")"
    ElseIf Not lst1.Selected(0) Then
        sSql = sSql & vbCrLf & "and userid in (" & selectSeries(lst1) & ")"
    End If
    
    If Not lst2.Selected(0) Then
        sSql = sSql & vbCrLf & "and SUBJ in (" & selectSeries(lst2) & ")"
    End If
    
    Select Case cbo3.ListIndex
    Case 0 'cbo3.AddItem "일"
        sSql = sSql & vbCrLf & "group by gijunilja , subj order by gijunilja "
    Case 1 'cbo3.AddItem "주"
        sSql = sSql & vbCrLf & "group by WW , subj ,yyyy order by yyyy,ww"
    Case 2 'cbo3.AddItem "월"
        sSql = sSql & vbCrLf & "group by M , subj ,yyyy order by yyyy,m"
    Case 3 '분기
        sSql = sSql & vbCrLf & "group by Q , subj ,yyyy order by yyyy,q"
    Case 4  '년"
        sSql = sSql & vbCrLf & "group by YYYY , subj  order by yyyy"
    Case 5 '1일요일, 7:토요일
        sSql = sSql & vbCrLf & "group by W , subj order by w"
    Case 6 '시간대별" '24개로 보여줍니다.
        sSql = sSql & vbCrLf & "group by H , subj order by h"
    Case Else
        Debug.Assert False
    End Select
    
    'Set localrs = New ADODB.Recordset
    
    If chkUser.Value = vbChecked Then
        sSql = Replace(sSql, "subj", "userid")
    End If
    
'   If 6 < cbo3.ListIndex Then
'    '------------------------------------------------------------------------------
'    'tg01 에서 조회하는것이 아닌 tu02에서 조회하는것으로 복습회수를 바꾼다.2004-12-11
'    '------------------------------------------------------------------------------
'        'If InStr(sSql, "select gijunilja , subj ,sum(review_cnt) from tg01") > 0 Then
''            sSql = Replace(sSql, "gijunilja", "update_ymd")
''            sSql = Replace(sSql, "where", "where gangyek > 0 and ")
''            sSql = Replace(sSql, "tg01", "tu02")
''            sSql = Replace(sSql, "review_cnt", "*")
''            sSql = Replace(sSql, "sum", "count")
''        ElseIf InStr(sSql, "select gijunilja , userid ,sum(review_cnt) from tg01") Then
'            sSql = Replace(sSql, "gijunilja", "update_ymd")
'            sSql = Replace(sSql, "where", "where gangyek > 0 and ")
'            sSql = Replace(sSql, "tg01", "tu02")
'            sSql = Replace(sSql, "o_cnt+x_cnt-review_cnt", "*")
'            sSql = Replace(sSql, "sum", "count")
''        End If
'    End If

Else '0~6조건이 아닌경우
    
    If lst1.List(0) <> "<전체>" Then
        userid_brace = " and userid in (" & selectSeriesFrom0(lst1) & ") "
        'userid_brace = Replace(userid_brace, ",''", "")
    ElseIf Not lst1.Selected(0) Then
        userid_brace = " and userid in (" & selectSeries(lst1) & ") "
        'userid_brace = Replace(userid_brace, ",''", "")
    End If
    
    If Not lst2.Selected(0) Then
        subj_brace = " and subj in (" & selectSeries(lst2) & ") "
        'subj_brace = Replace(subj_brace, ",''", "")
    End If

    Select Case cbo3.ListIndex
    Case 7 '기억량-일
        If InStr(Command, "/mdb") > 0 Then
            sSql = "select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct reserve_ymd from tu02 where 1=1 " & userid_brace & " and reserve_ymd <> '99999999' and reserve_ymd between '" & txtFrom.Text & "' and '" & txtTo.Text & "')  b where 1=1 " & userid_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>=b.reserve_ymd " & subj_brace & " group by b.reserve_ymd,a.subj order by 1"
        Else
            'sSql = "select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct reserve_ymd from tu02 where 1=1 " & userid_brace & " and reserve_ymd <> '99999999' and reserve_ymd between '" & txtFrom.Text & "' and '" & txtTo.Text & "')  b where 1=1 " & userid_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>=b.reserve_ymd " & subj_brace & " group by b.reserve_ymd,a.subj order by 1"
            sSql = "select '" & txtFrom.Text & "',a.subj,count(*) from tu02 a where 1=1 " & userid_brace & subj_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>='" & txtFrom.Text & "' group by a.subj union all "
            sSql = sSql + " select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct reserve_ymd from tu02 where 1=1 " & userid_brace & subj_brace & " and reserve_ymd between '" & txtFrom.Text & "' and '" & txtTo.Text & "')  b where 1=1 " & userid_brace & subj_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>=b.reserve_ymd " & subj_brace & " group by b.reserve_ymd,a.subj order by 1"
        End If
        
'        select '20150808',subj,count(*) from tu02
'where userid in ('박규선') and reserve_ymd <> '99999999' and
'reserve_ymd>='20150808' group by subj;

'기억량-일:
'기억량이란. 시간이 지남에 따라 기억에 남은 수가 줄어들게 되는데 그 수치를 의미합니다.
'기억량-월:select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct concat(substring(reserve_ymd,1,6),'01') reserve_ymd from tu02 where userid='영어' and reserve_ymd != '99999999' and reserve_ymd between '20051001' and '20051231')  b where a.userid='영어' and a.reserve_ymd != '99999999' and a.reserve_ymd>=b.reserve_ymd group by b.reserve_ymd,a.subj order by 1
        
        
    Case 8 '기억량-월
        If InStr(Command, "/mdb") > 0 Then
            sSql = "select b.reserve_y,a.subj,count(*) from tu02 a , (select distinct (mid(reserve_ymd,1,6) & '01') as reserve_y from tu02 where 1=1 " & userid_brace & " and reserve_ymd <> '99999999' and reserve_ymd between '" & txtFrom.Text & "' and '" & txtTo.Text & "')  b where 1=1 " & userid_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>=b.reserve_y " & subj_brace & " group by b.reserve_y,a.subj order by 1"
        Else
            'sSql = "select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct concat(substring(reserve_ymd,1,6),'01') reserve_ymd from tu02 where 1=1 " & userid_brace & " and reserve_ymd <> '99999999' and reserve_ymd between '" & txtFrom.Text & "' and '" & txtTo.Text & "')  b where 1=1 " & userid_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>=b.reserve_ymd " & subj_brace & " group by b.reserve_ymd,a.subj order by 1"
            sSql = "select '" & Mid(txtFrom.Text, 1, 6) & "01" & "',a.subj,count(*) from tu02 a where 1=1 " & userid_brace & subj_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>='" & Mid(txtFrom.Text, 1, 6) & "01" & "' group by a.subj union all "
            sSql = sSql + " select b.reserve_ymd,a.subj,count(*) from tu02 a , (select distinct concat(substring(reserve_ymd,1,6),'01') reserve_ymd from tu02 where 1=1 " & userid_brace & subj_brace & " and reserve_ymd between '" & txtFrom.Text & "' and '" & txtTo.Text & "')  b where 1=1 " & userid_brace & subj_brace & " and a.reserve_ymd <> '99999999' and a.reserve_ymd>=b.reserve_ymd " & subj_brace & " group by b.reserve_ymd,a.subj order by 1"
        End If
    End Select
    
    sSql = Replace(sSql, "1=1 and", "")
    If chkUser.Value = vbChecked Then
        'sSql = Replace(sSql, ",a.subj", ",a.userid") '콤마를 붙임에 유의!
        sSql = Replace(sSql, "a.subj", "a.userid") '콤마를 안붙여야 할 것 같음
        sSql = Replace(sSql, "union all", "union") '
    End If
    
End If

Set localrs = Fn_SQLExec(sSql, , , True).rs

If SQLCA.Error Then Exit Sub

Call makechart(mchart, localrs)

Set localrs = Nothing

strFootNote = "● 사용자: " & Replace(Replace(selectSeries(lst1), "''", ""), ",", " ") & Space(3) & " ● 과목: " & Replace(Replace(selectSeries(lst2), "''", ""), ",", " ") & " ● 기간:  " & txtFrom.Text & "  ~  " & txtTo

strFootNote = Replace(strFootNote, "● 사용자:     ●", " ●")
strFootNote = Replace(strFootNote, "● 과목:  ●", " ●")

strFootNote = strFootNote & Space(3) & "● 조회간격: " & cbo3.Text
strFootNote = strFootNote & "[" & cbo1.Text & "]"


If Len(mchart.Footnote.Text) > 0 Or mnuFootNote.Checked = True Then
    mchart.Footnote.Text = strFootNote
End If

If lastSelectIDX = 0 Then
    Call setChartColor
End If

End Sub


Sub setChartColor()
Dim serX As Series
Dim rval As Integer
Dim gVal As Integer
Dim bVal As Integer

Randomize

For Each serX In mchart.Plot.SeriesCollection
    rval = Rnd * 10000 Mod 256
    gVal = Rnd * 10000 Mod 256
    bVal = Rnd * 10000 Mod 256

    Call serX.DataPoints(-1).Brush.FillColor.Set(rval, gVal, bVal)
Next
End Sub


Private Sub dtp1_Change()
txtFrom.Text = date2Str(dtp1.Value)
End Sub

Private Sub dtp1_Click()
dtp1.Value = str2Date(txtFrom.Text)
End Sub

Private Sub dtp1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dtp1.Value = str2Date(txtFrom.Text)

End Sub

'Private Sub dtp1_KeyUp(KeyCode As Integer, Shift As Integer)
'
'End Sub

Private Sub dtp2_Click()
dtp2.Value = str2Date(txtTo.Text)
End Sub

Private Sub dtp2_CloseUp()
txtTo.Text = date2Str(dtp2.Value)
End Sub

Private Sub dtp2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dtp2.Value = str2Date(txtTo.Text)

End Sub

Private Sub Form_Load()

If Con.State Then
    conflag = True
Else
    Con_Open
End If

mnuChart.Visible = False

'모달로 떠있으므로 기존의 연결상태를 유지하거나
'새로운 연결로 계속 유지함.

    
    txtFrom.MaxVal = GETYMD() 'getMaxVal()
    txtFrom.MinVal = GETYMD() 'getMinVal()
    
    txtTo.MinVal = Left(txtTo.MaxVal, 4) & "0101" '  txtFrom.MinVal
    txtTo.MaxVal = txtTo.MaxVal
    
    txtFrom.Text = txtFrom.MinVal
    txtTo.Text = txtFrom.MaxVal
    
    Call adduser
    Call Addsubj
    
    Call addCHARTType
    Call addHangmok
    Call addGankyok

cbo2.ListIndex = 1
cbo3.ListIndex = 7 '
cbo1.ListIndex = 7

Dim i As Long
For i = 0 To lst1.ListCount - 1
    If gUserid = lst1.List(i) Then
        lst1.Selected(i) = True
    End If
Next

lst2.Selected(0) = True

cmdSearch_Click
End Sub

Public Function getMinVal() As Double
sSql = "select min(gijunilja) from tg01"
getMinVal = Fn_SQLExec(sSql).rs(0)
End Function

Public Function getMaxVal() As Double
sSql = "select max(gijunilja) from tg01"
getMaxVal = Fn_SQLExec(sSql).rs(0)
End Function

Public Sub adduser()
lst1.Clear
sSql = "select distinct userid from tg01 "
If gUserid <> "영어" Then
    sSql = sSql & " where userid = '" & gUserid & "' "
End If
Set localrs = Fn_SQLExec(sSql).rs
If gUserid = "영어" Then
    lst1.AddItem "<전체>"
End If
Do Until localrs.EOF
    lst1.AddItem localrs(0)
    localrs.MoveNext
Loop
End Sub

Public Sub Addsubj()
lst2.Clear
sSql = "select distinct subj from tg01"
Set localrs = Fn_SQLExec(sSql).rs
lst2.AddItem "<전체>"
Do Until localrs.EOF
    lst2.AddItem localrs(0)
    localrs.MoveNext
Loop

End Sub

Public Sub addCHARTType()

cbo2.Clear

cbo2.AddItem "2dArea" & Space(200) & vbTab & VtChChartType2dArea
cbo2.AddItem "2dBar" & Space(200) & vbTab & VtChChartType2dBar
cbo2.AddItem "2dCombination" & Space(200) & vbTab & VtChChartType2dCombination
cbo2.AddItem "2dLine" & Space(200) & vbTab & VtChChartType2dLine
cbo2.AddItem "2dPie" & Space(200) & vbTab & VtChChartType2dPie
cbo2.AddItem "2dStep" & Space(200) & vbTab & VtChChartType2dStep
cbo2.AddItem "2dXY" & Space(200) & vbTab & VtChChartType2dXY
cbo2.AddItem "3dArea" & Space(200) & vbTab & VtChChartType3dArea
cbo2.AddItem "3dBar" & Space(200) & vbTab & VtChChartType3dBar
cbo2.AddItem "3dCombination" & Space(200) & vbTab & VtChChartType3dCombination
cbo2.AddItem "3dLine" & Space(200) & vbTab & VtChChartType3dLine
cbo2.AddItem "3dStep" & Space(200) & vbTab & VtChChartType3dStep
    
End Sub

Public Sub addHangmok()
cbo1.Clear
cbo1.AddItem "학습시간(초)"
cbo1.AddItem "학습시간(분)"
cbo1.AddItem "학습시간(시)"

cbo1.AddItem "학습문항(manual)" 'next-pre

cbo1.AddItem "유효학습량" 'o+x-pre

cbo1.AddItem "신규학습량" 'new-pre/2

cbo1.AddItem "유효학습미보상량" 'new-pre/2

cbo1.AddItem "[기억량]" 'new-pre/2


End Sub


Private Sub Form_Resize()
Dim grdHeight As Single
Dim mchartWidth As Single

Dim Margin As Single

Margin = 10
grdHeight = Me.Height - grd.Top - Margin
If grdHeight < 10 Then
    grd.Height = 10
    mchart.Height = 10
Else
    grd.Height = grdHeight
    mchart.Height = grdHeight
    
End If

mchartWidth = Me.Width - mchart.Left - Margin - grd.Left * 2

If mchartWidth < 10 Then
    mchart.Width = 10
Else
    mchart.Width = mchartWidth
End If

imgSplitter.Height = grd.Height
   
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

If conflag Then

Else
    Con_Close
End If

End Sub

Private Sub addGankyok()
cbo3.Clear
cbo3.AddItem "일"

cbo3.AddItem "주"
cbo3.AddItem "월"

cbo3.AddItem "분기"

cbo3.AddItem "년"
cbo3.AddItem "요일별" '1일요일, 7:토요일
cbo3.AddItem "시간대별" '24개로 보여줍니다.

cbo3.AddItem "[기억량-일]"
cbo3.AddItem "[기억량-월]"


End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    With imgSplitter
        picSplitter.Move imgSplitter.Left, imgSplitter.Top, imgSplitter.Width \ 2, imgSplitter.Height - 20
 '   End With
    picSplitter.Visible = True
    mbMoving = True
    picSplitter.ZOrder 0
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit - 1500 Then
            picSplitter.Left = Me.Width - sglSplitLimit - 1500
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Sub SizeControls(X As Single)
    On Error Resume Next

    '너비를 설정합니다.
    If X < grd.Left + 30 Then
        X = grd.Left + 30
    End If
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    grd.Width = X - grd.Left
    imgSplitter.Left = X
    mchart.Left = X + 150
    mchart.Width = Me.Width - (grd.Left * 2 + 150 + grd.Width)
    
    '위쪽을 설정합니다.
  

'    If tbToolBar.Visible Then
'        grd.Top = tbToolBar.Height + picTitles.Height
'    Else
'        grd.Top = picTitles.Height
'    End If

  mchart.Top = grd.Top
    

    '높이를 설정합니다.
'    If sbStatusBar.Visible Then
'        grd.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
'    Else
'        grd.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
'    End If
    

    mchart.Height = grd.Height
    imgSplitter.Top = grd.Top
    imgSplitter.Height = grd.Height
    

End Sub

Public Sub makechart(chart As MSChart, rs As ADODB.Recordset)
Dim row As Long
Dim col As Long
Dim maxcol As Long
Dim maxrow As Long

If Not rs.EOF Then
maxrow = getMaxRow(rs)
Dim arr() As Variant
End If
Dim vSubj As Variant
Dim findsubj As Boolean

Dim preRowVal As String


row = 1
col = 1

'chart.ChartData

'arr(0)
'  arr(0,0) empty
'  arr(0,1) "국어"
'  arr(0,2) "영어"
'arr(1)
'   arr(1,0) "2004-02"
'   arr(1,1) 1
'   arr(1,2) 2
'arr(2)
'   arr(2,0) "2004-03"
'   arr(2,1) 3
'   arr(2,2) 4
'arr(3)
'   arr(3,0) "2004-04"
'   arr(3,1) 5
'   arr(3,2) 6

maxcol = 0

ReDim Preserve arr(maxrow, maxcol) As Variant
    
    
'With chart

'Debug.Assert rs.RecordCount > 0


If rs.BOF = False Then
'If rs.RecordCount > 0 Then
    rs.MoveFirst
End If

'If rs.EOF = False And (rs.RecordCount > 0 Or rs.RecordCount = -1) Then
'    rs.MoveFirst
'End If

Do Until rs.EOF
    preRowVal = "" & rs(0)
    arr(row, 0) = "　" & rs(0)
    
    col = 1
    
    findsubj = False
    Dim i As Long
    For i = 1 To UBound(arr, 2)
        If arr(0, i) = rs(1) Then
            findsubj = True
            arr(row, i) = Format(rs(2), "#0.00")
            Exit For
        End If
    Next
    
    If findsubj = False Then
        maxcol = maxcol + 1
        ReDim Preserve arr(maxrow, maxcol) As Variant
        arr(0, maxcol) = rs(1)
        arr(row, maxcol) = Format(rs(2), "#0.00")
    End If
    
    rs.MoveNext
    If Not rs.EOF Then
        If preRowVal <> "" & rs(0) Then
            row = row + 1
'            ReDim Preserve arr(maxrow, maxcol) As Variant
        End If
    End If
Loop
'End With

Dim j As Long
For i = 1 To UBound(arr, 1)
    For j = 0 To UBound(arr, 2)
        If IsEmpty(arr(i, j)) Then
            arr(i, j) = 0
        End If
    Next
Next

If UBound(arr, 1) > 0 Then
    mchart.ChartData = arr
    Call arr2grd(arr())
    mchart.Visible = True
    grd.Visible = True
Else
    MsgBox "조회된 자료가 없습니다.", vbExclamation
    mchart.Visible = False
    grd.Visible = False
End If

End Sub

Private Sub mchart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        If lastSelectIDX > 0 Then
           Call mchart.Plot.SeriesCollection.Item(lastSelectIDX).DataPoints(-1).Select
           mnuColor = True
           mnuP_Color = True
           mnuPattern = True
           
        Else
           mnuColor = False
           mnuP_Color = False
           mnuPattern = False
        End If
    
    
        Me.PopupMenu mnuChart
    End If
End Sub

Private Sub mchart_PlotSelected(MouseFlags As Integer, Cancel As Integer)
Debug.Print "plotselected", MouseFlags
End Sub

Private Sub MChart_SeriesActivated(Series As _
   Integer, MouseFlags As Integer, Cancel As Integer)

lastSelectIDX = Series

On Error GoTo ErrTrap
   ' The CommonDialog control is named dlgChart.
   
   
   Dim red, green, blue As Integer
   'With dlgChart ' CommonDialog object
      dlgChart.CancelError = True
      dlgChart.ShowColor
      red = RedFromRGB(dlgChart.Color)
      green = GreenFromRGB(dlgChart.Color)
      blue = BlueFromRGB(dlgChart.Color)
   'End With

   ' NOTE: Only the 2D and 3D line charts use the
   ' Pen object. All other types use the Brush.

   If mchart.chartType <> VtChChartType2dLine Or _
   mchart.chartType <> VtChChartType3dLine Then
      mchart.Plot.SeriesCollection(Series). _
         DataPoints(-1).Brush.FillColor. _
         Set red, green, blue
   Else
      mchart.Plot.SeriesCollection(Series).Pen. _
         VtColor.Set red, green, blue
   End If
Exit Sub
ErrTrap:

End Sub

Public Function RedFromRGB(ByVal rgb As Long) _
   As Integer
   ' The ampersand after &HFF coerces the number as a
   ' long, preventing Visual Basic from evaluating the
   ' number as a negative value. The logical And is
   ' used to return bit values.
   RedFromRGB = &HFF& And rgb
End Function

Public Function GreenFromRGB(ByVal rgb As Long) _
   As Integer
   ' The result of the And operation is divided by
   ' 256, to return the value of the middle bytes.
   ' Note the use of the Integer divisor.
   GreenFromRGB = (&HFF00& And rgb) \ 256
End Function

Public Function BlueFromRGB(ByVal rgb As Long) _
   As Integer
   ' This function works like the GreenFromRGB above,
   ' except you don't need the ampersand. The
   ' number is already a long. The result divided by
   ' 65536 to obtain the highest bytes.
   BlueFromRGB = (&HFF0000 And rgb) \ 65536
End Function

Private Function getMaxRow(rs As ADODB.Recordset) As Long
Dim preVal As String
Dim cnt As Long
cnt = 1

If Not rs.EOF Then
    rs.MoveFirst
End If

Do Until rs.EOF
    preVal = "" & rs(0)
    
    rs.MoveNext
    If Not rs.EOF Then
        If "" & rs(0) <> preVal Then
            cnt = cnt + 1
        End If
    End If
Loop
    
    getMaxRow = cnt
End Function



Sub arr2grd(arr() As Variant)
Dim R As Long, C As Long

Dim MaxWid() As Variant
Dim tmpWid As Single

grd.Clear

C = UBound(arr, 2)
R = UBound(arr, 1)

ReDim MaxWid(0 To C) As Variant

grd.rows = R + 1
grd.cols = C + 1

Dim i As Long
Dim j As Long

For i = 0 To R
    grd.row = i
    For j = 0 To C
        grd.col = j
        grd.Text = arr(i, j)
        tmpWid = Me.TextWidth(grd.Text)
        If tmpWid > MaxWid(j) Then
            MaxWid(j) = tmpWid
        End If
    Next

Next

For j = 0 To C

    grd.ColWidth(j) = MaxWid(j) + Me.TextWidth("가나")
    grd.row = 0
    grd.col = j
    grd.CellAlignment = 3
Next

End Sub

Public Function selectSeries(lst As ListBox) As String

Dim i As Integer
Dim nLen As Long
Dim str As String
For i = 1 To lst.ListCount - 1
    If lst.Selected(i) Then
        str = str & "'" & lst.List(i) & "',"
    End If
Next

str = str & "''"

selectSeries = str

End Function

Public Function selectSeriesFrom0(ByRef lst As ListBox) As String

Dim i As Integer
Dim nLen As Long
Dim str As String
For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then
        str = str & "'" & lst.List(i) & "',"
    End If
Next

str = str & "''"

selectSeriesFrom0 = str

End Function


Private Sub mchart_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
lastSelectIDX = Series
End Sub

Private Sub mnuColor_Click()
    Call MChart_SeriesActivated(lastSelectIDX, vbLeftButton, 0)
End Sub

Private Sub mnuFootNote_Click()
mnuFootNote.Checked = Not mnuFootNote.Checked
If mnuFootNote.Checked = True Then
    mchart.Footnote.Text = strFootNote
Else
    mchart.Footnote.Text = ""
End If
End Sub

Private Sub mnuForGeneral_Click()
Dim serX As Series
For Each serX In mchart.Plot.SeriesCollection
    serX.DataPoints(-1).Brush.Style = VtBrushStyleSolid
Next

End Sub

Private Sub mnuForPrint_Click()

Dim serX As Series
For Each serX In mchart.Plot.SeriesCollection
    serX.DataPoints(-1).Brush.Style = VtBrushStyleNull
Next


End Sub

Private Sub mnuP_Color_Click()

'lastSelectIDX = Series

On Error GoTo ErrTrap
   ' The CommonDialog control is named dlgChart.
   
   
   Dim red, green, blue As Integer
   'With dlgChart ' CommonDialog object
      dlgChart.CancelError = True
      dlgChart.ShowColor
      red = RedFromRGB(dlgChart.Color)
      green = GreenFromRGB(dlgChart.Color)
      blue = BlueFromRGB(dlgChart.Color)
   'End With

   ' NOTE: Only the 2D and 3D line charts use the
   ' Pen object. All other types use the Brush.

    mchart.Plot.SeriesCollection(lastSelectIDX). _
       DataPoints(-1).Brush.PatternColor. _
       Set red, green, blue

Exit Sub
ErrTrap:
End Sub

Private Sub mnuSavePicture_Click()
On Error GoTo ErrTrap
Dim fname As String
mchart.EditCopy

dlgChart.CancelError = True

dlgChart.Filter = "bmp파일(.bmp)|*.BMP"

dlgChart.ShowSave

fname = dlgChart.FileName

SavePicture Clipboard.GetData, fname

Exit Sub

ErrTrap:

End Sub

Private Sub mnuType_Click(Index As Integer)
    Select Case Index
    Case 0
        mchart.Plot.SeriesCollection(lastSelectIDX).DataPoints(-1).Brush.Style = VtBrushStyleNull
        'null
    Case 18
        'solid
        mchart.Plot.SeriesCollection(lastSelectIDX).DataPoints(-1).Brush.Style = VtBrushStyleSolid
    Case 19
        '빗금
        mchart.Plot.SeriesCollection(lastSelectIDX).DataPoints(-1).Brush.Style = VtBrushStyleHatched
    Case Else
        '패턴
        mchart.Plot.SeriesCollection(lastSelectIDX).DataPoints(-1).Brush.Style = VtBrushStylePattern
        mchart.Plot.SeriesCollection(lastSelectIDX).DataPoints(-1).Brush.Index = Index
    End Select
End Sub
