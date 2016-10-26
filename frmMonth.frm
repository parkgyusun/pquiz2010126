VERSION 5.00
Object = "{6B787742-7A47-11D3-9441-50AC0EC10000}#1.0#0"; "ctmonth.ocx"
Begin VB.Form frmMonth 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   Icon            =   "frmMonth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "FORM"
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   225
      TabIndex        =   1
      Top             =   180
      Width           =   8880
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "새로고침"
         Height          =   420
         Left            =   4365
         TabIndex        =   5
         Top             =   225
         Width           =   1050
      End
      Begin VB.CheckBox Check1 
         Caption         =   "주 번호"
         Height          =   285
         Left            =   810
         TabIndex        =   4
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "출력"
         Height          =   420
         Left            =   5955
         TabIndex        =   3
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         Height          =   420
         Left            =   7560
         TabIndex        =   2
         Top             =   225
         Width           =   1050
      End
   End
   Begin MonthLibCtl.ctMonth ctMonth1 
      Height          =   6315
      Left            =   675
      TabIndex        =   0
      Top             =   1440
      Width           =   8430
      _Version        =   65536
      _ExtentX        =   14870
      _ExtentY        =   11139
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMonth.frx":000C
      MonthNames      =   "1,2,3,4,5,6,7,8,9,10,11,12"
      PrintSubTitle   =   ""
      SelectBackColor =   14737632
      TitleTextColor  =   4210688
      HeaderBackColor =   13434828
      MaskColor       =   12632256
      Date            =   42651
      TitleType       =   1
      TitleSize       =   1
      HeaderBorder    =   1
      DateAlign       =   1
      DatePosition    =   0
      WeekNumAlign    =   0
      PrintMapMode    =   0
      Day             =   10
      Month           =   10
      Year            =   2016
      DateXOffset     =   -3
      DateYOffset     =   2
      KeyboardScan    =   0   'False
      MultiText       =   -1  'True
      TitleButtons    =   -1  'True
      TextWrap        =   -1  'True
      PicArray0       =   "frmMonth.frx":0028
      PicArray1       =   "frmMonth.frx":0044
      PicArray2       =   "frmMonth.frx":0060
      PicArray3       =   "frmMonth.frx":007C
      PicArray4       =   "frmMonth.frx":0098
      PicArray5       =   "frmMonth.frx":00B4
      PicArray6       =   "frmMonth.frx":00D0
      PicArray7       =   "frmMonth.frx":00EC
      PicArray8       =   "frmMonth.frx":0108
      PicArray9       =   "frmMonth.frx":0124
      PicArray10      =   "frmMonth.frx":0140
      PicArray11      =   "frmMonth.frx":015C
      PicArray12      =   "frmMonth.frx":0178
      PicArray13      =   "frmMonth.frx":0194
      PicArray14      =   "frmMonth.frx":01B0
      PicArray15      =   "frmMonth.frx":01CC
      PicArray16      =   "frmMonth.frx":01E8
      PicArray17      =   "frmMonth.frx":0204
      PicArray18      =   "frmMonth.frx":0220
      PicArray19      =   "frmMonth.frx":023C
      PicArray20      =   "frmMonth.frx":0258
      PicArray21      =   "frmMonth.frx":0274
      PicArray22      =   "frmMonth.frx":0290
      PicArray23      =   "frmMonth.frx":02AC
      PicArray24      =   "frmMonth.frx":02C8
      PicArray25      =   "frmMonth.frx":02E4
      PicArray26      =   "frmMonth.frx":0300
      PicArray27      =   "frmMonth.frx":031C
      PicArray28      =   "frmMonth.frx":0338
      PicArray29      =   "frmMonth.frx":0354
      PicArray30      =   "frmMonth.frx":0370
      PicArray31      =   "frmMonth.frx":038C
      PicArray32      =   "frmMonth.frx":03A8
      PicArray33      =   "frmMonth.frx":03C4
      PicArray34      =   "frmMonth.frx":03E0
      PicArray35      =   "frmMonth.frx":03FC
      PicArray36      =   "frmMonth.frx":0418
      PicArray37      =   "frmMonth.frx":0434
      PicArray38      =   "frmMonth.frx":0450
      PicArray39      =   "frmMonth.frx":046C
      PicArray40      =   "frmMonth.frx":0488
      PicArray41      =   "frmMonth.frx":04A4
      PicArray42      =   "frmMonth.frx":04C0
      PicArray43      =   "frmMonth.frx":04DC
      PicArray44      =   "frmMonth.frx":04F8
      PicArray45      =   "frmMonth.frx":0514
      PicArray46      =   "frmMonth.frx":0530
      PicArray47      =   "frmMonth.frx":054C
      PicArray48      =   "frmMonth.frx":0568
      PicArray49      =   "frmMonth.frx":0584
      PicArray50      =   "frmMonth.frx":05A0
      PicArray51      =   "frmMonth.frx":05BC
      PicArray52      =   "frmMonth.frx":05D8
      PicArray53      =   "frmMonth.frx":05F4
      PicArray54      =   "frmMonth.frx":0610
      PicArray55      =   "frmMonth.frx":062C
      PicArray56      =   "frmMonth.frx":0648
      PicArray57      =   "frmMonth.frx":0664
      PicArray58      =   "frmMonth.frx":0680
      PicArray59      =   "frmMonth.frx":069C
      PicArray60      =   "frmMonth.frx":06B8
      PicArray61      =   "frmMonth.frx":06D4
      PicArray62      =   "frmMonth.frx":06F0
      PicArray63      =   "frmMonth.frx":070C
      PicArray64      =   "frmMonth.frx":0728
      PicArray65      =   "frmMonth.frx":0744
      PicArray66      =   "frmMonth.frx":0760
      PicArray67      =   "frmMonth.frx":077C
      PicArray68      =   "frmMonth.frx":0798
      PicArray69      =   "frmMonth.frx":07B4
      PicArray70      =   "frmMonth.frx":07D0
      PicArray71      =   "frmMonth.frx":07EC
      PicArray72      =   "frmMonth.frx":0808
      PicArray73      =   "frmMonth.frx":0824
      PicArray74      =   "frmMonth.frx":0840
      PicArray75      =   "frmMonth.frx":085C
      PicArray76      =   "frmMonth.frx":0878
      PicArray77      =   "frmMonth.frx":0894
      PicArray78      =   "frmMonth.frx":08B0
      PicArray79      =   "frmMonth.frx":08CC
      PicArray80      =   "frmMonth.frx":08E8
      PicArray81      =   "frmMonth.frx":0904
      PicArray82      =   "frmMonth.frx":0920
      PicArray83      =   "frmMonth.frx":093C
      PicArray84      =   "frmMonth.frx":0958
      PicArray85      =   "frmMonth.frx":0974
      PicArray86      =   "frmMonth.frx":0990
      PicArray87      =   "frmMonth.frx":09AC
      PicArray88      =   "frmMonth.frx":09C8
      PicArray89      =   "frmMonth.frx":09E4
      PicArray90      =   "frmMonth.frx":0A00
      PicArray91      =   "frmMonth.frx":0A1C
      PicArray92      =   "frmMonth.frx":0A38
      PicArray93      =   "frmMonth.frx":0A54
      PicArray94      =   "frmMonth.frx":0A70
      PicArray95      =   "frmMonth.frx":0A8C
      PicArray96      =   "frmMonth.frx":0AA8
      PicArray97      =   "frmMonth.frx":0AC4
      PicArray98      =   "frmMonth.frx":0AE0
      PicArray99      =   "frmMonth.frx":0AFC
   End
End
Attribute VB_Name = "frmMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public parent As frmMain

Dim nYear As Integer
Dim nMonth As Integer
Dim lRs As ADODB.Recordset

Dim startDate As String
Dim endDate As String

Dim startDatePre As String
Dim endDatePre As String

Dim nCnt As Long

Private Sub Check1_Click()
    If Check1.Value = 0 Then
       ctMonth1.WeekNumbers = False
    Else
       ctMonth1.WeekNumbers = True
    End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

    ctMonth1.PrintCalendar
    
End Sub

Private Sub Command1_Click()
Call disp(2004, 8)
End Sub

Private Sub cmdRefresh_Click()
ctMonth1.ClearDays
Call disp(ctMonth1.Year, ctMonth1.Month)
End Sub

Private Sub ctMonth1_DateChange(nDOW As Integer, nDay As Integer, pnMonth As Integer, pnYear As Integer)
'If nYear <= 1900 Then Exit Sub
If pnMonth <> nMonth Or pnYear <> nYear Then
    ctMonth1.ClearDays
    Call disp(pnYear, pnMonth)
End If

End Sub

Private Sub Form_Load()

ctMonth1.Year = Year(Now)
ctMonth1.Month = Month(Now)

Call disp(ctMonth1.Year, ctMonth1.Month)

End Sub

Private Sub Form_Resize()
Dim lLeft As Single
Dim lTop As Single
Dim lFra As Frame

ctMonth1.Left = (Me.Width - ctMonth1.Width) / 2
ctMonth1.Top = (Me.Height - ctMonth1.Height) / 2
End Sub

Private Function disp(y As Integer, M As Integer)

'Debug.Assert False

nYear = y
nMonth = M


startDate = nYear & Right("0" & nMonth, 2) & "01"
endDate = nYear & Right("0" & nMonth, 2) & Right("0" & getLastDay(nYear, nMonth), 2)

startDatePre = date2Str(str2Date(startDate) - 1)
endDatePre = date2Str(str2Date(endDate) - 1)


'1 선긋는 놈을 찾는다.(pcode=0 fromilja >= 이달말일 and toilja <=이달초일
'If Not disp_pro1() Then Exit Function

'2 위에서 전체 갯수를 구한다.
'If Not disp_pro2() Then Exit Function

'3. 위에서  현재Row의 255비율을 구한다. = 255*현재Row/전체Row
'If Not disp_pro3() Then Exit Function

'4. 선을 긋는다.
'If Not disp_pro4() Then Exit Function

'5. pcode<>0 으로 조회했을때 toilja인 tu02와 tu02 join해서 count구해서tu03 o+x=0인것을 newcount를 구한다.
If Not disp_pro5() Then Exit Function

'6.노란색으로 복습일정 색칠하기 복습일정이 있는 날인 갯수를 뿌려주고 R123
If Not disp_pro6() Then Exit Function

'7 선긋는 색상: N: (55 153 55) R(152 153 55) NR(55 153 150)
If Not disp_pro7() Then Exit Function

If Not disp_pro1() Then Exit Function
If Not disp_pro2() Then Exit Function
If Not disp_pro3() Then Exit Function

End Function

'==============================================================================
'1 선긋는 놈을 찾는다.(pcode=0 fromilja >= 이달말일 and toilja <=이달초일
'==============================================================================
Private Function disp_pro1() As Boolean
    
    sSql = "select * from tp02 where userid='" & gUserid & "' and hidden=0 and pcode=0 and from_ymd <= '" & endDate & "' and to_ymd >= '" & startDate & "' order by from_ymd desc"
    Set lRs = Fn_SQLExec(sSql).rs
    
    nCnt = lRs.RecordCount
disp_pro1 = True
End Function

'==============================================================================
'2 위에서 전체 갯수를 구한다.
'==============================================================================
Private Function disp_pro2() As Boolean
'구했다 nCnt
disp_pro2 = True
End Function

'==============================================================================
'3. 위에서  현재Row의 360비율을 구한다. = 360*현재Row/전체Row
'==============================================================================
Private Function disp_pro3() As Boolean
Dim hue As Long
Dim i As Long
Dim nColor As Long
Do Until lRs.EOF
    i = i + 1
    hue = Round15(rndVal(0, 239))
    nColor = HSLToRGB201(CLng(hue), 100, 60)
    Call disp_pro4(lRs("from_ymd"), lRs("to_ymd"), nColor)
    lRs.MoveNext
Loop
lRs.Close

disp_pro3 = True
End Function

'==============================================================================
'4. 선을 긋는다.
'==============================================================================
Private Function disp_pro4(sFrom As String, sTo As String, nColor As Long) As Boolean

Static barIdx As Integer

barIdx = barIdx + 1

If barIdx > 100 Then
    barIdx = 0
End If

    Dim i1 As Integer
    Dim i2 As Integer
    
    Dim i As Integer
    
    If Mid(sFrom, 1, 6) = nYear & Right("0" & nMonth, 2) Then
        i1 = CInt(Right(sFrom, 2))
    Else
        i1 = 1
    End If
    
    If Mid(sTo, 1, 6) = nYear & Right("0" & nMonth, 2) Then
        i2 = CInt(Right(sTo, 2))
    Else
        i2 = getLastDay(nYear, nMonth)
    End If
    
    If (barIdx Mod 3) = 1 Then
        For i = i1 To i2
            ctMonth1.TaskBarColor1(i) = nColor
        Next
    ElseIf (barIdx Mod 3) = 2 Then
        For i = i1 To i2
            ctMonth1.TaskBarColor2(i) = nColor
        Next
    ElseIf (barIdx Mod 3) = 0 Then
        For i = i1 To i2
            ctMonth1.TaskBarColor3(i) = nColor
        Next
    End If
    
    disp_pro4 = True
    
End Function


'==============================================================================
'5. pcode<>0 으로 조회했을때 toilja인 tu02와 tu02 join해서 count구해서
'   tu03 o+x=0인것을 newcount를 구한다.
'==============================================================================
Private Function disp_pro5() As Boolean
    Dim i1 As Integer
    sSql = "select from_ymd,count(*) as cnt from tp02 a,tp03 b where a.userid='" & gUserid & "' and "
    sSql = sSql & "a.userid=b.userid and a.pcode<>0 and a.from_ymd <= '" & endDate & "' "
    sSql = sSql & "and a.to_ymd >= '" & startDate & "' and a.pocketnm=b.pocketnm and a.chasu=b.chasu and b.o+b.x = 0 "
    sSql = sSql & "group by a.from_ymd having a.from_ymd between '" & startDate & "' and '" & endDate & "' and count(*) > 0"
    
    Set lRs = Fn_SQLExec(sSql, , , True).rs
    
    If lRs Is Nothing Then Exit Function
    
    Do Until lRs.EOF
        i1 = CInt(Right(lRs("from_ymd"), 2))
        ctMonth1.DateText(i1) = "ⓝ" & lRs("cnt")
        ctMonth1.DateColor(i1) = rgb(182, 228, 255)
        lRs.MoveNext
    Loop
    
    disp_pro5 = True
End Function

'==============================================================================
'6.노란색으로 복습일정 색칠하기 복습일정이 있는 날인 갯수를 뿌려주고 R123
'==============================================================================
Private Function disp_pro6() As Boolean

    Dim i1 As Integer
    sSql = "select reserve_ymd,count(*) as cnt from tu02 where userid='" & gUserid & "' and reserve_ymd between '" & startDatePre & "' and '" & endDatePre & "' "
    sSql = sSql & "group by reserve_ymd"
    
    Set lRs = Fn_SQLExec(sSql, , , True).rs
    
    Do Until lRs.EOF
        i1 = CInt(Format(str2Date(lRs("reserve_ymd")) + 1, "d"))
        
        If Len(ctMonth1.DateText(i1)) > 0 Then
            ctMonth1.DateText(i1) = ctMonth1.DateText(i1) & " " & "ⓡ" & lRs("cnt") '중복
            ctMonth1.DateColor(i1) = rgb(255, 174, 239)
        Else
            ctMonth1.DateText(i1) = "ⓡ" & lRs("cnt")
            ctMonth1.DateColor(i1) = rgb(255, 247, 179)
        End If
        lRs.MoveNext
    Loop
disp_pro6 = True
End Function

'==============================================================================
'7 선긋는 색상: N: (55 153 55) R(152 153 55) NR(55 153 150)
'==============================================================================
Private Function disp_pro7() As Boolean

disp_pro7 = True

End Function


