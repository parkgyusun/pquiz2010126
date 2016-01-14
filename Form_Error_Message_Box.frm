VERSION 5.00
Begin VB.Form Form_Error_Message_Box 
   BorderStyle     =   4  '고정 도구 창
   ClientHeight    =   3735
   ClientLeft      =   3375
   ClientTop       =   3015
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form_Error_Message_Box"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text_View 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   3
      Top             =   60
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   9255
   End
   Begin VB.CommandButton Command_OK 
      Caption         =   "확인"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command_SaveTo 
      Caption         =   "파일로 저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "Form_Error_Message_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'******************************************************************************************
'   프로그램 에러 확인기
'------------------------------------------------------------------------------------------
'   작성자          : 염대우
'   작성일          : 1999년 9월 2일 목요일
'   제작 기간       : 1999년 8월 말~1999년 9월 2일
'------------------------------------------------------------------------------------------
'    이 소스는 염대우가 만들었습니다.
'   이 소스 코드는 VB의 여러가지 프로그램 에러에 대한 확인을 해주는 코드 입니다.
'   재 배포시에 위의 출처를 반드시 표시해야 합니다.
'   소스코드 수정시 마지막 수정자 이름, 마지막 수정날짜, 수정 기간과, 현제 버전을 수정하세요.
'   Copyright ⓒ 1999-2001 YDWSOFT. All rights reserved.
'   염대우(YDWSOFT) E-Mail : ydwsoftware@popsmail.com
'------------------------------------------------------------------------------------------
'   마지막 수정자   : 염대우
'   마지막 수정     : 1999년 9월 6일 월요일
'   수정 기간       : 1999년 9월 5일~1999년 9월 6일
'   현제 버전(빌드) : 1.02
'------------------------------------------------------------------------------------------
'    이 소스는 염대우가 수정했습니다.
'   재 배포시에 위의 출처를 반드시 표시해야 합니다.
'------------------------------------------------------------------------------------------
'******************************************************************************************
Private Sub Command_OK_Click()
    Unload Me
End Sub

Private Sub Command_SaveTo_Click()
    On Error GoTo err
    Screen.MousePointer = 11
    Dim TempFileName As String
    TempFileName = App.Path & "/" & Me.Caption & ".DAT"
    
    Dim Text_Back
    Dim Text_Next As String
    Open TempFileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, Text_Next
            Text_Back = Text_Back & vbCrLf & Text_Next
        Loop
    Close #1
    'Text_View = Text_Back
    'Exit Sub
    
    Open TempFileName For Output As #1
        Print #1, Text_View.Text & vbCrLf & vbCrLf & _
        "━━━━━━━━━━━━━━━━━━━━" & _
        "━━━━━━━━━━━━━━━━━━━━" _
                                                & vbCrLf & Text_Back
    Close #1
    
    Screen.MousePointer = 0
    
    Exit Sub
    
err:
    Select Case err
        Case 53
        Open TempFileName For Output As #1
            Print #1, Text_View.Text & vbCrLf & vbCrLf & _
            "━━━━━━━━━━━━━━━━━━━━" & _
            "━━━━━━━━━━━━━━━━━━━━" _
                                                    & vbCrLf & vbCrLf & Text_Back
    
        Close #1
        
        Case Else
'        ErrMsgProcM "Private Sub Command_SaveTo_Click()"
        
    End Select
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Text_View.Move 60, 60, Me.ScaleWidth - 120, Me.Height - 955
    Frame1.Move 0, Text_View.Height + 105, Me.Width
    Command_OK.Move Me.Width - 1560, Text_View.Height + 195
    Command_SaveTo.Move Me.Width - 3000, Text_View.Height + 195
End Sub
