VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "도움말(build date:2015-08-06)"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8565
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00C0FFFF&
      Height          =   4185
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmHelp.frx":000C
      Top             =   60
      Width           =   8445
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'기능설명:
'- 형광펜 기능
'   바탕창에 마우스 왼쪽버튼을 클릭하여 형광펜 쓰기가 가능합니다.
'- 키조작
'   숫자 1,2,3,4,5 키로 선택가능
'   NumPad에서 숫자 선택시 선택 후 엔터 기능
'   화살표키 조작으로 위아래 번호 선택이동 가능
'   화살표키 조작행위 말고도 k 키는 위 화살표, j키는 아래화살표 기능 (UNIX vi 에디터 커서 이동 방식)
'   스페이스키 및 N키는 엔터기능이며 다음 진행
'   빠져나가기:esc 키 기능
'   - A S D F G 키는 각각 1번 2번 3번 4번 5번 문항을 선택하는 것이며 엔터자동입력이 되려면 두번 눌러야 합니다.
'   - 하지만 두번 누르지 않아도 되는 선택 후 다음버튼 까지 클릭되는 기능은 숫자버튼(NumPad포함)입니다.
'- TTS설정하여도 환경설정에서 복습문제는 읽지 않음에 체크되어있으면 복습문제는 읽지 않음.
'- 모션클릭: 마우스를 클릭하지 않고 버튼 위에서 체크하는 모양의 원을 그리면 선택이 됩니다.
'- I,R,E 키: 시간흐름 토글
'- H 키: 힌트
'- Z 키: 선택해제
'
'진도는 100이 넘었는데 안푼문제가 있는 이유:
':자식시험지를 만들경우 상위시험지에서 안푼문제로 체크를 합니다.
' 그렇게 한 이유: 자식시험지를 만들고 자식시험지를 지운경우 미풀이 문항으로 생성이 가능하게 하기 위함
'- 케릭터가 마음에 안들어서 다른 것으로 바꾸고 싶어요:
' 프로그램이 실행되는 동일한 폴더에 확장자가 acs인 파일이 있습니다. 원하시는 케릭터의 acs파일만 동일폴더에 두시면 원하시는 케릭터만 나타납니다.
' 실행도중에 케릭터사용 체크를 변경하면 다른 케릭터로 바뀌게 됩니다.
'- 박스주위에하얀점으로 4지선다형의 답의 힌트를 알 수 있습니다.
'  환경설정을 참고하세요.
'  좌상에 하얀점: 정답1
'  우상에 하얀점: 정답2
'  우하에 빨간점: 정답3
'  좌하에 빨간점: 정답4
'- 회원가입시 이메일 주소적는 란에 공백을 먼저 입력하면 메일주소 형식이 아닌 형태로 회원가입이 가능하며 앞에 공백은 회원아이디로 식별되지 않습니다.
'
Private Sub Form_Load()
    Me.Caption = "도움말(build date: " + BUILD_NUMBER & ")"
End Sub

Private Sub Form_Resize()
txtHelp.Width = Me.Width - 100
txtHelp.Height = Me.Height - 380
txtHelp.Left = 0
txtHelp.Top = 0
End Sub
