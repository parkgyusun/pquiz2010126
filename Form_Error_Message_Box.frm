VERSION 5.00
Begin VB.Form Form_Error_Message_Box 
   BorderStyle     =   4  '���� ���� â
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
         Name            =   "����"
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
      ScrollBars      =   3  '�����
      TabIndex        =   3
      Top             =   60
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "Ȯ��"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���Ϸ� ����"
      BeginProperty Font 
         Name            =   "����"
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
'   ���α׷� ���� Ȯ�α�
'------------------------------------------------------------------------------------------
'   �ۼ���          : �����
'   �ۼ���          : 1999�� 9�� 2�� �����
'   ���� �Ⱓ       : 1999�� 8�� ��~1999�� 9�� 2��
'------------------------------------------------------------------------------------------
'    �� �ҽ��� ����찡 ��������ϴ�.
'   �� �ҽ� �ڵ�� VB�� �������� ���α׷� ������ ���� Ȯ���� ���ִ� �ڵ� �Դϴ�.
'   �� �����ÿ� ���� ��ó�� �ݵ�� ǥ���ؾ� �մϴ�.
'   �ҽ��ڵ� ������ ������ ������ �̸�, ������ ������¥, ���� �Ⱓ��, ���� ������ �����ϼ���.
'   Copyright �� 1999-2001 YDWSOFT. All rights reserved.
'   �����(YDWSOFT) E-Mail : ydwsoftware@popsmail.com
'------------------------------------------------------------------------------------------
'   ������ ������   : �����
'   ������ ����     : 1999�� 9�� 6�� ������
'   ���� �Ⱓ       : 1999�� 9�� 5��~1999�� 9�� 6��
'   ���� ����(����) : 1.02
'------------------------------------------------------------------------------------------
'    �� �ҽ��� ����찡 �����߽��ϴ�.
'   �� �����ÿ� ���� ��ó�� �ݵ�� ǥ���ؾ� �մϴ�.
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
        "����������������������������������������" & _
        "����������������������������������������" _
                                                & vbCrLf & Text_Back
    Close #1
    
    Screen.MousePointer = 0
    
    Exit Sub
    
err:
    Select Case err
        Case 53
        Open TempFileName For Output As #1
            Print #1, Text_View.Text & vbCrLf & vbCrLf & _
            "����������������������������������������" & _
            "����������������������������������������" _
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
