VERSION 5.00
Begin VB.Form frmTH01 
   Caption         =   "�����Ű���"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmTH01.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox txtUserID 
      Enabled         =   0   'False
      Height          =   330
      Left            =   960
      TabIndex        =   4
      Top             =   60
      Width           =   2655
   End
   Begin VB.TextBox txtHint 
      Height          =   3315
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   2
      Top             =   420
      Width           =   6375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�(&C)"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3840
      Width           =   1515
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "��ġ�Ϸ�(&S)"
      Height          =   375
      Left            =   1380
      TabIndex        =   0
      Top             =   3840
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "��ġ�Ͻ�"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmTH01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sSubj As String
Public sSeq As String
Public sUserid As String

Public saved As Boolean

Public Sub fillAll()
Con_Open

sSql = "select * from th01 where subj='" & sSubj & "' and seq=" & sSeq & " and userid like '@%' order by userid desc"

Set rs = Fn_SQLExec(sSql, , , True).rs

Dim sHint As Variant

'�̺κ� while ������ �ϸ� �ȵǳ�~������
If Not rs.EOF Then
    
    sHint = rs("hint")
    If IsNull(sHint) Then
        txtHint.Text = ""
    Else
        txtHint.Text = sHint
    End If
    
    sUserid = rs("userid")
    
    txtUserID.Text = sUserid
    
End If

Con_Close
End Sub

Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub cmdSave_Click()
Dim affected As Integer
Con_Open

sSql = "update th01 set "
sSql = sSql & " hint= '" & MakeDbTxt(txtHint.Text) & "' ,"
sSql = sSql & " userid=  concat('@',date_format(now(),'%Y%m%d%H%i%s'))"
sSql = sSql & " where subj='" & MakeDbTxt(sSubj) & "'"
sSql = sSql & " and seq=" & sSeq & " and userid='" & sUserid & "'"

affected = Fn_SQLExec(sSql).nrow
Debug.Assert affected = 1

Con_Close
saved = True

Me.Hide

End Sub

