VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "����(build date:2015-08-06)"
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
'��ɼ���:
'- ������ ���
'   ����â�� ���콺 ���ʹ�ư�� Ŭ���Ͽ� ������ ���Ⱑ �����մϴ�.
'- Ű����
'   ���� 1,2,3,4,5 Ű�� ���ð���
'   NumPad���� ���� ���ý� ���� �� ���� ���
'   ȭ��ǥŰ �������� ���Ʒ� ��ȣ �����̵� ����
'   ȭ��ǥŰ �������� ���� k Ű�� �� ȭ��ǥ, jŰ�� �Ʒ�ȭ��ǥ ��� (UNIX vi ������ Ŀ�� �̵� ���)
'   �����̽�Ű �� NŰ�� ���ͱ���̸� ���� ����
'   ����������:esc Ű ���
'   - A S D F G Ű�� ���� 1�� 2�� 3�� 4�� 5�� ������ �����ϴ� ���̸� �����ڵ��Է��� �Ƿ��� �ι� ������ �մϴ�.
'   - ������ �ι� ������ �ʾƵ� �Ǵ� ���� �� ������ư ���� Ŭ���Ǵ� ����� ���ڹ�ư(NumPad����)�Դϴ�.
'- TTS�����Ͽ��� ȯ�漳������ ���������� ���� ������ üũ�Ǿ������� ���������� ���� ����.
'- ���Ŭ��: ���콺�� Ŭ������ �ʰ� ��ư ������ üũ�ϴ� ����� ���� �׸��� ������ �˴ϴ�.
'- I,R,E Ű: �ð��帧 ���
'- H Ű: ��Ʈ
'- Z Ű: ��������
'
'������ 100�� �Ѿ��µ� ��Ǭ������ �ִ� ����:
':�ڽĽ������� ������ �������������� ��Ǭ������ üũ�� �մϴ�.
' �׷��� �� ����: �ڽĽ������� ����� �ڽĽ������� ������ ��Ǯ�� �������� ������ �����ϰ� �ϱ� ����
'- �ɸ��Ͱ� ������ �ȵ� �ٸ� ������ �ٲٰ� �;��:
' ���α׷��� ����Ǵ� ������ ������ Ȯ���ڰ� acs�� ������ �ֽ��ϴ�. ���Ͻô� �ɸ����� acs���ϸ� ���������� �νø� ���Ͻô� �ɸ��͸� ��Ÿ���ϴ�.
' ���൵�߿� �ɸ��ͻ�� üũ�� �����ϸ� �ٸ� �ɸ��ͷ� �ٲ�� �˴ϴ�.
'- �ڽ��������Ͼ������� 4���������� ���� ��Ʈ�� �� �� �ֽ��ϴ�.
'  ȯ�漳���� �����ϼ���.
'  �»� �Ͼ���: ����1
'  ��� �Ͼ���: ����2
'  ���Ͽ� ������: ����3
'  ���Ͽ� ������: ����4
'- ȸ�����Խ� �̸��� �ּ����� ���� ������ ���� �Է��ϸ� �����ּ� ������ �ƴ� ���·� ȸ�������� �����ϸ� �տ� ������ ȸ�����̵�� �ĺ����� �ʽ��ϴ�.
'
Private Sub Form_Load()
    Me.Caption = "����(build date: " + BUILD_NUMBER & ")"
End Sub

Private Sub Form_Resize()
txtHelp.Width = Me.Width - 100
txtHelp.Height = Me.Height - 380
txtHelp.Left = 0
txtHelp.Top = 0
End Sub
