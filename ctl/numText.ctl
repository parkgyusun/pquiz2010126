VERSION 5.00
Begin VB.UserControl numText 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   ScaleHeight     =   2880
   ScaleWidth      =   4215
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      CausesValidation=   0   'False
      Height          =   1500
      Left            =   180
      TabIndex        =   0
      Text            =   "0"
      Top             =   270
      Width           =   3300
   End
End
Attribute VB_Name = "numText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'UserControl�� �׽�Ʈ�ϰų� ������ϰų� �˼��ϴµ� �ʿ��� �۾����� �Ʒ��� ������� �����߽��ϴ�.
'
'A) UserControl �׽�Ʈ ���α׷� �����
'
'ǥ�� EXE�� ��Ʈ���� �����ߴ��� �Ǵ� UserControl�� ���� ActiveX ��Ʈ�� ������Ʈ�� ����������� ���� UserControl �׽�Ʈ ���α׷��� ��ġ�ϴ� ����� �� ���� �ֽ��ϴ�.
'
'ActiveX ��Ʈ�� ������Ʈ�� ������ٸ� �Ʒ� ������ ���� �׽�Ʈ ���α׷��� ��ġ�մϴ�.
'
'1)  UserControl�� �����մϴ�.
'2) ���� ��忡�� ��Ʈ���� ��ġ�Ϸ��� UserControl�� �����̳ʸ� �ݽ��ϴ�.
'3) �׽�Ʈ ������Ʈ�� ���� ������ �ʾҴٸ� [����] �޴����� [������Ʈ �߰�]�� �����ؼ� ǥ�� EXE ������Ʈ�� �߰��մϴ�.
'4) ���� ���ڿ��� UserControl �������� �� �� ���� Form1�� UserControl �ν��Ͻ��� ��ġ�մϴ�. �ʿ��ϴٸ� ��Ʈ���� �ű�ų� ũ�⸦ ������ �� �ֽ��ϴ�.
'5) ������Ʈ �׷��� �����մϴ�. ������ ���߰� �׽�Ʈ ���ǿ��� ������Ʈ �׷��� ���� �Ѳ����� �� ������Ʈ�� ��� �� �� �ֽ��ϴ�.
'
'���� ǥ�� EXE ������Ʈ�� UserControl�� �����ߴٸ� �Ʒ� ������ �����ʽÿ�.
'
'1)  UserControl�� �����մϴ�.
'2)  ���� ��忡�� ��Ʈ���� ��ġ�Ϸ��� UserControl�� �����̳ʸ� �ݽ��ϴ�.
'3)  ������Ʈ â���� ǥ�� EXE ������Ʈ�� Form1�� �� �� ���� �����̳ʸ� ���ϴ�.
'4)  ���� ���ڿ��� UserControl �������� �� �� ���� Form1�� UserControl �ν��Ͻ��� ��ġ�մϴ�. �ʿ��ϴٸ� ��Ʈ���� �ű�ų� ũ�⸦ ������ �� �ֽ��ϴ�.
'
'B) ������ ���� ���� ��忡�� ��Ʈ�� ������ �׽�Ʈ�ϱ�
'
'1)  �׽�Ʈ ������Ʈ���� Form1�� ��ġ�� ��Ʈ���� �����ϰ� <F4>Ű�� ���� �Ӽ� â�� ���ϴ�. ��Ʈ�ѿ� �߰��� �Ӽ��� ��Ÿ������ �׸��� ������ �� �ִ��� Ȯ���մϴ�.
'2)  Form1�� ���� �� �ٽ� ���ϴ�. �׷� ���� ��Ʈ���� �Ӽ����� �ùٷ� ����ǰ� �˻��Ǿ����� Ȯ���մϴ�.
'3)  Form1�� ��ġ�� ��Ʈ���� �� �� ������ �ڵ� â�� ���콺 ������ ����(���ν���)�� ������ �� ��Ÿ���� ��Ӵٿ� ��Ͽ� �ش� �̺�Ʈ�� ��Ÿ������ Ȯ���մϴ�
'4)  ��Ʈ���� �̺�Ʈ ���ν����� �׽�Ʈ �ڵ带 �߰��մϴ�.
'5)  �ٸ� ��Ʈ�ѵ��� �߰��ϰ� ��Ʈ���� �̺�Ʈ ���ν����� �ڵ带 ������ ���� ��Ʈ���� �Ӽ��� �޼����� ��Ÿ�� ������ �׽�Ʈ�մϴ�.
'6)  <F5>Ű�� ���� �׽�Ʈ ������Ʈ�� �����ϰ� ��Ʈ���� ��Ÿ�� ������ �׽�Ʈ�մϴ�.
'
'C) �˼��� ������ ������ ����� ���� ��Ʈ�� �����(�����簡 �������� ���� �ڵ� �߰�)
'
'1)  ���� ���� ��� ��Ʈ���� ������ �Ϻ� �̺�Ʈ�� �Ӽ��� ���� ���� ��� ��Ʈ�ѿ� �����ؾ� �մϴ�. ���� ��� BackColor �Ӽ��� �Ƹ��� UserControl�� Label ��Ʈ���� BackColor �Ӽ��� �����ؾ� �� ���Դϴ�. MouseMove �̺�Ʈ�� ��� ���� ��� ��Ʈ���� MouseMove �̺�Ʈ�� �����ؾ� �մϴ�.
'2)  X�� Y ��ǥ�� �����ϴ� MouseMove ���� ��� �̺�Ʈ�� ��ǥ ��ȯ�� �߰��մϴ�.
'3)  MoursePointer�� Borderstyle ���� ���Ÿ� ������ �Ӽ��� ��� �Ӽ��� ������ ������ �ش� ���� �̸����� �����Ͽ� ���� ��Ұ� �Ӽ� â�� ��Ÿ���� �մϴ�. �� ��쿡�� MousePointerConstants �� BorderStyleConstants�Դϴ�.
'4)  �ڽ��� ������ �Ӽ��� �ʿ��� ����� ���� ���ſ� �ڵ带 �߰��ؼ� ����� ���� ���Ÿ� ��ȿȭ�մϴ�.
'5)  ReadProperties �̺�Ʈ�� ���� ��⸦ �߰��Ͽ� .frm ���Ͽ� �������� �����ؼ� ���� �Ͱ��� �߸��� ���̳� ������ �������κ��� ��ȣ�մϴ�. ��� �Ӽ��� �̷� ������ �߻��ϸ� �ڵ带 �߰��ؼ� �⺻ ���������� ��ȯ�մϴ�. (�¶��� �������� "��Ʈ�� �Ӽ� ����"�� "������ ��� �Ǵ� ���� ��� ���� �Ӽ� �ۼ�"�� �����Ͻʽÿ�.)
'6)  ���� ��� ��Ʈ���� �ִٸ� UserControl_Resize�� �ڵ带 �߰��ؼ� ��Ʈ���� ũ�⸦ ������ �� ���� ��� ��Ʈ���� ũ�⵵ �����մϴ�.
'7)  Enabled �Ӽ��� ���� ���ν��� ID�� �����Ͻʽÿ�. �׷��� ��Ʈ���� ��밡���� ���� ��� �Ұ����� ��쿡 �ٸ� ActiveX ��Ʈ�Ѱ� �����ϰ� �۵��մϴ�.
'8)  �� ������� ��Ʈ���� �Ӽ��� ����� �̸��� ���� ���� ��� ��Ʈ��( UserControl)�� �Ӽ��� �����մϴ�. ��쿡 ���� �� �Ӽ��� �̸��� �ٸ� �Ӽ��� �����Ϸ��� ��쵵 ���� �� �ս��ϴ�. ���� ��� ShapeLabel�� ���� ��� Shape ��Ʈ���� FillColor�� BackColor�� �����մϴ�. �̷��� ������Ϸ��� �������� �ؾ� �մϴ�.
'9)  ��Ʈ���� ũ�⿡ ������ �ִ� �Ӽ���(AutoSize �Ӽ��� ���� ��Ʈ�ѿ����� �۲� ũ��)�� Property Let���� ũ�� ���� �ڵ带 ȣ���ؾ� �մϴ�.
'10) ����ڰ� ���� ��Ʈ���̸� UserControl�� Paint �̺�Ʈ�� �ڵ带 �߰��Ͽ� ��Ʈ���� ����� �׸��ϴ�. (�¶��� �������� "����� �׸��� ��Ʈ��"�� "��Ʈ�ѿ� ���� ��Ŀ�� ��� ���"�� �����Ͻʽÿ�.)
'11)  ��Ʈ���� �ϳ� �̻��� �Ӽ��� �����Ϳ� ���ε��Ǹ� �¶��� �������� "��Ʈ���� ������ �������� ���ε�"�� �����Ͻʽÿ�.
'12)  ��Ʈ�ѿ� �̿��� ����� �߰��մϴ�. �¶��� �������� "Visual Basic ActiveX ��Ʈ�� Ư¡'�� �ִ� �׸���� �ڼ��� ������ ������ �� ���Դϴ�.
'
'(�̷��� �۾� �׸� ���� ������ ������ CtlPlus.vbg ���� ���� ���α׷��� �����Ͻʽÿ�.)
'
'�����縦 �ٽ� �����ϰ� UserControl�� �����ؼ� ��Ʈ���� ������ �� �ֽ��ϴ�.
'
'������� UserControlsUse�� ���� �Ӽ� �������� ������� �Ӽ� ������ �����縦 ����մϴ�.
'
'ActiveX ��Ʈ�� ������ �׽�Ʈ�� ���� �ڼ��� ������ 4�� "ActiveX ��Ʈ�� �ۼ�"�� 9�� "ActiveX ��Ʈ�� ����"�� �����Ͻʽÿ�.
'
'6�� "���� ��� �������� �Ϲ� ����"�� 7�� " ActiveX ���� ����� �����, �׽�Ʈ, ����"���� ������ �߰� ������ ���� �� �ֽ��ϴ�.'�⺻ �Ӽ� ��:

'Const m_def_FormatMask = "##0.##"
'Const m_def_MinVal = -999.99
'Const m_def_MaxVal = 999.99
'�Ӽ� ����:
Option Explicit
Option Base 0

Dim m_AllowNull As Boolean


Dim m_FormatMask As String
Dim m_MinVal As Double
Dim m_MaxVal As Double
Dim focusFlag As Boolean

'Dim m_FormatMask As String
'Dim m_MinVal As Double
'Dim m_MaxVal As Double
'�̺�Ʈ ����:
Event Validate(Cancel As Boolean) 'MappingInfo=Text1,Text1,-1,Validate
Attribute Validate.VB_Description = "��Ʈ���� ��ȿ�� �˻縦 ����Ų ��Ʈ�ѿ� ���� ��Ŀ���� ���� �� �߻��մϴ�."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=Text1,Text1,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "OLE ����/���� �۾��� ���� �Ǵ� �ڵ����� ���۵� �� �߻��մϴ�."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=Text1,Text1,-1,OLESetData
Attribute OLESetData.VB_Description = "OLEDragStart �̺�Ʈ �� DataObject �� �־����� ���� �����͸� ���� ����� ��û�� �� OLE ����/���� ���� ��Ʈ�ѿ��� �߻��մϴ�."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=Text1,Text1,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "���콺 Ŀ���� ��ȭ�� �ʿ��� �� OLE ����/���� �۾��� ���� ��Ʈ�ѿ��� �߻��մϴ�."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=Text1,Text1,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "OLEDropMode �Ӽ��� �������� ������ ������ ��� OLE ����/���� �۾� �� ���콺�� ��Ʈ�� ���� �̵��ϸ� �߻��մϴ�."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "OLE ����/���� �۾��� ���� �����͸� ��Ʈ�ѿ� ���� �� �׸��� OLEDropMode�� �������� ������ �� �߻��մϴ�."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=Text1,Text1,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "���� �Ǵ� �ڵ� ����/���Ⱑ �Ϸ�ǰų� ��ҵ� �� OLE ����/���� ���� ��Ʈ�ѿ��� �߻��մϴ�."
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Attribute Change.VB_Description = "��Ʈ���� ������ ����� �� �߻��մϴ�."
Event Click() 'MappingInfo=Text1,Text1,-1,Click
Attribute Click.VB_Description = "��ü���� ���콺 ���߸� �����ٰ� ���� �� �߻��մϴ�."
Event DblClick() 'MappingInfo=Text1,Text1,-1,DblClick
Attribute DblClick.VB_Description = "���콺 ���߸� ��ü���� ������ ���� �� �ٽ� ������ ������ �߻��մϴ�."
Event KeyDown(keycode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Attribute KeyDown.VB_Description = "��ü�� ��Ŀ���� ���� �� Ű�� ������ �߻��մϴ�."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Attribute KeyPress.VB_Description = "ANSIŰ�� ������ ������ ��� �߻��մϴ�."
Event KeyUp(keycode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp
Attribute KeyUp.VB_Description = "��ü�� ��Ŀ���� ���� �� Ű�� ������ �߻��մϴ�."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseDown
Attribute MouseDown.VB_Description = "��ü�� ��Ŀ���� ���� �� ���콺 ���߸� ������ �߻��մϴ�."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseMove
Attribute MouseMove.VB_Description = "���콺�� ������ ��� �߻��մϴ�."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseUp
Attribute MouseUp.VB_Description = "��ü�� ��Ŀ���� ���� �� ���콺 ���߸� ������ �߻��մϴ�."
'�⺻ �Ӽ� ��:
Const m_def_AllowNull = False

Const m_def_FormatMask = "#,##0"
Const m_def_MinVal = -9999999999#
Const m_def_MaxVal = 9999999999#

Dim onReadProperte As Boolean




'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "��ü�� �ؽ�Ʈ�� �׷����� ǥ���ϱ� ���� ���Ǵ� ������ ��ȯ�ϰų� �����մϴ�."
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "��ü���� �ؽ�Ʈ�� �׷����� ǥ���ϴ� ������� ��ȯ�ϰų� �����մϴ�."
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����ڰ� ���� �̺�Ʈ�� ���� ��ü�� ������ �� �ִ����� ���θ� �����ϴ� ���� ��ȯ�ϰų� �����մϴ�."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font ��ü�� ��ȯ�մϴ�."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property
'
''���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
''MappingInfo=UserControl,UserControl,-1,BackStyle
'Public Property Get BackStyle() As Integer
'    BackStyle = UserControl.BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    UserControl.BackStyle() = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "��ü �׵θ� ������ ��ȯ�ϰų� �����մϴ�."
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Text1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "��ü�� ������ �ٽ� �׸��� �մϴ�."
    Text1.Refresh
End Sub

Private Sub Text1_Click()
    RaiseEvent Click
End Sub

Private Sub Text1_DblClick()
'focusFlag = True
'If focusFlag Then
'Else
    focusFlag = True
    selall
'End If
'    RaiseEvent DblClick
End Sub

Private Sub Text1_GotFocus()
selall
End Sub

Sub selall()
Dim ctl As Control
Debug.Print "selall"
Set ctl = Text1

If isNothing(ctl) Then
    Set ctl = Nothing
    Debug.Print "selall nothing"
    Exit Sub
End If
Debug.Print "seled"
ctl.SelStart = 0
ctl.SelLength = Len(ctl)

End Sub
Private Sub Text1_KeyDown(keycode As Integer, Shift As Integer)
    RaiseEvent KeyDown(keycode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

focusFlag = True

If Len(m_FormatMask) > 0 Then
    If Not IsNumeric(Asc(KeyAscii)) Then
        Select Case KeyAscii
        Case Asc("-")
            If Not (Left(m_FormatMask, 1) = "-") Then
                KeyAscii = 0
            End If
        Case Asc(".")
            If InStr(m_FormatMask, ".") = 0 Or InStr(Text1.Text, ".") > 0 Then
                KeyAscii = 0
            Else
                '
            End If
        Case vbKeyBack
            
        Case vbKeyDelete
        Case Else
            KeyAscii = 0
        End Select
    End If
End If

    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(keycode As Integer, Shift As Integer)
    RaiseEvent KeyUp(keycode, Shift)
End Sub



Private Sub Text1_LostFocus()
If AllowNull = False And Trim(Text1.Text) = "" Then
    MsgBox "������ ������ �ʽ��ϴ�.", vbExclamation
    Text1.SetFocus
Else
    focusFlag = False
End If
End Sub

'Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub

'Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub
'
Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If focusFlag Then
Else
    focusFlag = True
    selall
End If
'    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
''
''���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
''MemberInfo=13,0,0,##0.##
'Public Property Get FormatMask() As String
'    FormatMask = m_FormatMask
'End Property
'
'Public Property Let FormatMask(ByVal New_FormatMask As String)
'    m_FormatMask = New_FormatMask
'    PropertyChanged "FormatMask"
'End Property
'
''���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
''MemberInfo=4,0,0,-999.99
'Public Property Get MinVal() As Double
'    MinVal = m_MinVal
'End Property
'
'Public Property Let MinVal(ByVal New_MinVal As Double)
'    m_MinVal = New_MinVal
'    PropertyChanged "MinVal"
'End Property
'
''���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
''MemberInfo=4,0,0,999.99
'Public Property Get MaxVal() As Double
'    MaxVal = m_MaxVal
'End Property
'
'Public Property Let MaxVal(ByVal New_MaxVal As Double)
'    m_MaxVal = New_MaxVal
'    PropertyChanged "MaxVal"
'End Property

'����� ���� ��Ʈ�ѿ� ���� �Ӽ��� �ʱ�ȭ�մϴ�.
Private Sub UserControl_InitProperties()
'    m_FormatMask = m_def_FormatMask
'    m_MinVal = m_def_MinVal
'    m_MaxVal = m_def_MaxVal
    m_FormatMask = m_def_FormatMask
    m_MinVal = m_def_MinVal
    m_MaxVal = m_def_MaxVal
    m_AllowNull = m_def_AllowNull
End Sub

'����ҿ��� �Ӽ����� �ε��մϴ�.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

onReadProperte = True

    Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
'    m_FormatMask = PropBag.ReadProperty("FormatMask", m_def_FormatMask)
'    m_MinVal = PropBag.ReadProperty("MinVal", m_def_MinVal)
'    m_MaxVal = PropBag.ReadProperty("MaxVal", m_def_MaxVal)
    m_FormatMask = PropBag.ReadProperty("FormatMask", m_def_FormatMask)
    m_MinVal = PropBag.ReadProperty("MinVal", m_def_MinVal)
    m_MaxVal = PropBag.ReadProperty("MaxVal", m_def_MaxVal)
    Text1.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    Text1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    Text1.SelText = PropBag.ReadProperty("SelText", "")
    Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text1.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    Text1.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Text1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Text1.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    Text1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    Text1.LinkTopic = PropBag.ReadProperty("LinkTopic", "")
    Text1.LinkTimeout = PropBag.ReadProperty("LinkTimeout", 50)
    Text1.LinkMode = PropBag.ReadProperty("LinkMode", 0)
    Text1.LinkItem = PropBag.ReadProperty("LinkItem", "")
    ''Text1.IMEMode = PropBag.ReadProperty("IMEMode", 0)''���������� ������ �߻���.
    Text1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Text1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Text1.FontSize = PropBag.ReadProperty("FontSize", 14)
    Text1.FontName = PropBag.ReadProperty("FontName", "����")
    Text1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Text1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Text1.DataMember = PropBag.ReadProperty("DataMember", "")
    Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
    Text1.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    Text1.Appearance = PropBag.ReadProperty("Appearance", 1)
    Text1.Alignment = PropBag.ReadProperty("Alignment", 1)
    m_AllowNull = PropBag.ReadProperty("AllowNull", m_def_AllowNull)
    Text1.Text = PropBag.ReadProperty("Text", "0")
onReadProperte = False
End Sub

Private Sub UserControl_Resize()

Text1.Move 0, 0, UserControl.Width, UserControl.Height

End Sub

'�Ӽ����� ����ҿ� ����մϴ�.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
'    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
'    Call PropBag.WriteProperty("FormatMask", m_FormatMask, m_def_FormatMask)
'    Call PropBag.WriteProperty("MinVal", m_MinVal, m_def_MinVal)
'    Call PropBag.WriteProperty("MaxVal", m_MaxVal, m_def_MaxVal)
    Call PropBag.WriteProperty("FormatMask", m_FormatMask, m_def_FormatMask)
    Call PropBag.WriteProperty("MinVal", m_MinVal, m_def_MinVal)
    Call PropBag.WriteProperty("MaxVal", m_MaxVal, m_def_MaxVal)
    Call PropBag.WriteProperty("WhatsThisHelpID", Text1.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("ToolTipText", Text1.ToolTipText, "")
    Call PropBag.WriteProperty("SelText", Text1.SelText, "")
    Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
    Call PropBag.WriteProperty("RightToLeft", Text1.RightToLeft, False)
    Call PropBag.WriteProperty("PasswordChar", Text1.PasswordChar, "")
    Call PropBag.WriteProperty("OLEDropMode", Text1.OLEDropMode, 0)
    Call PropBag.WriteProperty("OLEDragMode", Text1.OLEDragMode, 0)
    Call PropBag.WriteProperty("MousePointer", Text1.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
    Call PropBag.WriteProperty("Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("LinkTopic", Text1.LinkTopic, "")
    Call PropBag.WriteProperty("LinkTimeout", Text1.LinkTimeout, 50)
    Call PropBag.WriteProperty("LinkMode", Text1.LinkMode, 0)
    Call PropBag.WriteProperty("LinkItem", Text1.LinkItem, "")
    'Call PropBag.WriteProperty("IMEMode", Text1.IMEMode, 0)
    Call PropBag.WriteProperty("FontUnderline", Text1.FontUnderline, 0)
    Call PropBag.WriteProperty("FontStrikethru", Text1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontSize", Text1.FontSize, 14)
    Call PropBag.WriteProperty("FontName", Text1.FontName, "")
    Call PropBag.WriteProperty("FontItalic", Text1.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", Text1.FontBold, 0)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataMember", Text1.DataMember, "")
    Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
    Call PropBag.WriteProperty("CausesValidation", Text1.CausesValidation, True)
    Call PropBag.WriteProperty("Appearance", Text1.Appearance, 1)
    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 1)
    Call PropBag.WriteProperty("AllowNull", m_AllowNull, m_def_AllowNull)
    Call PropBag.WriteProperty("Text", Text1.Text, "0")
End Sub
'
Public Property Get FormatMask() As String
    FormatMask = m_FormatMask
End Property

Public Property Let FormatMask(ByVal New_FormatMask As String)
    m_FormatMask = New_FormatMask
    PropertyChanged "FormatMask"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=4,0,0,-999.99
Public Property Get MinVal() As Double
    MinVal = m_MinVal
End Property

Public Property Let MinVal(ByVal New_MinVal As Double)
    m_MinVal = New_MinVal
    PropertyChanged "MinVal"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=4,0,0,999.99
Public Property Get MaxVal() As Double
    MaxVal = m_MaxVal
End Property

Public Property Let MaxVal(ByVal New_MaxVal As Double)
    m_MaxVal = New_MaxVal
    PropertyChanged "MaxVal"
End Property


'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "��ü�� ���� ���ؽ�Ʈ ��ȣ�� ��ȯ�ϰų� �����մϴ�."
    WhatsThisHelpID = Text1.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    Text1.WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

Private Sub Text1_Validate(Cancel As Boolean)
If AllowNull = True Then
    If Trim(Text1) = "" Then
        MsgBox "���鹮�ڴ� ������ �ʽ��ϴ�.", vbExclamation
        Cancel = True
    End If
End If
    'RaiseEvent Validate(Cancel)
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "���콺�� ��Ʈ�ѿ� �Ͻ� �����Ǿ� ���� �� ǥ�õǴ� �ؽ�Ʈ�� ��ȯ�ϰų� �����մϴ�."
    ToolTipText = Text1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Text1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "���� ���õ� �ؽ�Ʈ�� �����ϴ� ���ڿ��� ��ȯ�ϰų� �����մϴ�."
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Text1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "���õ� �ؽ�Ʈ�� �������� ��ȯ�ϰų� �����մϴ�."
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "���õ� ������ ������ ��ȯ�ϰų� �����մϴ�."
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "��ü�� ����/���� ��ũ�� ���븦 ���������� ���θ� ��Ÿ���� ���� ��ȯ�ϰų� �����մϴ�."
    ScrollBars = Text1.ScrollBars
End Property

'"Scale" ������ ������ ������ VBA�� ������̹Ƿ�
'�ʿ��մϴ�.
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=5
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "�ؽ�Ʈ ǥ�� ������ �����ϰ� ����� �ý��ۿ��� �ð��� ����� �����մϴ�."
    RightToLeft = Text1.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    Text1.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "����ڰ� �Է��� ���� �Ǵ� �ڸ� ǥ���� ���ڰ� ��Ʈ�ѿ� ǥ�õǾ����� �����ϴ� ���� ��ȯ�ϰų� �����մϴ�. "
    PasswordChar = Text1.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    Text1.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

Private Sub Text1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Text1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "�� ��ü�� OLE ���� ������� ����� �� �ִ��� �׸��� �̰��� �ڵ����� Ȥ�� ���α׷��� �����Ͽ� �Ͼ������ ���θ� ��ȯ�ϰų� �����մϴ�."
    OLEDropMode = Text1.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    Text1.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,OLEDragMode
Public Property Get OLEDragMode() As Integer
Attribute OLEDragMode.VB_Description = "�� ��Ʈ���� OLE ����/���� �������� ����ɼ� �ִ��� �׸��� �� ���μ����� �ڵ����� Ȥ�� ���α׷��� �����Ͽ� ���۵ɼ� �ִ����� ���θ� ��ȯ�ϰų� �����մϴ�."
    OLEDragMode = Text1.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
    Text1.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "�־��� ��Ʈ���� �������� ����Ͽ� OLE ����/���� �̺�Ʈ�� �����մϴ�."
    Text1.OLEDrag
End Sub

Private Sub Text1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "��Ʈ���� ���� ���� �ؽ�Ʈ�� �޾Ƶ��� �� �ִ��� ���θ� �����ϴ� ���� ��ȯ�ϰų� �����մϴ�."
    MultiLine = Text1.MultiLine
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "��ü�� ���콺 �����Ͱ� ���� ��� ǥ�õǴ� ���콺 �������� ������ ��ȯ�ϰų� �����մϴ�."
    MousePointer = Text1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Text1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "����� ���� ���콺 �������� �����մϴ�."
    Set MouseIcon = Text1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Text1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "��Ʈ�ѿ� �� �� �ִ� ������ �ִ���� ��ȯ�ϰų� �����մϴ�."
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "��Ʈ���� ���� ���� ���θ� �����մϴ�."
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "���� ���� ���α׷��� ��� ��Ʈ���� �׸��� ��ȯ�ϰų� �����մϴ�."
    LinkTopic = Text1.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
    Text1.LinkTopic() = New_LinkTopic
    PropertyChanged "LinkTopic"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "��Ʈ���� DDE �޽����� ������ ��ٸ��� �ð��� ��ȯ�ϰų� �����մϴ�. "
    LinkTimeout = Text1.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
    Text1.LinkTimeout() = New_LinkTimeout
    PropertyChanged "LinkTimeout"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkSend
Public Sub LinkSend()
Attribute LinkSend.VB_Description = "PictureBox�� ������ DDE ��ȭ�� ��� ���� ���α׷����� �����մϴ�."
    Text1.LinkSend
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkRequest
Public Sub LinkRequest()
Attribute LinkRequest.VB_Description = "���� DDE ���� ���α׷��� ��û�Ͽ� Label�̳� PictureBox, Textbox ��Ʈ�� ������ ������Ʈ�ϰ� �մϴ�."
    Text1.LinkRequest
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkPoke
Public Sub LinkPoke()
Attribute LinkPoke.VB_Description = "DDE ��ȭ�� ���� ���� ���α׷����� Label, PictureBox, TextBox�� ������ �����մϴ�."
    Text1.LinkPoke
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkMode
Public Property Get LinkMode() As Integer
Attribute LinkMode.VB_Description = "DDE ��ȭ�� ���� ���� ���¸� ��ȯ�ϰų� �����ϰ� ������ Ȱ��ȭ��ŵ�ϴ�."
    LinkMode = Text1.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As Integer)
    Text1.LinkMode() = New_LinkMode
    PropertyChanged "LinkMode"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "�ٸ� ���� ���α׷��� �Բ� DDE ��ȭ�� ��� ��Ʈ���� ����� �����͸� ��ȯ�ϰų� �����մϴ�."
    LinkItem = Text1.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
    Text1.LinkItem() = New_LinkItem
    PropertyChanged "LinkItem"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,LinkExecute
Public Sub LinkExecute(ByVal Command As String)
Attribute LinkExecute.VB_Description = "DDE ��ȭ�� ���� ���� ���α׷��� ��� ���ڿ��� �����մϴ�."
    Text1.LinkExecute Command
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,IMEMode
Public Property Get IMEMode() As Integer
Attribute IMEMode.VB_Description = "�Է� �޼��� �������� ���� � ��带 ��ȯ�ϰų� �����մϴ�."
    ''IMEMode = Text1.IMEMode''���������� ������ �߻���.
End Property

Public Property Let IMEMode(ByVal New_IMEMode As Integer)
    'Text1.IMEMode() = New_IMEMode''���������� ������ �߻���.
    PropertyChanged "IMEMode"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Microsoft Windows�� �ڵ��� ��ü�� â���� ��ȯ�մϴ�."
    hwnd = UserControl.hwnd
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "��Ʈ���� ��Ŀ���� �Ҿ��� ��� Masked edit ��Ʈ���� ������ ���������� ���θ� �����մϴ�."
    HideSelection = Text1.HideSelection
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "���� �۲� ������ ��ȯ�ϰų� �����մϴ�."
    FontUnderline = Text1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Text1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "��Ҽ� �۲� ������ ��ȯ�ϰų� �����մϴ�."
    FontStrikethru = Text1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Text1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "�־��� �ܰ��� �� �࿡ ��Ÿ���� �۲� ũ�⸦ ����Ʈ ������ �����մϴ�."
    FontSize = Text1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Text1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "�־��� �ܰ��� �� �࿡ ��Ÿ���� �۲��� �̸��� �����մϴ�."
    FontName = Text1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Text1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "����� �۲� ������ ��ȯ�ϰų� �����մϴ�."
    FontItalic = Text1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Text1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "���� �۲� ������ ��ȯ�ϰų� �����մϴ�."
    FontBold = Text1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Text1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,DataSource
Public Property Get DataSource() As Variant ' DataSource
Attribute DataSource.VB_Description = "Data ��Ʈ���� �����ϴ� ���� �����մϴ�. �� Data ��Ʈ���� ���Ͽ� ���� ��Ʈ���� �����ͺ��̽��� ���ε��˴ϴ�."
    Set DataSource = Text1.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As Variant)
    Set Text1.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "������ ���ῡ ���� DataMember�� �����ϴ� ���� ��ȯ�ϰų� �����մϴ�."
    DataMember = Text1.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    Text1.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,DataFormat
Public Property Get DataFormat() As Variant  ' OLE_OPTEXCLUSIVE ' IStdDataFormatDisp
Attribute DataFormat.VB_Description = "�� ���� ����� ���ε� ���� �Ӽ��� ���� ����ϱ� ���� DataFormat ��ü�� ��ȯ�մϴ�."
    Set DataFormat = Text1.DataFormat
End Property

Public Property Set DataFormat(ByVal New_DataFormat As Variant) ' IStdDataFormatDisp)
    Set Text1.DataFormat = New_DataFormat
    PropertyChanged "DataFormat"
End Property

Private Sub Text1_Change()
Static onChange As Boolean

If onReadProperte Then Exit Sub

Dim prePos As Long
Dim preLen As Long
Dim ctl As Variant

Dim lastPos As Long
Dim lastLen As Long

 

err.Clear
On Error Resume Next
Set ctl = Text1

If onChange Then Exit Sub
RaiseEvent Change
onChange = True


If Len(m_FormatMask) > 0 And IsNumeric(Text1.Text) Then
    On Error Resume Next
    prePos = ctl.SelStart
    On Error GoTo 0
    preLen = Len(ctl)
    
    Text1.Text = Format(Text1, m_FormatMask)
    
    If m_MaxVal < CDbl(Text1.Text) Then
        MsgBox "�ִ밪 ���� ����.(�ִ밪: " & m_MaxVal & ")", vbExclamation, "����"
        onChange = False
        Text1.Text = m_MaxVal
        
    ElseIf m_MinVal > CDbl(Text1.Text) Then
        MsgBox "�ּҰ� ���� ����.(�ּҰ�: " & m_MinVal & ")", vbExclamation, "����"
        onChange = False
        Text1.Text = m_MinVal
    Else
        lastLen = Len(Text1)
        If (lastLen - preLen) > 0 Then
            ctl.SelStart = prePos + (lastLen - preLen)
            ctl.SelLength = 0
        Else
            ctl.SelStart = prePos
            ctl.SelLength = 0
        End If
        
    End If
End If
onChange = False
   
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,CausesValidation
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "��Ŀ���� ���� ��Ʈ�ѿ��� ��ȿ�� �˻縦 �ϴ��� ���θ� ��ȯ�ϰų� �����մϴ�."
    CausesValidation = Text1.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    Text1.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "���� ��忡�� ��ü�� 3D ȿ���� �׷��������� ���θ� ��ȯ�ϰų� �����մϴ�."
    Appearance = Text1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    Text1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "CheckBox�� OptionButton �Ǵ�  ��Ʈ�� �ؽ�Ʈ�� ������ ��ȯ�ϰų� �����մϴ�."
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Function isNothing(ctl As Variant) As Boolean
On Error Resume Next
err.Clear
If Len(ctl) > 0 Then
End If


If err.Number = 91 Then
    isNothing = True
Else
    isNothing = False
End If
On Error GoTo 0

End Function
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0,0,0,False
Public Property Get AllowNull() As Boolean
Attribute AllowNull.VB_Description = "NULL�� ��������� ����"
    AllowNull = m_AllowNull

End Property

Public Property Let AllowNull(ByVal New_AllowNull As Boolean)
    m_AllowNull = New_AllowNull
    If m_AllowNull Then
        CausesValidation = True
    End If
    PropertyChanged "AllowNull"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "��Ʈ�ѿ� ���Ե� �ؽ�Ʈ�� ��ȯ�ϰų� �����մϴ�."
Attribute Text.VB_UserMemId = 0
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

