VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colAddSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsAddSec"
Attribute VB_Ext_KEY = "Member0" ,"clsAddSec"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

'�÷����� ������ ���� ����
Private mCol As Collection

Public Function Add(sigan As String, gijunilja As String, subj As String, userid As String, Optional ByVal sKey As String) As Boolean
'�� ��ü�� ����ϴ�.
    Dim objNewMember As clsAddSec
    Set objNewMember = New clsAddSec


    '�޼���� ���޵Ǵ� �Ӽ��� �����մϴ�.
    objNewMember.sigan = sigan
    objNewMember.gijunilja = gijunilja
    objNewMember.subj = subj
    objNewMember.userid = userid
   
    If Len(sKey) = 0 Then
        sKey = userid & "_" & subj & "_" & gijunilja & "_" & sigan
    End If
    
On Error GoTo ErrTrap
    mCol.Add objNewMember, sKey
    mCol(sKey).totalsec = 1

Exit Function

ErrTrap:
'    Dim tmr As Double
'    tmr = Timer
    Static preTmr As Double
    
    'Timer()�� ���������� �ʸ� ��ġ�� �����ִ� �Լ��̴�.
    
    If Abs(preTmr - Timer) > 60 Then
        preTmr = Timer - 1
    
    End If
    mCol(sKey).totalsec = mCol(sKey).totalsec + Round15(Timer - preTmr)
    Debug.Assert mCol(sKey).totalsec > 0
    preTmr = Timer
    If mCol(sKey).totalsec > 60 * 5 Then '5�д����� �����ϰԵ�...
        Add = True
    Else
        Add = False
    End If


End Function

Public Property Get Item(vntIndexKey As Variant) As clsAddSec
Attribute Item.VB_UserMemId = 0
    '�÷��ǿ� �ִ� ��Ҹ� ������ �� ���˴ϴ�.
    'vntIndexKey���� �÷��ǿ� ���� �ε����� Ű�� ���������
    '�� ������ Variant�� ����˴ϴ�.
    '����: Set foo = x.Item(xyz) �Ǵ� Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get count() As Long
    '�÷��ǿ� �ִ� ����� ���� �˻��� �� ���˴ϴ�.
    '����: Debug.Print x.Count
    count = mCol.count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '�÷��ǿ��� ��Ҹ� ������ �� ���˴ϴ�.
    'vntIndexKey���� �ε����� Ű�� ��������� �� ������ Variant��
    '����˴ϴ�.
    '����: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '�� �Ӽ��� ����ϸ� For...Each ������ ����Ͽ�
    '�� �÷����� ������ �� �ֽ��ϴ�.
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '�� Ŭ������ �ۼ��� �� �÷����� �ۼ��մϴ�.
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '�� Ŭ������ ������ �� �÷����� �����մϴ�.
    Set mCol = Nothing
End Sub
