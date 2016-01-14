VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTg01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsTg01"
Attribute VB_Ext_KEY = "Member0" ,"clsTg01"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'�÷����� ������ ���� ����
Private mCol As Collection
Private sSql As String
Public Con As ADODB.Connection
Public LocalRs As ADODB.Recordset

Public Function Add(gijunilja As String, userid As String, subj As String, totalsec As Long, new_cnt As Long, review_cnt As Long, o_cnt As Long, x_cnt As Long, chk_cnt As Long, back_cnt As Long, next_cnt As Long, Optional sKey As String) As clsTg01
    '�� ��ü�� ����ϴ�.
    Dim objNewMember As clsTg01
    Set objNewMember = New clsTg01


    '�޼���� ���޵Ǵ� �Ӽ��� �����մϴ�.
    objNewMember.gijunilja = gijunilja
    objNewMember.userid = userid
    objNewMember.subj = subj
    objNewMember.totalsec = totalsec
    objNewMember.new_cnt = new_cnt
    objNewMember.review_cnt = review_cnt
    objNewMember.o_cnt = o_cnt
    objNewMember.x_cnt = x_cnt
    objNewMember.chk_cnt = chk_cnt
    objNewMember.back_cnt = back_cnt
    objNewMember.next_cnt = next_cnt
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '������� ��ü�� ��ȯ�մϴ�.
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsTg01
Attribute Item.VB_UserMemId = 0
    '�÷��ǿ� �ִ� ��Ҹ� ������ �� ���˴ϴ�.
    'vntIndexKey���� �÷��ǿ� ���� �ε����� Ű�� ���������
    '�� ������ Variant�� ����˴ϴ�.
    '����: Set foo = x.Item(xyz) �Ǵ� Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '�÷��ǿ� �ִ� ����� ���� �˻��� �� ���˴ϴ�.
    '����: Debug.Print x.Count
    Count = mCol.Count
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

Public Function read_tg01(con_state As Long, refCon As ADODB.Connection, gijunilja As String, strid As String, Optional subj As String) As clsTg01
If con_state = 1 Then
    Call prcRead(gijunilja, strid, subj)
Else
    Set Con = refCon
    Con_Open
    Call prcRead(gijunilja, strid, subj)
    Con_Close
    
End If
End Function

Sub prcRead(gijunilja As String, strid As String, Optional subj As String)

End Sub

Public Sub write_tg01(strid As String, subj As String)

End Sub

Public Sub addsec()

End Sub

Public Function Con_Open() As Boolean
On Error GoTo ERRTRAP
If Con.STATE = 0 Then
    Con.Open STRCON
    Con.CommandTimeout = 300
End If
Exit Function
ERRTRAP:
    Con_Open = True
    MsgBox Err.Description, vbCritical
    
End Function

Public Function Con_Close() As Boolean
On Error GoTo ERRTRAP
If Con.STATE = 1 Then
    Con.Close
End If

Exit Function

ERRTRAP:
    Con_Close = True
    MsgBox Err.Description, vbCritical
    
End Function
