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

'컬렉션을 포함할 지역 변수
Private mCol As Collection

Public Function Add(sigan As String, gijunilja As String, subj As String, userid As String, Optional ByVal sKey As String) As Boolean
'새 개체를 만듭니다.
    Dim objNewMember As clsAddSec
    Set objNewMember = New clsAddSec


    '메서드로 전달되는 속성을 설정합니다.
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
    
    'Timer()는 자정부터의 초를 수치로 보여주는 함수이다.
    
    If Abs(preTmr - Timer) > 60 Then
        preTmr = Timer - 1
    
    End If
    mCol(sKey).totalsec = mCol(sKey).totalsec + Round15(Timer - preTmr)
    Debug.Assert mCol(sKey).totalsec > 0
    preTmr = Timer
    If mCol(sKey).totalsec > 60 * 5 Then '5분단위로 저장하게됨...
        Add = True
    Else
        Add = False
    End If


End Function

Public Property Get Item(vntIndexKey As Variant) As clsAddSec
Attribute Item.VB_UserMemId = 0
    '컬렉션에 있는 요소를 참조할 때 사용됩니다.
    'vntIndexKey에는 컬렉션에 대한 인덱스나 키가 들어있으며
    '이 때문에 Variant로 선언됩니다.
    '구문: Set foo = x.Item(xyz) 또는 Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get count() As Long
    '컬렉션에 있는 요소의 수를 검색할 때 사용됩니다.
    '구문: Debug.Print x.Count
    count = mCol.count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '컬렉션에서 요소를 제거할 때 사용됩니다.
    'vntIndexKey에는 인덱스나 키가 들어있으며 이 때문에 Variant로
    '선언됩니다.
    '구문: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '이 속성을 사용하면 For...Each 구문을 사용하여
    '이 컬렉션을 열거할 수 있습니다.
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '이 클래스가 작성될 때 컬렉션을 작성합니다.
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '이 클래스를 종료할 때 컬렉션을 제거합니다.
    Set mCol = Nothing
End Sub

