VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPocketQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

'속성값을 가질 지역 변수
Private mvartotal As Integer '지역 복사
Private mvaro_cnt As Integer '지역 복사
Private mvarx_cnt As Integer '지역 복사
Private mvarpocketnm As String '지역 복사
Private mvarchk_cnt As Integer '지역 복사
'속성값을 가질 지역 변수



Public Property Let chk_cnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.chk_cnt = 5
    mvarchk_cnt = vData
End Property


Public Property Get chk_cnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.chk_cnt
    chk_cnt = mvarchk_cnt
End Property



Public Property Let pocketnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.pocketnm = 5
    mvarpocketnm = vData
End Property


Public Property Get pocketnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.pocketnm
    pocketnm = mvarpocketnm
End Property



Public Property Let x_cnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.x_cnt = 5
    mvarx_cnt = vData
End Property


Public Property Get x_cnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.x_cnt
    x_cnt = mvarx_cnt
End Property



Public Property Let o_cnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.o_cnt = 5
    mvaro_cnt = vData
End Property


Public Property Get o_cnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.o_cnt
    o_cnt = mvaro_cnt
End Property



Public Property Let total(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.total = 5
    mvartotal = vData
End Property


Public Property Get total() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.total
    total = mvartotal
End Property



