VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

Public ccPocketnm As Collection

Public ccSubj As Collection

Public bCorrect As Boolean

Public bIncorrect As Boolean

Public bIncorrectRate As Boolean

'속성값을 가질 지역 변수
Private mvarCorrectCnt As Integer '지역 복사
Private mvarIncorrectCnt As Integer '지역 복사
Private mvarIncorrectRate As Integer '지역 복사
Public ccSubSubj As Collection
Public idx_ccSubSubj As Long
Private mvaritm_ccSubSubj_01 As Variant '지역 복사
Private mvaritm_ccSubSubj_02 As Variant '지역 복사
Private mvaritm_ccSubSubj_03 As Variant '지역 복사
Private mvaritm_ccSubSubj_04 As Variant '지역 복사
Private mvarselMethod As Integer '지역 복사
Public sInitial As String
Public bInitial As Boolean
Public sPocketNm As String
Public sPocketNmFull As String
Public sSubNm1Format As String
Public sSubNm2Format As String
Public sSubNm3Format As String
Public bSubNm1Format As Boolean
Public bSubNm2Format As Boolean
Public bSubNm3Format As Boolean
Public sSubChar1 As String
Public sSubChar2 As String
Public sSubChar3 As String
Public bQuizSwap As Boolean
Attribute bQuizSwap.VB_VarDescription = "문제섞기 여부"
Public bQuizSwap_all As Boolean
Attribute bQuizSwap_all.VB_VarDescription = "전체 문제를 섞기"
Public bQuizSwap_subj As Boolean
Attribute bQuizSwap_subj.VB_VarDescription = "과목범위에서 문제 섞기로 설정을 합니다."
'속성값을 가질 지역 변수
Private mvarbQuizSwap_chasu As Boolean '지역 복사
'속성값을 가질 지역 변수
Private mvarbReserveChk As Boolean '지역 복사
Public bReserveFromChk As Boolean
Public bReserveToChk As Boolean
Public sReserveFrom As String
Public sReserveTo As String
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent ReserveChkChanged[(arg1, arg2, ... , argn)]
Public Event ReserveChkChanged(ByVal bChk As Boolean)
Public nPerX As Double
Attribute nPerX.VB_VarDescription = "x일마다"
Public nPerY As Double
Attribute nPerY.VB_VarDescription = "Y문항씩"
Public nTotalCnt As Long
'속성값을 가질 지역 변수
Private mvarnDayTotal As Long '지역 복사
Public Property Let nDayTotal(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.nDayTotal = 5
    mvarnDayTotal = vData
End Property


Public Property Get nDayTotal() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.nDayTotal
    nDayTotal = mvarnDayTotal
End Property





Public Property Let bReserveChk(ByVal vData As Boolean)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.bReserveChk = 5
    mvarbReserveChk = vData
    RaiseEvent ReserveChkChanged(vData)
End Property


Public Property Get bReserveChk() As Boolean
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.bReserveChk
    bReserveChk = mvarbReserveChk
End Property



Public Property Let bQuizSwap_chasu(ByVal vData As Boolean)
Attribute bQuizSwap_chasu.VB_Description = "회차 내에서 문제 섞기를 합니다."
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.bQuizSwap_chasu = 5
    mvarbQuizSwap_chasu = vData
End Property


Public Property Get bQuizSwap_chasu() As Boolean
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.bQuizSwap_chasu
    bQuizSwap_chasu = mvarbQuizSwap_chasu
End Property


Public Property Let selMethod(ByVal vData As Integer)
Attribute selMethod.VB_Description = "선택방식\r\n1: 시작일, 종료일, 일마다\r\n2: 시작일, 종료일, 문항수\r\n3: 시작일, 일마다, 문항수"
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.selMethod = 5
    mvarselMethod = vData
End Property


Public Property Get selMethod() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.selMethod
    selMethod = mvarselMethod
End Property



Public Property Let itm_ccSubSubj_04(ByVal vData As Variant)
Attribute itm_ccSubSubj_04.VB_Description = "순서"
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.itm_ccSubSubj_04 = 5
    mvaritm_ccSubSubj_04 = vData
End Property


Public Property Set itm_ccSubSubj_04(ByVal vData As Variant)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.itm_ccSubSubj_04 = Form1
    Set mvaritm_ccSubSubj_04 = vData
End Property


Public Property Get itm_ccSubSubj_04() As Variant
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.itm_ccSubSubj_04
    If IsObject(mvaritm_ccSubSubj_04) Then
        Set itm_ccSubSubj_04 = mvaritm_ccSubSubj_04
    Else
        itm_ccSubSubj_04 = mvaritm_ccSubSubj_04
    End If
End Property



Public Property Let itm_ccSubSubj_03(ByVal vData As Variant)
Attribute itm_ccSubSubj_03.VB_Description = "문항수"
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.itm_ccSubSubj_03 = 5
    mvaritm_ccSubSubj_03 = vData
End Property


Public Property Set itm_ccSubSubj_03(ByVal vData As Variant)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.itm_ccSubSubj_03 = Form1
    Set mvaritm_ccSubSubj_03 = vData
End Property


Public Property Get itm_ccSubSubj_03() As Variant
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.itm_ccSubSubj_03
    If IsObject(mvaritm_ccSubSubj_03) Then
        Set itm_ccSubSubj_03 = mvaritm_ccSubSubj_03
    Else
        itm_ccSubSubj_03 = mvaritm_ccSubSubj_03
    End If
End Property



Public Property Let itm_ccSubSubj_02(ByVal vData As Variant)
Attribute itm_ccSubSubj_02.VB_Description = "과목명"
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.itm_ccSubSubj_02 = 5
    mvaritm_ccSubSubj_02 = vData
End Property


Public Property Set itm_ccSubSubj_02(ByVal vData As Variant)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.itm_ccSubSubj_02 = Form1
    Set mvaritm_ccSubSubj_02 = vData
End Property


Public Property Get itm_ccSubSubj_02() As Variant
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.itm_ccSubSubj_02
    If IsObject(mvaritm_ccSubSubj_02) Then
        Set itm_ccSubSubj_02 = mvaritm_ccSubSubj_02
    Else
        itm_ccSubSubj_02 = mvaritm_ccSubSubj_02
    End If
End Property



Public Property Let itm_ccSubSubj_01(ByVal vData As Variant)
Attribute itm_ccSubSubj_01.VB_Description = "선택여부"
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.itm_ccSubSubj_01 = 5
    mvaritm_ccSubSubj_01 = vData
End Property


Public Property Set itm_ccSubSubj_01(ByVal vData As Variant)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.itm_ccSubSubj_01 = Form1
    Set mvaritm_ccSubSubj_01 = vData
End Property


Public Property Get itm_ccSubSubj_01() As Variant
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.itm_ccSubSubj_01
    If IsObject(mvaritm_ccSubSubj_01) Then
        Set itm_ccSubSubj_01 = mvaritm_ccSubSubj_01
    Else
        itm_ccSubSubj_01 = mvaritm_ccSubSubj_01
    End If
End Property





Public Property Let IncorrectRate(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.IncorrectRate = 5
    mvarIncorrectRate = vData
End Property


Public Property Get IncorrectRate() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.IncorrectRate
    IncorrectRate = mvarIncorrectRate
End Property



Public Property Let IncorrectCnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.IncorrectCnt = 5
    mvarIncorrectCnt = vData
End Property


Public Property Get IncorrectCnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.IncorrectCnt
    IncorrectCnt = mvarIncorrectCnt
End Property



Public Property Let CorrectCnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CorrectCnt = 5
    mvarCorrectCnt = vData
End Property


Public Property Get CorrectCnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CorrectCnt
    CorrectCnt = mvarCorrectCnt
End Property



Public Function getPocketNm() As String
    
    If bInitial Then
        getPocketNm = sInitial
    End If
    
    getPocketNm = Mid(getPocketNm & " " & sPocketNm, 1, 30)

End Function


Public Function getSubPocketNm(Optional p_chasu As Variant, Optional dtfrom As Variant, Optional dtto As Variant) As String
    
If IsMissing(p_chasu) Then
    p_chasu = 1
    dtfrom = Now
    dtto = Now
End If
If Me.bSubNm1Format And Len(Me.sSubNm1Format) > 0 Then
    getSubPocketNm = Format(p_chasu, Me.sSubNm1Format)
End If

getSubPocketNm = getSubPocketNm & sSubChar1

If Me.bSubNm2Format And Len(Me.sSubNm2Format) > 0 Then
    getSubPocketNm = getSubPocketNm & Format(dtfrom, Me.sSubNm2Format)
End If

getSubPocketNm = getSubPocketNm & sSubChar2

If Me.bSubNm3Format And Len(Me.sSubNm3Format) > 0 Then
    getSubPocketNm = getSubPocketNm & Format(dtto, Me.sSubNm3Format)
End If

getSubPocketNm = Mid(getSubPocketNm & sSubChar3, 1, 30)

End Function


