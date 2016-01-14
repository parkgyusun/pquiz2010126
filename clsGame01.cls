VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGame01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'속성값을 가질 지역 변수
Private mvarcols As Integer '지역 복사
Private mvarrows As Integer '지역 복사
Private mvarquestion As String '지역 복사
Private mvarquestion_queue1 As String '지역 복사
Private mvaranswer As String '지역 복사
Private mvaranswer_queue1 As String '지역 복사
Private mvarquestion_queue2 As String '지역 복사
Private mvarquestion_queue3 As String '지역 복사
Private mvarquestion_queue4 As String '지역 복사
Private mvaranswer_queue2 As String '지역 복사
Private mvaranswer_queue3 As String '지역 복사
Private mvaranswer_queue4 As String '지역 복사
Private mvarobjQuiz As clsQuiz '지역 복사
Private mvarobjQuiz1 As clsQuiz '지역 복사
Private mvarobjQuiz2 As clsQuiz '지역 복사
Private mvarobjQuiz3 As clsQuiz '지역 복사
Private mvarobjQuiz4 As clsQuiz '지역 복사
Private mvarstrNews As Variant '지역 복사
Private mvarspendTime As Double '지역 복사
Private mvarposX As Long '지역 복사
Private mvarposY As Long '지역 복사
Private mvartotalLife As Long '지역 복사
Private mvarscore As Integer '지역 복사
Private mvartotalSecondsStudy As Long '지역 복사
Private mvaritemIndex As Long '지역 복사
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent ended[(arg1, arg2, ... , argn)]
Public Event ended()
Private mvarremainLoad As Long '지역 복사
Private mvarcorrectAnswerNumber As Integer '지역 복사
Private mvaritemIndexName As String '지역 복사
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent doEnding[(arg1, arg2, ... , argn)]
Public Event doEnding()
Private mvarcorrectPosX As Long '지역 복사
Private mvarcorrectPosY As Long '지역 복사
Private mvardeadPosY As Long '지역 복사
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent pouse[(arg1, arg2, ... , argn)]
Public Event pouse()
Private mvarstrQuery As String '지역 복사
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent fireQuery[(arg1, arg2, ... , argn)]
Public Event fireQuery()
Private mvarmutex As Boolean '지역 복사
'속성값을 가질 지역 변수
Private mvarbSoundOn_one As Boolean '지역 복사
Private mvarbSoundOn_every As Boolean '지역 복사
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent ttsOn[(arg1, arg2, ... , argn)]
Public Event ttsOn()
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent ttsOff[(arg1, arg2, ... , argn)]
Public Event ttsOff()
'이 이벤트를 발생시키려면 다음과 같은 구문을 가지는 RaiseEvent를 사용하십시오.
'RaiseEvent ready[(arg1, arg2, ... , argn)]
Public Event ready()



Public Property Let bSoundOn_every(ByVal vData As Boolean)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.bSoundOn_every = 5
    mvarbSoundOn_every = vData
End Property


Public Property Get bSoundOn_every() As Boolean
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.bSoundOn_every
    bSoundOn_every = mvarbSoundOn_every
End Property



Public Property Let bSoundOn_one(ByVal vData As Boolean)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.bSoundOn_one = 5
    mvarbSoundOn_one = vData
End Property


Public Property Get bSoundOn_one() As Boolean
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.bSoundOn_one
    bSoundOn_one = mvarbSoundOn_one
End Property



Public Sub isX()
End Sub

Public Sub isO()
End Sub

Public Sub keyPressed(keycode As Integer)
End Sub

Public Property Let mutex(ByVal vData As Boolean)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.mutex = 5
    mvarmutex = vData
End Property


Public Property Get mutex() As Boolean
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.mutex
    mutex = mvarmutex
End Property



Public Sub removeItem(idx As Integer)
End Sub

Public Function isFull() As Boolean
End Function

Public Sub swap()
End Sub

Public Sub doStart(doll_index As Long, doll_Name As String, beforeStudiedSeconds As Long, aimCount As Long)
End Sub

Public Sub oneSecondsAdded()
End Sub


Public Property Let strQuery(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.strQuery = 5
    mvarstrQuery = vData
End Property


Public Property Get strQuery() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.strQuery
    strQuery = mvarstrQuery
End Property



Public Sub makeQuery()
End Sub

Public Sub resumeGame()
End Sub


Public Sub sendToDB(spendSeconds As Double, objQuiz As clsQuiz, doll_index As Long)
End Sub

Public Sub doCheck()
End Sub

Public Property Let deadPosY(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.deadPosY = 5
    mvardeadPosY = vData
End Property


Public Property Get deadPosY() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.deadPosY
    deadPosY = mvardeadPosY
End Property



Public Property Let correctPosY(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.correctPosY = 5
    mvarcorrectPosY = vData
End Property


Public Property Get correctPosY() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.correctPosY
    correctPosY = mvarcorrectPosY
End Property



Public Property Let correctPosX(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.correctPosX = 5
    mvarcorrectPosX = vData
End Property


Public Property Get correctPosX() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.correctPosX
    correctPosX = mvarcorrectPosX
End Property




Public Property Let itemIndexName(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.itemIndexName = 5
    mvaritemIndexName = vData
End Property


Public Property Get itemIndexName() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.itemIndexName
    itemIndexName = mvaritemIndexName
End Property



Public Property Let correctAnswerNumber(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.correctAnswerNumber = 5
    mvarcorrectAnswerNumber = vData
End Property


Public Property Get correctAnswerNumber() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.correctAnswerNumber
    correctAnswerNumber = mvarcorrectAnswerNumber
End Property



Public Property Let remainLoad(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.remainLoad = 5
    mvarremainLoad = vData
End Property


Public Property Get remainLoad() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.remainLoad
    remainLoad = mvarremainLoad
End Property



Public Sub newGame(count As Long)
End Sub


Public Property Let itemIndex(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.itemIndex = 5
    mvaritemIndex = vData
End Property


Public Property Get itemIndex() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.itemIndex
    itemIndex = mvaritemIndex
End Property



Public Property Let totalSecondsStudy(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.totalSecondsStudy = 5
    mvartotalSecondsStudy = vData
End Property


Public Property Get totalSecondsStudy() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.totalSecondsStudy
    totalSecondsStudy = mvartotalSecondsStudy
End Property



Public Property Let score(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.score = 5
    mvarscore = vData
End Property


Public Property Get score() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.score
    score = mvarscore
End Property



Public Property Let totalLife(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.totalLife = 5
    mvartotalLife = vData
End Property


Public Property Get totalLife() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.totalLife
    totalLife = mvartotalLife
End Property



Public Property Let posY(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.posY = 5
    mvarposY = vData
End Property


Public Property Get posY() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.posY
    posY = mvarposY
End Property



Public Property Let posX(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.posX = 5
    mvarposX = vData
End Property


Public Property Get posX() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.posX
    posX = mvarposX
End Property



Public Property Let spendTime(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.spendTime = 5
    mvarspendTime = vData
End Property


Public Property Get spendTime() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.spendTime
    spendTime = mvarspendTime
End Property



Public Property Let strNews(ByVal vData As Variant)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.strNews = 5
    mvarstrNews = vData
End Property


Public Property Set strNews(ByVal vData As Variant)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.strNews = Form1
    Set mvarstrNews = vData
End Property


Public Property Get strNews() As Variant
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.strNews
    If IsObject(mvarstrNews) Then
        Set strNews = mvarstrNews
    Else
        strNews = mvarstrNews
    End If
End Property



Public Property Set objQuiz4(ByVal vData As clsQuiz)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.objQuiz4 = Form1
    Set mvarobjQuiz4 = vData
End Property


Public Property Get objQuiz4() As clsQuiz
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.objQuiz4
    Set objQuiz4 = mvarobjQuiz4
End Property



Public Property Set objQuiz3(ByVal vData As clsQuiz)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.objQuiz3 = Form1
    Set mvarobjQuiz3 = vData
End Property


Public Property Get objQuiz3() As clsQuiz
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.objQuiz3
    Set objQuiz3 = mvarobjQuiz3
End Property



Public Property Set objQuiz2(ByVal vData As clsQuiz)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.objQuiz2 = Form1
    Set mvarobjQuiz2 = vData
End Property


Public Property Get objQuiz2() As clsQuiz
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.objQuiz2
    Set objQuiz2 = mvarobjQuiz2
End Property



Public Property Set objQuiz1(ByVal vData As clsQuiz)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.objQuiz1 = Form1
    Set mvarobjQuiz1 = vData
End Property


Public Property Get objQuiz1() As clsQuiz
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.objQuiz1
    Set objQuiz1 = mvarobjQuiz1
End Property



Public Property Set objQuiz(ByVal vData As clsQuiz)
'Set 문의 왼쪽에서 속성에 개체를 할당할 때 사용됩니다.
'Syntax: Set x.objQuiz = Form1
    Set mvarobjQuiz = vData
End Property


Public Property Get objQuiz() As clsQuiz
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.objQuiz
    Set objQuiz = mvarobjQuiz
End Property



Public Sub addQuiz(question As String, answer As String)
End Sub

Public Sub keyUPUP()
End Sub

Public Sub moveUP()
End Sub

Public Sub moveRight()
End Sub

Public Sub moveLeft()
End Sub

Public Sub moveDown()
End Sub

Public Property Let answer_queue4(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.answer_queue4 = 5
    mvaranswer_queue4 = vData
End Property


Public Property Get answer_queue4() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.answer_queue4
    answer_queue4 = mvaranswer_queue4
End Property



Public Property Let answer_queue3(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.answer_queue3 = 5
    mvaranswer_queue3 = vData
End Property


Public Property Get answer_queue3() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.answer_queue3
    answer_queue3 = mvaranswer_queue3
End Property



Public Property Let answer_queue2(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.answer_queue2 = 5
    mvaranswer_queue2 = vData
End Property


Public Property Get answer_queue2() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.answer_queue2
    answer_queue2 = mvaranswer_queue2
End Property



Public Property Let question_queue4(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.question_queue4 = 5
    mvarquestion_queue4 = vData
End Property


Public Property Get question_queue4() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.question_queue4
    question_queue4 = mvarquestion_queue4
End Property



Public Property Let question_queue3(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.question_queue3 = 5
    mvarquestion_queue3 = vData
End Property


Public Property Get question_queue3() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.question_queue3
    question_queue3 = mvarquestion_queue3
End Property



Public Property Let question_queue2(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.question_queue2 = 5
    mvarquestion_queue2 = vData
End Property


Public Property Get question_queue2() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.question_queue2
    question_queue2 = mvarquestion_queue2
End Property



Public Property Let answer_queue1(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.answer_queue1 = 5
    mvaranswer_queue1 = vData
End Property


Public Property Get answer_queue1() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.answer_queue1
    answer_queue1 = mvaranswer_queue1
End Property



Public Property Let answer(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.answer = 5
    mvaranswer = vData
End Property


Public Property Get answer() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.answer
    answer = mvaranswer
End Property



Public Property Let question_queue1(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.question_queue1 = 5
    mvarquestion_queue1 = vData
End Property


Public Property Get question_queue1() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.question_queue1
    question_queue1 = mvarquestion_queue1
End Property



Public Property Let question(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.question = 5
    mvarquestion = vData
End Property


Public Property Get question() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.question
    question = mvarquestion
End Property



Public Property Let rows(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rows = 5
    mvarrows = vData
End Property


Public Property Get rows() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rows
    rows = mvarrows
End Property



Public Property Let cols(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.cols = 5
    mvarcols = vData
End Property


Public Property Get cols() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.cols
    cols = mvarcols
End Property



