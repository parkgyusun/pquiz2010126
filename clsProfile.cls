VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colIP_DerivativesMissing"
Attribute VB_Ext_KEY = "Member1" ,"colIP_derivativesOK"
Attribute VB_Ext_KEY = "Member2" ,"colIP_StockMissing"
Attribute VB_Ext_KEY = "Member3" ,"colIP_StockOk"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

Public Event saveProfile()
'�Ӽ����� ���� ���� ����
Private mvarfontsize As Single '���� ����
'�Ӽ����� ���� ���� ����
Private mvarbSetTimeOut As Boolean '���� ����
Private mvarbSetTimeOutStudy As Boolean '���� ����
Private mvarTimeOutSec As Integer '���� ����
Private mvarTimeOutSecStudy As Integer '���� ����
Private mvarCreateRepeatInterval As Integer '���� ����
Private mvarDirectTimeSec As Integer '���� ����
Private mvarbStreateOutCntCheck As Boolean '���� ����
Private mvarbAlarm1 As Boolean '���� ����
Private mvarCntOfAlarm1 As Integer '���� ����
Private mvarbAlarm2 As Boolean '���� ����
Private mvarbAlarm3 As Boolean '���� ����
Private mvarbAlarm4 As Boolean '���� ����
Private mvarbAlarm5 As Boolean '���� ����
Private mvarCntOfAlarm2 As Integer '���� ����
Private mvarCntOfAlarm3 As Integer '���� ����
Private mvarCntOfAlarm4 As Integer '���� ����
Private mvarCntOfAlarm5 As Integer '���� ����
Private mvarbChkAlarm1 As Boolean '���� ����
Private mvarbChkAlarm2 As Boolean '���� ����
Private mvarbChkAlarm3 As Boolean '���� ����
Private mvarbChkAlarm4 As Boolean '���� ����
Private mvarbChkAlarm5 As Boolean '���� ����
Private mvarCntOfStreatOutSetting As Integer '���� ����
Private mvarCntOfStreatOutNow As Integer '���� ����
'�Ӽ����� ���� ���� ����
Private mvarfontBold As Boolean '���� ����
'�Ӽ����� ���� ���� ����
Private mvarDelayTime1 As Integer '���� ����
Private mvarDelayTime2 As Integer '���� ����
'�Ӽ����� ���� ���� ����
Private mvarnansu As Long '���� ����
'�Ӽ����� ���� ���� ����
Private mvarcompactDB As Boolean '���� ����
Private mvarsetPoP1 As Boolean '���� ����
Private mvarsetPoP2 As Variant '���� ����
Private mvarShowWhiteSpot As Boolean '���� ����
Private mvarcheckWhenResize As Boolean '���� ����
'�Ӽ����� ���� ���� ����
Private mvarsetPoP3 As Boolean '���� ����
Private mvarLastUser As String
Private mvarFontName As String
'�Ӽ����� ���� ���� ����
Private mvarendSettingLogin As String '���� ����
Private mvarendSettingLoginDefaultMin As String '���� ����
Private mvarendSettingQuiz As String '���� ����
Private mvarendSettingQuizDefaultMin As String '���� ����
'�Ӽ����� ���� ���� ����
Private mvarendSettingSetMin As String '���� ����
Private mvarendSettingQuestion As Boolean '���� ����
Private mvarendSettingTimer As Boolean '���� ����
Private mvarbChkTTSuse As Boolean '���� ����
Private mvarbChkTTSuse1 As Boolean '���� ����
Private mvarbChkTTSuse2 As Boolean '���� ����
Private mvarbChkTTSuse3 As Boolean '���� ����
Private mvarbChkTTSuse4 As Boolean '���� ����
Private mvarPenColor As Long '���� ����
Public Property Let bChkTTSuse4(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkTTSuse4 = 5
    mvarbChkTTSuse4 = vData
End Property


Public Property Get bChkTTSuse4() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkTTSuse4
    bChkTTSuse4 = mvarbChkTTSuse4
End Property



Public Property Let bChkTTSuse3(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkTTSuse3 = 5
    mvarbChkTTSuse3 = vData
End Property


Public Property Get bChkTTSuse3() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkTTSuse3
    bChkTTSuse3 = mvarbChkTTSuse3
End Property



Public Property Let bChkTTSuse2(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkTTSuse2 = 5
    mvarbChkTTSuse2 = vData
End Property


Public Property Get bChkTTSuse2() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkTTSuse2
    bChkTTSuse2 = mvarbChkTTSuse2
End Property



Public Property Let bChkTTSuse1(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkTTSuse1 = 5
    mvarbChkTTSuse1 = vData
End Property


Public Property Get bChkTTSuse1() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkTTSuse1
    bChkTTSuse1 = mvarbChkTTSuse1
End Property



Public Property Let bChkTTSuse(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkTTSuse = 5
    mvarbChkTTSuse = vData
End Property


Public Property Get bChkTTSuse() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkTTSuse
    bChkTTSuse = mvarbChkTTSuse
End Property

Public Property Let endSettingTimer(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.endSettingTimer = 5
    mvarendSettingTimer = vData
End Property


Public Property Get endSettingTimer() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.endSettingTimer
    endSettingTimer = mvarendSettingTimer
End Property



Public Property Let endSettingQuestion(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.endSettingQuestion = 5
    mvarendSettingQuestion = vData
End Property


Public Property Get endSettingQuestion() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.endSettingQuestion
    endSettingQuestion = mvarendSettingQuestion
End Property



Public Property Let endSettingSetMin(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.endSettingSetMin = 5
    mvarendSettingSetMin = vData
End Property


Public Property Get endSettingSetMin() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.endSettingSetMin
    endSettingSetMin = mvarendSettingSetMin
End Property



Public Property Let endSettingQuizDefaultMin(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.endSettingQuizDefaultMin = 5
    mvarendSettingQuizDefaultMin = vData
End Property


Public Property Get endSettingQuizDefaultMin() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.endSettingQuizDefaultMin
    endSettingQuizDefaultMin = mvarendSettingQuizDefaultMin
End Property



Public Property Let endSettingQuiz(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.endSettingQuiz = 5
    mvarendSettingQuiz = vData
End Property


Public Property Get endSettingQuiz() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.endSettingQuiz
    endSettingQuiz = mvarendSettingQuiz
End Property



Public Property Let endSettingLoginDefaultMin(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.endSettingLoginDefaultMin = 5
    mvarendSettingLoginDefaultMin = vData
End Property


Public Property Get endSettingLoginDefaultMin() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.endSettingLoginDefaultMin
    endSettingLoginDefaultMin = mvarendSettingLoginDefaultMin
End Property



Public Property Let endSettingLogin(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.endSettingLogin = 5
    mvarendSettingLogin = vData
End Property


Public Property Get endSettingLogin() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.endSettingLogin
    endSettingLogin = mvarendSettingLogin
End Property




Public Property Let setPoP3(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.setPoP3 = 5
    mvarsetPoP3 = vData
End Property


Public Property Get setPoP3() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.setPoP3
    setPoP3 = mvarsetPoP3
End Property



Public Property Let checkWhenResize(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.checkWhenResize = 5
    mvarcheckWhenResize = vData
End Property


Public Property Get checkWhenResize() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.checkWhenResize
    checkWhenResize = mvarcheckWhenResize
End Property



Public Property Let ShowWhiteSpot(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.ShowWhiteSpot = 5
    mvarShowWhiteSpot = vData
End Property

Public Property Get LastUser() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.checkWhenResize
    LastUser = mvarLastUser
End Property



Public Property Let LastUser(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.ShowWhiteSpot = 5
    mvarLastUser = vData
End Property

Public Property Get ShowWhiteSpot() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.ShowWhiteSpot
    ShowWhiteSpot = mvarShowWhiteSpot
End Property

Public Property Let setPoP2(ByVal vData As Variant)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.setPoP2 = 5
    mvarsetPoP2 = vData
End Property


Public Property Set setPoP2(ByVal vData As Variant)
'Set ���� ���ʿ��� �Ӽ��� ��ü�� �Ҵ��� �� ���˴ϴ�.
'Syntax: Set x.setPoP2 = Form1
    Set mvarsetPoP2 = vData
End Property


Public Property Get setPoP2() As Variant
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.setPoP2
    If IsObject(mvarsetPoP2) Then
        Set setPoP2 = mvarsetPoP2
    Else
        setPoP2 = mvarsetPoP2
    End If
End Property



Public Property Let setPoP1(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.setPoP1 = 5
    mvarsetPoP1 = vData
End Property


Public Property Get setPoP1() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.setPoP1
    setPoP1 = mvarsetPoP1
End Property



Public Property Let compactDB(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.compactDB = 5
    mvarcompactDB = vData
End Property


Public Property Get compactDB() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.compactDB
    compactDB = mvarcompactDB
End Property



Public Property Let nansu(ByVal vData As Long)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.nansu = 5
    mvarnansu = vData
End Property


Public Property Get nansu() As Long
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.nansu
    nansu = mvarnansu
End Property



Public Property Let DelayTime2(ByVal vData As Integer)
Attribute DelayTime2.VB_Description = "���� ���� Ű�� �����Ͽ� ��������\r\n������ ��Ǭ��� �ش� �ð��̳����� \r\n""���� ������ Ǫ����"" ��� �˸��ϴ�."
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.DelayTime2 = 5
    mvarDelayTime2 = vData
End Property


Public Property Get DelayTime2() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.DelayTime2
    DelayTime2 = mvarDelayTime2
End Property



Public Property Let DelayTime1(ByVal vData As Integer)
Attribute DelayTime1.VB_Description = "Sleep �ð��� mSec�� �����մϴ�."
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.DelayTime1 = 5
    mvarDelayTime1 = vData
End Property


Public Property Get DelayTime1() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.DelayTime1
    DelayTime1 = mvarDelayTime1
End Property



Public Property Let FontBold(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.fontBold = 5
    mvarfontBold = vData
End Property


Public Property Get FontBold() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.fontBold
    FontBold = mvarfontBold
End Property



Public Property Let CntOfStreatOutNow(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CntOfStreatOutNow = 5
    mvarCntOfStreatOutNow = vData
End Property


Public Property Get CntOfStreatOutNow() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CntOfStreatOutNow
    CntOfStreatOutNow = mvarCntOfStreatOutNow
End Property



Public Property Let CntOfStreatOutSetting(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CntOfStreatOutSetting = 5
    mvarCntOfStreatOutSetting = vData
End Property


Public Property Get CntOfStreatOutSetting() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CntOfStreatOutSetting
    CntOfStreatOutSetting = mvarCntOfStreatOutSetting
End Property



Public Property Let bChkAlarm5(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkAlarm5 = 5
    mvarbChkAlarm5 = vData
End Property


Public Property Get bChkAlarm5() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkAlarm5
    bChkAlarm5 = mvarbChkAlarm5
End Property



Public Property Let bChkAlarm4(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkAlarm4 = 5
    mvarbChkAlarm4 = vData
End Property


Public Property Get bChkAlarm4() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkAlarm4
    bChkAlarm4 = mvarbChkAlarm4
End Property



Public Property Let bChkAlarm3(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkAlarm3 = 5
    mvarbChkAlarm3 = vData
End Property


Public Property Get bChkAlarm3() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkAlarm3
    bChkAlarm3 = mvarbChkAlarm3
End Property



Public Property Let bChkAlarm2(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkAlarm2 = 5
    mvarbChkAlarm2 = vData
End Property


Public Property Get bChkAlarm2() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkAlarm2
    bChkAlarm2 = mvarbChkAlarm2
End Property



Public Property Let bChkAlarm1(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bChkAlarm1 = 5
    mvarbChkAlarm1 = vData
End Property


Public Property Get bChkAlarm1() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bChkAlarm1
    bChkAlarm1 = mvarbChkAlarm1
End Property



Public Property Let CntOfAlarm5(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CntOfAlarm5 = 5
    mvarCntOfAlarm5 = vData
End Property


Public Property Get CntOfAlarm5() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CntOfAlarm5
    CntOfAlarm5 = mvarCntOfAlarm5
End Property



Public Property Let CntOfAlarm4(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CntOfAlarm4 = 5
    mvarCntOfAlarm4 = vData
End Property


Public Property Get CntOfAlarm4() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CntOfAlarm4
    CntOfAlarm4 = mvarCntOfAlarm4
End Property



Public Property Let CntOfAlarm3(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CntOfAlarm3 = 5
    mvarCntOfAlarm3 = vData
End Property


Public Property Get CntOfAlarm3() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CntOfAlarm3
    CntOfAlarm3 = mvarCntOfAlarm3
End Property



Public Property Let CntOfAlarm2(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CntOfAlarm2 = 5
    mvarCntOfAlarm2 = vData
End Property


Public Property Get CntOfAlarm2() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CntOfAlarm2
    CntOfAlarm2 = mvarCntOfAlarm2
End Property



Public Property Let bAlarm5(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bAlarm5 = 5
    mvarbAlarm5 = vData
End Property


Public Property Get bAlarm5() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bAlarm5
    bAlarm5 = mvarbAlarm5
End Property



Public Property Let bAlarm4(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bAlarm4 = 5
    mvarbAlarm4 = vData
End Property


Public Property Get bAlarm4() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bAlarm4
    bAlarm4 = mvarbAlarm4
End Property



Public Property Let bAlarm3(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bAlarm3 = 5
    mvarbAlarm3 = vData
End Property


Public Property Get bAlarm3() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bAlarm3
    bAlarm3 = mvarbAlarm3
End Property



Public Property Let bAlarm2(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bAlarm2 = 5
    mvarbAlarm2 = vData
End Property


Public Property Get bAlarm2() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bAlarm2
    bAlarm2 = mvarbAlarm2
End Property



Public Property Let CntOfAlarm1(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CntOfAlarm1 = 5
    mvarCntOfAlarm1 = vData
End Property


Public Property Get CntOfAlarm1() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CntOfAlarm1
    CntOfAlarm1 = mvarCntOfAlarm1
End Property



Public Property Let bAlarm1(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bAlarm1 = 5
    mvarbAlarm1 = vData
End Property


Public Property Get bAlarm1() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bAlarm1
    bAlarm1 = mvarbAlarm1
End Property



Public Property Let bStreateOutCntCheck(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bStreateOutCntCheck = 5
    mvarbStreateOutCntCheck = vData
End Property


Public Property Get bStreateOutCntCheck() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bStreateOutCntCheck
    bStreateOutCntCheck = mvarbStreateOutCntCheck
End Property



Public Property Let DirectTimeSec(ByVal vData As Integer)
Attribute DirectTimeSec.VB_Description = "������ ���̳��� ���� �������� �ʰ�\r\n������ �����Ҷ� �޼����� �˸� �⺻20��"
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.DirectTimeSec = 5
    mvarDirectTimeSec = vData
End Property


Public Property Get DirectTimeSec() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.DirectTimeSec
    DirectTimeSec = mvarDirectTimeSec
End Property



Public Property Let CreateRepeatInterval(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.CreateRepeatInterval = 5
    mvarCreateRepeatInterval = vData
End Property


Public Property Get CreateRepeatInterval() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.CreateRepeatInterval
    CreateRepeatInterval = mvarCreateRepeatInterval
End Property



Public Property Let TimeOutSecStudy(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.TimeOutSecStudy = 5
    mvarTimeOutSecStudy = vData
End Property


Public Property Get TimeOutSecStudy() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.TimeOutSecStudy
    TimeOutSecStudy = mvarTimeOutSecStudy
End Property



Public Property Let TimeOutSec(ByVal vData As Integer)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.TimeOutSec = 5
    mvarTimeOutSec = vData
End Property


Public Property Get TimeOutSec() As Integer
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.TimeOutSec
    TimeOutSec = mvarTimeOutSec
End Property



Public Property Let bSetTimeOutStudy(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bSetTimeOutStudy = 5
    mvarbSetTimeOutStudy = vData
End Property


Public Property Get bSetTimeOutStudy() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bSetTimeOutStudy
    bSetTimeOutStudy = mvarbSetTimeOutStudy
End Property



Public Property Let bSetTimeOut(ByVal vData As Boolean)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.bSetTimeOut = 5
    mvarbSetTimeOut = vData
End Property


Public Property Get bSetTimeOut() As Boolean
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.bSetTimeOut
    bSetTimeOut = mvarbSetTimeOut
End Property



Public Property Let FontSize(ByVal vData As Single)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.fontsize = 5
    mvarfontsize = vData
End Property


Public Property Get FontSize() As Single
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.fontsize
    FontSize = mvarfontsize
End Property

Public Property Let FontName(ByVal vData As String)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.fontsize = 5
    mvarFontName = vData
End Property


Public Property Get FontName() As String
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.fontsize
    FontName = mvarFontName
End Property


Public Property Let PenColor(ByVal vData As Long)
'���Թ��� ���ʿ��� �Ӽ��� ���� �Ҵ��� �� ���˴ϴ�.
'Syntax: X.fontsize = 5
    mvarPenColor = vData
    WritePrivateProfileString gUserid, "PenColor", CStr(mvarPenColor), App.Path & cFTP.getINIFILE()
End Property


Public Property Get PenColor() As Long
'���Թ��� �����ʿ��� �Ӽ����� �˻��� �� ���˴ϴ�.
'Syntax: Debug.Print X.fontsize
    PenColor = mvarPenColor
End Property



Public Function save()
    RaiseEvent saveProfile
End Function
