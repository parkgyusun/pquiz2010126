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
'UserControl을 테스트하거나 디버깅하거나 검수하는데 필요한 작업들을 아래에 목록으로 정리했습니다.
'
'A) UserControl 테스트 프로그램 만들기
'
'표준 EXE에 컨트롤을 삽입했는지 또는 UserControl에 대한 ActiveX 컨트롤 프로젝트를 만들었는지에 따라 UserControl 테스트 프로그램을 설치하는 방법이 두 가지 있습니다.
'
'ActiveX 컨트롤 프로젝트를 만들었다면 아래 절차에 따라 테스트 프로그램을 설치합니다.
'
'1)  UserControl을 저장합니다.
'2) 실행 모드에서 컨트롤을 배치하려면 UserControl의 디자이너를 닫습니다.
'3) 테스트 프로젝트를 아직 만들지 않았다면 [파일] 메뉴에서 [프로젝트 추가]를 선택해서 표준 EXE 프로젝트를 추가합니다.
'4) 도구 상자에서 UserControl 아이콘을 두 번 눌러 Form1에 UserControl 인스턴스를 배치합니다. 필요하다면 컨트롤을 옮기거나 크기를 조정할 수 있습니다.
'5) 프로젝트 그룹을 저장합니다. 이후의 개발과 테스트 세션에서 프로젝트 그룹을 열어 한꺼번에 두 프로젝트를 모두 열 수 있습니다.
'
'기존 표준 EXE 프로젝트에 UserControl을 삽입했다면 아래 절차를 따르십시오.
'
'1)  UserControl을 저장합니다.
'2)  실행 모드에서 컨트롤을 배치하려면 UserControl의 디자이너를 닫습니다.
'3)  프로젝트 창에서 표준 EXE 프로젝트의 Form1을 두 번 눌러 디자이너를 엽니다.
'4)  도구 상자에서 UserControl 아이콘을 두 번 눌러 Form1에 UserControl 인스턴스를 배치합니다. 필요하다면 컨트롤을 옮기거나 크기를 조정할 수 있습니다.
'
'B) 디자인 모드와 실행 모드에서 컨트롤 동작을 테스트하기
'
'1)  테스트 프로젝트에서 Form1에 배치한 컨트롤을 선택하고 <F4>키를 눌러 속성 창을 엽니다. 컨트롤에 추가한 속성이 나타나는지 그리고 변경할 수 있는지 확인합니다.
'2)  Form1을 닫은 후 다시 엽니다. 그런 다음 컨트롤의 속성값이 올바로 저장되고 검색되었는지 확인합니다.
'3)  Form1에 배치한 컨트롤을 두 번 누르고 코드 창의 마우스 오른쪽 단추(프로시저)를 눌렀을 때 나타나는 드롭다운 목록에 해당 이벤트가 나타나는지 확인합니다
'4)  컨트롤의 이벤트 프로시저에 테스트 코드를 추가합니다.
'5)  다른 컨트롤들을 추가하고 컨트롤의 이벤트 프로시저에 코드를 삽입한 다음 컨트롤의 속성과 메서드의 런타임 동작을 테스트합니다.
'6)  <F5>키를 눌러 테스트 프로젝트를 실행하고 컨트롤의 런타임 동작을 테스트합니다.
'
'C) 검수가 끝나고 완전한 기능을 갖춘 컨트롤 만들기(마법사가 제공하지 않은 코드 추가)
'
'1)  폼에 구성 요소 컨트롤이 있으면 일부 이벤트와 속성을 다중 구성 요소 컨트롤에 매핑해야 합니다. 예를 들어 BackColor 속성은 아마도 UserControl과 Label 컨트롤의 BackColor 속성에 매핑해야 할 것입니다. MouseMove 이벤트는 모든 구성 요소 컨트롤의 MouseMove 이벤트에 매핑해야 합니다.
'2)  X와 Y 좌표를 지정하는 MouseMove 등의 모든 이벤트에 좌표 변환을 추가합니다.
'3)  MoursePointer와 Borderstyle 등의 열거를 가지는 속성은 모두 속성의 데이터 형식을 해당 열거 이름으로 변경하여 열거 요소가 속성 창에 나타나게 합니다. 이 경우에는 MousePointerConstants 및 BorderStyleConstants입니다.
'4)  자신의 고유한 속성에 필요한 사용자 정의 열거와 코드를 추가해서 사용자 정의 열거를 유효화합니다.
'5)  ReadProperties 이벤트에 오류 잡기를 추가하여 .frm 파일에 수동으로 편집해서 생긴 것같은 잘못된 값이나 데이터 형식으로부터 보호합니다. 모든 속성을 이런 오류가 발생하면 코드를 추가해서 기본 설정값으로 전환합니다. (온라인 설명서에서 "컨트롤 속성 저장"과 "디자인 모드 또는 실행 모드 전용 속성 작성"을 참조하십시오.)
'6)  구성 요소 컨트롤이 있다면 UserControl_Resize에 코드를 추가해서 컨트롤의 크기를 조정할 때 구성 요소 컨트롤의 크기도 조정합니다.
'7)  Enabled 속성에 대한 프로시저 ID를 설정하십시오. 그러면 컨트롤이 사용가능할 경우와 사용 불가능한 경우에 다른 ActiveX 컨트롤과 동일하게 작동합니다.
'8)  이 마법사는 컨트롤의 속성을 비슷한 이름을 가진 구성 요소 컨트롤( UserControl)의 속성에 매핑합니다. 경우에 따라 한 속성을 이름이 다른 속성에 매핑하려는 경우도 있을 수 잇습니다. 예를 들어 ShapeLabel은 구성 요소 Shape 컨트롤의 FillColor에 BackColor를 매핑합니다. 이렇게 재매핑하려면 수동으로 해야 합니다.
'9)  컨트롤의 크기에 영향을 주는 속성들(AutoSize 속성을 가진 컨트롤에서의 글꼴 크기)은 Property Let에서 크기 조정 코드를 호출해야 합니다.
'10) 사용자가 만든 컨트롤이면 UserControl의 Paint 이벤트에 코드를 추가하여 컨트롤의 모양을 그립니다. (온라인 설명서에서 "사용자 그리기 컨트롤"와 "컨트롤에 대한 포커스 사용 방법"을 참조하십시오.)
'11)  컨트롤의 하나 이상의 속성이 데이터에 바인딩되면 온라인 설명서에서 "컨트롤을 데이터 원본으로 바인딩"을 참조하십시오.
'12)  컨트롤에 이외의 기능을 추가합니다. 온라인 설명서에서 "Visual Basic ActiveX 컨트롤 특징'에 있는 항목들을 자세히 읽으면 도움이 될 것입니다.
'
'(이러한 작업 항목에 대한 예제를 보려면 CtlPlus.vbg 예제 응용 프로그램을 참조하십시오.)
'
'마법사를 다시 실행하고 UserControl을 선택해서 컨트롤을 수정할 수 있습니다.
'
'사용자의 UserControlsUse에 대한 속성 페이지를 만들려면 속성 페이지 마법사를 사용합니다.
'
'ActiveX 컨트롤 만들기와 테스트에 대한 자세한 내용은 4장 "ActiveX 컨트롤 작성"과 9장 "ActiveX 컨트롤 구성"을 참조하십시오.
'
'6장 "구성 요소 디자인의 일반 원리"와 7장 " ActiveX 구성 요소의 디버깅, 테스트, 배포"에서 유용한 추가 정보를 얻을 수 있습니다.'기본 속성 값:

'Const m_def_FormatMask = "##0.##"
'Const m_def_MinVal = -999.99
'Const m_def_MaxVal = 999.99
'속성 변수:
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
'이벤트 선언:
Event Validate(Cancel As Boolean) 'MappingInfo=Text1,Text1,-1,Validate
Attribute Validate.VB_Description = "컨트롤이 유효성 검사를 일으킨 컨트롤에 대한 포커스를 잃을 때 발생합니다."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=Text1,Text1,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "OLE 끌기/놓기 작업이 수동 또는 자동으로 시작될 때 발생합니다."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=Text1,Text1,-1,OLESetData
Attribute OLESetData.VB_Description = "OLEDragStart 이벤트 중 DataObject 에 주어지지 않은 데이터를 놓기 대상이 요청할 때 OLE 끌기/놓기 원본 컨트롤에서 발생합니다."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=Text1,Text1,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "마우스 커서의 변화가 필요할 때 OLE 끌기/놓기 작업의 원본 컨트롤에서 발생합니다."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=Text1,Text1,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "OLEDropMode 속성이 수동으로 설정된 상태일 경우 OLE 끌기/놓기 작업 중 마우스를 컨트롤 위로 이동하면 발생합니다."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "OLE 끌기/놓기 작업을 통해 데이터를 컨트롤에 놓을 때 그리고 OLEDropMode가 수동으로 설정될 때 발생합니다."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=Text1,Text1,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "수동 또는 자동 끌기/놓기가 완료되거나 취소된 후 OLE 끌기/놓기 원본 컨트롤에서 발생합니다."
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Attribute Change.VB_Description = "컨트롤의 내용이 변경될 때 발생합니다."
Event Click() 'MappingInfo=Text1,Text1,-1,Click
Attribute Click.VB_Description = "개체에서 마우스 단추를 눌렀다가 놓을 때 발생합니다."
Event DblClick() 'MappingInfo=Text1,Text1,-1,DblClick
Attribute DblClick.VB_Description = "마우스 단추를 개체에서 누르고 놓은 후 다시 누르고 놓으면 발생합니다."
Event KeyDown(keycode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Attribute KeyDown.VB_Description = "개체에 포커스가 있을 때 키를 누르면 발생합니다."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Attribute KeyPress.VB_Description = "ANSI키를 누르고 놓았을 경우 발생합니다."
Event KeyUp(keycode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp
Attribute KeyUp.VB_Description = "개체에 포커스가 있을 때 키를 놓으면 발생합니다."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseDown
Attribute MouseDown.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 누르면 발생합니다."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseMove
Attribute MouseMove.VB_Description = "마우스를 움직일 경우 발생합니다."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseUp
Attribute MouseUp.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 놓으면 발생합니다."
'기본 속성 값:
Const m_def_AllowNull = False

Const m_def_FormatMask = "#,##0"
Const m_def_MinVal = -9999999999#
Const m_def_MaxVal = 9999999999#

Dim onReadProperte As Boolean




'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "개체의 텍스트나 그래픽을 표시하기 위해 사용되는 배경색을 반환하거나 설정합니다."
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "개체에서 텍스트나 그래픽을 표시하는 전경색을 반환하거나 설정합니다."
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "사용자가 만든 이벤트에 대해 개체가 응답할 수 있는지의 여부를 결정하는 값을 반환하거나 설정합니다."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font 개체를 반환합니다."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property
'
''경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
''MappingInfo=UserControl,UserControl,-1,BackStyle
'Public Property Get BackStyle() As Integer
'    BackStyle = UserControl.BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    UserControl.BackStyle() = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "개체 테두리 유형을 반환하거나 설정합니다."
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Text1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "개체를 완전히 다시 그리게 합니다."
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
    MsgBox "공백은 허용되지 않습니다.", vbExclamation
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
''경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
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
''경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
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
''경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
''MemberInfo=4,0,0,999.99
'Public Property Get MaxVal() As Double
'    MaxVal = m_MaxVal
'End Property
'
'Public Property Let MaxVal(ByVal New_MaxVal As Double)
'    m_MaxVal = New_MaxVal
'    PropertyChanged "MaxVal"
'End Property

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
'    m_FormatMask = m_def_FormatMask
'    m_MinVal = m_def_MinVal
'    m_MaxVal = m_def_MaxVal
    m_FormatMask = m_def_FormatMask
    m_MinVal = m_def_MinVal
    m_MaxVal = m_def_MaxVal
    m_AllowNull = m_def_AllowNull
End Sub

'저장소에서 속성값을 로드합니다.
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
    ''Text1.IMEMode = PropBag.ReadProperty("IMEMode", 0)''영문에서는 오류가 발생함.
    Text1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Text1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Text1.FontSize = PropBag.ReadProperty("FontSize", 14)
    Text1.FontName = PropBag.ReadProperty("FontName", "굴림")
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

'속성값을 저장소에 기록합니다.
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=4,0,0,-999.99
Public Property Get MinVal() As Double
    MinVal = m_MinVal
End Property

Public Property Let MinVal(ByVal New_MinVal As Double)
    m_MinVal = New_MinVal
    PropertyChanged "MinVal"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=4,0,0,999.99
Public Property Get MaxVal() As Double
    MaxVal = m_MaxVal
End Property

Public Property Let MaxVal(ByVal New_MaxVal As Double)
    m_MaxVal = New_MaxVal
    PropertyChanged "MaxVal"
End Property


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "개체의 관련 컨텍스트 번호를 반환하거나 설정합니다."
    WhatsThisHelpID = Text1.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    Text1.WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

Private Sub Text1_Validate(Cancel As Boolean)
If AllowNull = True Then
    If Trim(Text1) = "" Then
        MsgBox "공백문자는 허용되지 않습니다.", vbExclamation
        Cancel = True
    End If
End If
    'RaiseEvent Validate(Cancel)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "마우스가 컨트롤에 일시 정지되어 있을 때 표시되는 텍스트를 반환하거나 설정합니다."
    ToolTipText = Text1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Text1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "현재 선택된 텍스트를 포함하는 문자열을 반환하거나 설정합니다."
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Text1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "선택된 텍스트의 시작점을 반환하거나 설정합니다."
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "선택된 문자의 개수를 반환하거나 설정합니다."
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "개체가 수직/수평 스크롤 막대를 가지는지의 여부를 나타내는 값을 반환하거나 설정합니다."
    ScrollBars = Text1.ScrollBars
End Property

'"Scale" 다음에 나오는 밑줄은 VBA의 예약어이므로
'필요합니다.
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=5
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "텍스트 표시 방향을 결정하고 양방향 시스템에서 시각적 모양을 조정합니다."
    RightToLeft = Text1.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    Text1.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "사용자가 입력한 문자 또는 자리 표시자 문자가 컨트롤에 표시되었는지 결정하는 값을 반환하거나 설정합니다. "
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "이 개체가 OLE 놓기 대상으로 실행될 수 있는지 그리고 이것이 자동으로 혹은 프로그램의 제어하에 일어나는지의 여부를 반환하거나 설정합니다."
    OLEDropMode = Text1.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    Text1.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,OLEDragMode
Public Property Get OLEDragMode() As Integer
Attribute OLEDragMode.VB_Description = "이 컨트롤이 OLE 끌기/놓기 원본으로 실행될수 있는지 그리고 이 프로세스가 자동으로 혹은 프로그램의 제어하에 시작될수 있는지의 여부를 반환하거나 설정합니다."
    OLEDragMode = Text1.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
    Text1.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "주어진 컨트롤을 원본으로 사용하여 OLE 끌기/놓기 이벤트를 시작합니다."
    Text1.OLEDrag
End Sub

Private Sub Text1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "컨트롤이 여러 줄의 텍스트를 받아들일 수 있는지 여부를 결정하는 값을 반환하거나 설정합니다."
    MultiLine = Text1.MultiLine
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "개체에 마우스 포인터가 있을 경우 표시되는 마우스 포인터의 형식을 반환하거나 설정합니다."
    MousePointer = Text1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Text1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "사용자 정의 마우스 아이콘을 설정합니다."
    Set MouseIcon = Text1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Text1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "컨트롤에 들어갈 수 있는 문자의 최대수를 반환하거나 설정합니다."
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "컨트롤의 편집 가능 여부를 결정합니다."
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "원본 응용 프로그램과 대상 컨트롤의 항목을 반환하거나 설정합니다."
    LinkTopic = Text1.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
    Text1.LinkTopic() = New_LinkTopic
    PropertyChanged "LinkTopic"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "컨트롤이 DDE 메시지의 응답을 기다리는 시간을 반환하거나 설정합니다. "
    LinkTimeout = Text1.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
    Text1.LinkTimeout() = New_LinkTimeout
    PropertyChanged "LinkTimeout"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkSend
Public Sub LinkSend()
Attribute LinkSend.VB_Description = "PictureBox의 내용을 DDE 대화의 대상 응용 프로그램으로 전송합니다."
    Text1.LinkSend
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkRequest
Public Sub LinkRequest()
Attribute LinkRequest.VB_Description = "원본 DDE 응용 프로그램에 요청하여 Label이나 PictureBox, Textbox 컨트롤 내용을 업데이트하게 합니다."
    Text1.LinkRequest
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkPoke
Public Sub LinkPoke()
Attribute LinkPoke.VB_Description = "DDE 대화의 원본 응용 프로그램으로 Label, PictureBox, TextBox의 내용을 변경합니다."
    Text1.LinkPoke
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkMode
Public Property Get LinkMode() As Integer
Attribute LinkMode.VB_Description = "DDE 대화에 사용된 연결 형태를 반환하거나 설정하고 연결을 활성화시킵니다."
    LinkMode = Text1.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As Integer)
    Text1.LinkMode() = New_LinkMode
    PropertyChanged "LinkMode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "다른 응용 프로그램과 함께 DDE 대화의 대상 컨트롤을 통과한 데이터를 반환하거나 설정합니다."
    LinkItem = Text1.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
    Text1.LinkItem() = New_LinkItem
    PropertyChanged "LinkItem"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,LinkExecute
Public Sub LinkExecute(ByVal Command As String)
Attribute LinkExecute.VB_Description = "DDE 대화의 원본 응용 프로그램에 명령 문자열을 전송합니다."
    Text1.LinkExecute Command
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,IMEMode
Public Property Get IMEMode() As Integer
Attribute IMEMode.VB_Description = "입력 메서드 편집기의 현재 운영 모드를 반환하거나 설정합니다."
    ''IMEMode = Text1.IMEMode''영문에서는 오류가 발생함.
End Property

Public Property Let IMEMode(ByVal New_IMEMode As Integer)
    'Text1.IMEMode() = New_IMEMode''영문에서는 오류가 발생함.
    PropertyChanged "IMEMode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Microsoft Windows의 핸들을 개체의 창으로 반환합니다."
    hwnd = UserControl.hwnd
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "컨트롤이 포커스를 잃었을 경우 Masked edit 컨트롤을 숨기기로 선택할지의 여부를 지정합니다."
    HideSelection = Text1.HideSelection
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "밑줄 글꼴 유형을 반환하거나 설정합니다."
    FontUnderline = Text1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Text1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "취소선 글꼴 유형을 반환하거나 설정합니다."
    FontStrikethru = Text1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Text1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "주어진 단계의 각 행에 나타나는 글꼴 크기를 포인트 단위로 지정합니다."
    FontSize = Text1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Text1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "주어진 단계의 각 행에 나타나는 글꼴의 이름을 지정합니다."
    FontName = Text1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Text1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "기울임 글꼴 유형을 반환하거나 설정합니다."
    FontItalic = Text1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Text1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "굵게 글꼴 유형을 반환하거나 설정합니다."
    FontBold = Text1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Text1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,DataSource
Public Property Get DataSource() As Variant ' DataSource
Attribute DataSource.VB_Description = "Data 컨트롤을 지정하는 값을 설정합니다. 이 Data 컨트롤을 통하여 현재 컨트롤이 데이터베이스에 바인딩됩니다."
    Set DataSource = Text1.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As Variant)
    Set Text1.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "데이터 연결에 대한 DataMember를 설명하는 값을 반환하거나 설정합니다."
    DataMember = Text1.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    Text1.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,DataFormat
Public Property Get DataFormat() As Variant  ' OLE_OPTEXCLUSIVE ' IStdDataFormatDisp
Attribute DataFormat.VB_Description = "이 구성 요소의 바인딩 가능 속성에 대해 사용하기 위해 DataFormat 개체를 반환합니다."
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
        MsgBox "최대값 범위 오류.(최대값: " & m_MaxVal & ")", vbExclamation, "오류"
        onChange = False
        Text1.Text = m_MaxVal
        
    ElseIf m_MinVal > CDbl(Text1.Text) Then
        MsgBox "최소값 범위 오류.(최소값: " & m_MinVal & ")", vbExclamation, "오류"
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,CausesValidation
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "포커스를 잃은 컨트롤에서 유효성 검사를 하는지 여부를 반환하거나 설정합니다."
    CausesValidation = Text1.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    Text1.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "실행 모드에서 개체가 3D 효과로 그려지는지의 여부를 반환하거나 설정합니다."
    Appearance = Text1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    Text1.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "CheckBox나 OptionButton 또는  컨트롤 텍스트의 맞춤을 반환하거나 설정합니다."
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
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,False
Public Property Get AllowNull() As Boolean
Attribute AllowNull.VB_Description = "NULL을 허용할지의 여부"
    AllowNull = m_AllowNull

End Property

Public Property Let AllowNull(ByVal New_AllowNull As Boolean)
    m_AllowNull = New_AllowNull
    If m_AllowNull Then
        CausesValidation = True
    End If
    PropertyChanged "AllowNull"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "컨트롤에 포함된 텍스트를 반환하거나 설정합니다."
Attribute Text.VB_UserMemId = 0
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

