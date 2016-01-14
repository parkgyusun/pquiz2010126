Attribute VB_Name = "mduPutool"

Option Explicit

Public cmd As String

Public gbZipStop As Boolean

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
'  대화 모드에 대한 비트 필드
Global Const IME_CMODE_ALPHANUMERIC = &H0
Global Const IME_CMODE_NATIVE = &H1
Global Const IME_CMODE_CHINESE = IME_CMODE_NATIVE
Global Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Global Const IME_CMODE_JAPANESE = IME_CMODE_NATIVE
Global Const IME_CMODE_KATAKANA = &H2                   '  only effect under IME_CMODE_NATIVE
Global Const IME_CMODE_LANGUAGE = &H3
Global Const IME_CMODE_FULLSHAPE = &H8
Global Const IME_CMODE_ROMAN = &H10
Global Const IME_CMODE_CHARCODE = &H20
Global Const IME_CMODE_HANJACONVERT = &H40
Global Const IME_CMODE_SOFTKBD = &H80
Global Const IME_CMODE_NOCONVERSION = &H100
Global Const IME_CMODE_EUDC = &H200
Global Const IME_CMODE_SYMBOL = &H400
Global Const IME_CMODE_FIXED = &H800


Global Const IME_SMODE_NONE = &H0
Global Const IME_SMODE_PLAURALCLAUSE = &H1
Global Const IME_SMODE_SINGLECONVERT = &H2
Global Const IME_SMODE_AUTOMATIC = &H4
Global Const IME_SMODE_PHRASEPREDICT = &H8
Global Const IME_SMODE_CONVERSATION = &H10

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Public status As Panel

Global Const LISTVIEW_MODE0 = "큰 아이콘"
Global Const LISTVIEW_MODE1 = "작은 아이콘"
Global Const LISTVIEW_MODE2 = "목록 보기"
Global Const LISTVIEW_MODE3 = "자세히 보기"

Public gScreenHourGLS As Boolean

  
Public Type HSL
  hue As Long
  Saturation As Long
  Luminance As Long
End Type

Public STRCON   As String
Public Const GPWD2 = "d=zk"

Global gUserid As String
Global gPocket As String
Global gOrder As Long
Global gLastNew As Long
Global gHangSu As Long
Global gIsNew As Boolean
Global gNowNum As Long
Global gbReadTu03 As Boolean
Global gChasu As Long

Public fMainForm As frmMain
'Public Con As New ADODB.Connection
'Public rs As New ADODB.Recordset
Public sSql As String




Public SELFLAG As Boolean

Global cFTP As New clsFTP
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

'//결과값을 여러가지 형태로 보여주는 result
'Public Type ut_bRecordSet
'   result As Boolean       ' Query의 결과 성공( True) /  실패 ( False)
'   Error As Boolean        ' Error 유무   Error발생 (TRUE) / Error미 발생 (FALSE)
'   nrow As Integer
'   rs As ADODB.Recordset   ' Query의 결과에 의한 RecordSet
'End Type

'Dim conHandle As ADODB.Connection

'Public SQLCA As ut_bRecordSet


Sub Main()
'    MsgBox App.Path
    cmd = Command
    frmMain.Show
'    End
End Sub



'Function dbversioncheck() As Boolean
'
'    Dim sdbid As String
'    Dim lver As Long
'
'    sdbid = getTableVal("dbid", "ver", "")
'    lver = getTableVal("ver", "ver", "")
'
'    If sdbid = GDBID Or lver = GDBVER Then
'        dbversioncheck = True
'    Else
'        dbversioncheck = False
'        MsgBox "데이터와 프로그램의 버젼이 일치하지 않습니다.", vbCritical
'    End If
'
'End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next


    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer


    '폼의 캡션을 설정합니다.
    frm.Caption = LoadResString(CInt(frm.Tag))
    

    '글꼴을 설정합니다.
    Set fnt = frm.Font
    fnt.Name = LoadResString(20)
    fnt.Size = CInt(LoadResString(21))
    

    '메뉴 항목의 Caption 속성과
    '다른 모든 컨트롤의 Tag 속성을 사용하여
    '컨트롤의 캡션을 설정합니다.
    For Each ctl In frm.Controls
        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next

End Sub

'Public Function Con_Open() As Boolean
'On Error GoTo ErrTrap
'If Con.State = 0 Then
'   MsgBox "연결중입니다. 30초정도 걸립니다... (확인 클릭)", vbExclamation
'    Con.Open STRCON
'   MsgBox "연결완료               ", vbExclamation
'
''    Con.CommandTimeout = 300
'End If
'Exit Function
'ErrTrap:
'    Con_Open = True
'    MsgBox Err.Description, vbCritical
'    MsgBox "위와 같은 문제로 연결을 실패하여 프로그램을 종료합니다.(문의:010-4739-1183 박규선)"
'    End '프로그램 종료
'End Function
'
'Public Function Con_Close() As Boolean
'
'Exit Function
'
'On Error GoTo ErrTrap
'If Con.State = 1 Then
'    Con.Close
'End If
'
'Exit Function
'
'ErrTrap:
'    Con_Close = True
'    MsgBox Err.Description, vbCritical
'
'End Function

Public Function GETYMD() As String
GETYMD = Format(Now, "YYYYMMDD")
End Function

Public Function GETTIME() As String
GETTIME = Format(Now, "hhMMSS")
End Function

Public Sub selall(ctl As Control)
ctl.SelStart = 0
ctl.SelLength = Len(ctl)
End Sub

Public Function ONLYNUM(KeyAscii As Integer) As Integer
If IsNumeric(Chr(KeyAscii)) Then
    ONLYNUM = KeyAscii
Else
    ONLYNUM = 0
End If
End Function
Function IDEMODE() As Boolean
On Error GoTo errHandler
    Debug.Print 1 / 0
Exit Function
errHandler:
IDEMODE = True
End Function


Public Function GETLASTTXT(SN As String, DL As String) As String
    Dim Pos As Long
    Pos = InStrRev(SN, DL)
    If Pos = 0 Then
       GETLASTTXT = ""
    Else
        GETLASTTXT = Mid(SN, Pos + Len(DL))
    End If

End Function


Public Sub AlwaysOnTop(f As Form, OnTop As Boolean)
    'hwndInsertAfter values
    Const HWND_TOP = 0
    Const HWND_BOTTOM = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    'wFlags values
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER = &H4
    Const SWP_NOREDRAW = &H8
    Const SWP_NOACTIVATE = &H10
    Const SWP_FRAMECHANGED = &H20           'The frame changed: send WM_NCCALCSIZE
    Const SWP_SHOWWINDOW = &H40
    Const SWP_HIDEWINDOW = &H80
    Const SWP_NOCOPYBITS = &H100
    Const SWP_NOOWNERZORDER = &H200    'Don't do owner Z ordering
    Const SWP_DRAWFRAME = SWP_FRAMECHANGED
    Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

    If OnTop = True Then
        'Turn on the TopMost attribute.
        SetWindowPos f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    ElseIf OnTop = False Then
        'Turn off the TopMost attribute.
        SetWindowPos f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub
Public Function selectSeries(lst As Object, Optional dli As String) As String

Dim i As Integer
Dim nLen As Long
Dim str As String

If IsMissing(dli) Then
    For i = 1 To lst.ListCount - 1
        If lst.Selected(i) Then
            str = str & "'" & lst.List(i) & "',"
        End If
    Next
Else
    For i = 1 To lst.ListCount - 1
        If lst.Selected(i) Then
            str = str & "'" & GETLASTTXT(lst.List(i), dli) & "',"
        End If
    Next
End If
str = str & "''"

selectSeries = str

End Function

Public Function itemSeries2(lst As Object, Optional dli As String) As String

Dim i As Integer
Dim nLen As Long
Dim str As String
Dim ListText As String


If IsMissing(dli) Then
    For i = lst.ListCount - 1 To 0 Step -1
        ListText = lst.ListText(i)
        str = str & "'" & ListText & "',"
    Next
Else
    For i = lst.ListCount - 1 To 0 Step -1
        ListText = lst.ListText(i)
        str = str & "'" & GETLASTTXT(ListText, dli) & "',"
    Next
End If

str = str & "''"

itemSeries2 = str

End Function


Function str2Date(ByVal str As String) As Date

Dim ymd As String

If str = "" Then
str = GETYMD
End If

Debug.Assert Len(str) = 8

ymd = Mid(str, 1, 4) & "-" & Mid(str, 5, 2) & "-" & Mid(str, 7, 2)

str2Date = CDate(ymd)

'Debug.Print str2Date

End Function

Function date2Str(ByVal dt As Date) As String

date2Str = Format(dt, "yyyyMMdd")

End Function

'====================================================
'버그수정 round round(2.5)=> 2 Error Not 3
' round(213.14,-1)=> 210 But RunTime Error 5
'====================================================
Public Function Round15( _
    ByVal dblNumber As Double, _
    Optional ByVal lngDecimals As Long = 0) _
    As Double
' By Gustav, gustav@cactus.dk, 20020509
' Modification of Round14
' by Donald, donald@xbeat.net, 20020419
' modification of Round10 inspired by Jost's Round13 (Variant = CDec!)

  Static dblNumberPrevious    As Double
  Static lngDecimalsPrevious  As Long
  Static varFactor            As Variant
  Static dblFactorInv         As Double
  Static dblValue             As Double

  Dim booNewDecimals          As Boolean
  Dim booNewNumber            As Boolean

  ' Ignore excessive values of lngDecimals and dblValue.
  On Error GoTo Err_Round15

  booNewNumber = dblNumber <> dblNumberPrevious
  booNewDecimals = lngDecimals <> lngDecimalsPrevious Or dblFactorInv = 0

  If booNewDecimals = True Then
    ' Calculate factor for this number of decimals.
    varFactor = CDec(10 ^ lngDecimals)
    dblFactorInv = 10 ^ -lngDecimals
    lngDecimalsPrevious = lngDecimals
  End If

  If booNewDecimals = True Or booNewNumber = True Then
    dblNumberPrevious = dblNumber
    If dblNumber = 0 Then
      dblValue = 0
    ElseIf dblNumber > 0 Then
      dblValue = Int(dblNumber * varFactor + 0.5)
      dblValue = dblValue * dblFactorInv
    Else
      dblValue = -Int(-dblNumber * varFactor + 0.5)
      dblValue = dblValue * dblFactorInv
    End If
  End If

Exit_Round15:
  Round15 = dblValue
  Exit Function

Err_Round15:
  ' Return input value unmodified.
  dblValue = dblNumber
  Resume Exit_Round15

End Function


Public Function HSLToRGB201( _
    ByVal hue As Long, _
    ByVal Saturation As Long, _
    ByVal Luminance As Long, _
    Optional Validate As Boolean _
    ) As Long
' by Donald, donald@xbeat.net, 20011126
' after seeing Thomas Kabir's code (contact@vbfrood.de, http://www.vbfrood.de)
  Dim rR As Single, rG As Single, rB As Single
  Dim rH As Single, rL As Single, rs As Single
  Dim rMin As Single, rMax As Single, rDiff As Single
  
'  If Validate Then ValidateHSL hue, Saturation, Luminance
  
  If Saturation = 0 Then
    ' CLng(CSng(...)) else 127.5 -> 127
    HSLToRGB201 = CLng(CSng(2.55 * Luminance)) * &H10101
  Else
    rH = hue / 60: rs = Saturation / 100: rL = Luminance / 100
    If rL <= 0.5 Then
      rMin = rL * (1 - rs)
    Else
      rMin = rL - rs * (1 - rL)
    End If
    rMax = 2 * rL - rMin
    rDiff = rMax - rMin
    
    Select Case hue \ 60
    Case 0
      rR = rMax
      rB = rMin
      rG = rH * rDiff + rMin
    Case 1
      rG = rMax
      rB = rMin
      rR = rMin - (rH - 2) * rDiff
    Case 2
      rG = rMax
      rR = rMin
      rB = (rH - 2) * rDiff + rMin
    Case 3
      rB = rMax
      rR = rMin
      rG = rMin - (rH - 4) * rDiff
    Case 4
      rB = rMax
      rG = rMin
      rR = (rH - 4) * rDiff + rMin
    Case Else
      rR = rMax
      rG = rMin
      rB = rMin - (rH - 6) * rDiff
    End Select
    HSLToRGB201 = CLng(rB * 255) * &H10000 + CLng(rG * 255) * &H100 + CLng(rR * 255)
  End If
  
End Function


Public Function RGBToHSL201(ByVal RGBValue As Long) As HSL
' by Donald, donald@xbeat.net, 20011126
  Dim r As Long, G As Long, B As Long
  Dim lMax As Long, lMin As Long, lDiff As Long, lSum As Long

  r = RGBValue And &HFF&
  G = (RGBValue And &HFF00&) \ &H100&
  B = (RGBValue And &HFF0000) \ &H10000

  If r > G Then lMax = r: lMin = G Else lMax = G: lMin = r
  If B > lMax Then lMax = B Else If B < lMin Then lMin = B

  lDiff = lMax - lMin
  lSum = lMax + lMin
  
  ' Luminance
  RGBToHSL201.Luminance = lSum / 5.1!
  
  If lDiff Then
    ' Saturation
    If RGBToHSL201.Luminance <= 50& Then
      RGBToHSL201.Saturation = 100 * lDiff / lSum
    Else
      RGBToHSL201.Saturation = 100 * lDiff / (510 - lSum)
    End If
    ' Hue
    Dim q As Single: q = 60 / lDiff
    Select Case lMax
    Case r
      If G < B Then
        RGBToHSL201.hue = 360& + q * (G - B)
      Else
        RGBToHSL201.hue = q * (G - B)
      End If
    Case G
      RGBToHSL201.hue = 120& + q * (B - r)
    Case B
      RGBToHSL201.hue = 240& + q * (r - G)
    End Select
  End If
  
End Function


Public Sub SplitRGB02(ByVal lColor As Long, _
    ByRef lRed As Long, _
    ByRef lGreen As Long, _
    ByRef lBlue As Long)
' by Donald, donald@xbeat.net, 20010922
  lRed = lColor And &HFF
  lGreen = (lColor And &HFF00&) \ &H100&
  lBlue = (lColor And &HFF0000) \ &H10000
End Sub



Public Function GrayScale09(ByVal lColor As Long) As Long
' by Donald, donald@xbeat.net, 20011123
  GrayScale09 = ((77& * (lColor And &HFF&) + _
                 152& * (lColor And &HFF00&) \ &H100& + _
                  28& * ((lColor And &HFF0000) \ &H10000)) \ 256&) * &H10101
End Function


'<<
Public Function ShiftLeft06(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011001
  Select Case ShiftCount
  Case 0&
    ShiftLeft06 = Value
  Case 1&
    If Value And &H40000000 Then
      ShiftLeft06 = (Value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If Value And &H20000000 Then
      ShiftLeft06 = (Value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If Value And &H10000000 Then
      ShiftLeft06 = (Value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If Value And &H8000000 Then
      ShiftLeft06 = (Value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If Value And &H4000000 Then
      ShiftLeft06 = (Value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If Value And &H2000000 Then
      ShiftLeft06 = (Value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If Value And &H1000000 Then
      ShiftLeft06 = (Value And &HFFFFFF) * &H80& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If Value And &H800000 Then
      ShiftLeft06 = (Value And &H7FFFFF) * &H100& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If Value And &H400000 Then
      ShiftLeft06 = (Value And &H3FFFFF) * &H200& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If Value And &H200000 Then
      ShiftLeft06 = (Value And &H1FFFFF) * &H400& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If Value And &H100000 Then
      ShiftLeft06 = (Value And &HFFFFF) * &H800& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFFF) * &H800&
    End If
  Case 12&
    If Value And &H80000 Then
      ShiftLeft06 = (Value And &H7FFFF) * &H1000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If Value And &H40000 Then
      ShiftLeft06 = (Value And &H3FFFF) * &H2000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If Value And &H20000 Then
      ShiftLeft06 = (Value And &H1FFFF) * &H4000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If Value And &H10000 Then
      ShiftLeft06 = (Value And &HFFFF&) * &H8000& Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If Value And &H8000& Then
      ShiftLeft06 = (Value And &H7FFF&) * &H10000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FFF&) * &H10000
    End If
  Case 17&
    If Value And &H4000& Then
      ShiftLeft06 = (Value And &H3FFF&) * &H20000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FFF&) * &H20000
    End If
  Case 18&
    If Value And &H2000& Then
      ShiftLeft06 = (Value And &H1FFF&) * &H40000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FFF&) * &H40000
    End If
  Case 19&
    If Value And &H1000& Then
      ShiftLeft06 = (Value And &HFFF&) * &H80000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFFF&) * &H80000
    End If
  Case 20&
    If Value And &H800& Then
      ShiftLeft06 = (Value And &H7FF&) * &H100000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7FF&) * &H100000
    End If
  Case 21&
    If Value And &H400& Then
      ShiftLeft06 = (Value And &H3FF&) * &H200000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3FF&) * &H200000
    End If
  Case 22&
    If Value And &H200& Then
      ShiftLeft06 = (Value And &H1FF&) * &H400000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1FF&) * &H400000
    End If
  Case 23&
    If Value And &H100& Then
      ShiftLeft06 = (Value And &HFF&) * &H800000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &HFF&) * &H800000
    End If
  Case 24&
    If Value And &H80& Then
      ShiftLeft06 = (Value And &H7F&) * &H1000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7F&) * &H1000000
    End If
  Case 25&
    If Value And &H40& Then
      ShiftLeft06 = (Value And &H3F&) * &H2000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3F&) * &H2000000
    End If
  Case 26&
    If Value And &H20& Then
      ShiftLeft06 = (Value And &H1F&) * &H4000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1F&) * &H4000000
    End If
  Case 27&
    If Value And &H10& Then
      ShiftLeft06 = (Value And &HF&) * &H8000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &HF&) * &H8000000
    End If
  Case 28&
    If Value And &H8& Then
      ShiftLeft06 = (Value And &H7&) * &H10000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H7&) * &H10000000
    End If
  Case 29&
    If Value And &H4& Then
      ShiftLeft06 = (Value And &H3&) * &H20000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H3&) * &H20000000
    End If
  Case 30&
    If Value And &H2& Then
      ShiftLeft06 = (Value And &H1&) * &H40000000 Or &H80000000
    Else
      ShiftLeft06 = (Value And &H1&) * &H40000000
    End If
  Case 31&
    If Value And &H1& Then
      ShiftLeft06 = &H80000000
    Else
      ShiftLeft06 = &H0&
    End If
  End Select
End Function


'>>
Public Function ShiftRight08(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011010
  Select Case ShiftCount
  Case 0&:  ShiftRight08 = Value
  Case 1&:  ShiftRight08 = (Value And &HFFFFFFFE) \ &H2&
  Case 2&:  ShiftRight08 = (Value And &HFFFFFFFC) \ &H4&
  Case 3&:  ShiftRight08 = (Value And &HFFFFFFF8) \ &H8&
  Case 4&:  ShiftRight08 = (Value And &HFFFFFFF0) \ &H10&
  Case 5&:  ShiftRight08 = (Value And &HFFFFFFE0) \ &H20&
  Case 6&:  ShiftRight08 = (Value And &HFFFFFFC0) \ &H40&
  Case 7&:  ShiftRight08 = (Value And &HFFFFFF80) \ &H80&
  Case 8&:  ShiftRight08 = (Value And &HFFFFFF00) \ &H100&
  Case 9&:  ShiftRight08 = (Value And &HFFFFFE00) \ &H200&
  Case 10&: ShiftRight08 = (Value And &HFFFFFC00) \ &H400&
  Case 11&: ShiftRight08 = (Value And &HFFFFF800) \ &H800&
  Case 12&: ShiftRight08 = (Value And &HFFFFF000) \ &H1000&
  Case 13&: ShiftRight08 = (Value And &HFFFFE000) \ &H2000&
  Case 14&: ShiftRight08 = (Value And &HFFFFC000) \ &H4000&
  Case 15&: ShiftRight08 = (Value And &HFFFF8000) \ &H8000&
  Case 16&: ShiftRight08 = (Value And &HFFFF0000) \ &H10000
  Case 17&: ShiftRight08 = (Value And &HFFFE0000) \ &H20000
  Case 18&: ShiftRight08 = (Value And &HFFFC0000) \ &H40000
  Case 19&: ShiftRight08 = (Value And &HFFF80000) \ &H80000
  Case 20&: ShiftRight08 = (Value And &HFFF00000) \ &H100000
  Case 21&: ShiftRight08 = (Value And &HFFE00000) \ &H200000
  Case 22&: ShiftRight08 = (Value And &HFFC00000) \ &H400000
  Case 23&: ShiftRight08 = (Value And &HFF800000) \ &H800000
  Case 24&: ShiftRight08 = (Value And &HFF000000) \ &H1000000
  Case 25&: ShiftRight08 = (Value And &HFE000000) \ &H2000000
  Case 26&: ShiftRight08 = (Value And &HFC000000) \ &H4000000
  Case 27&: ShiftRight08 = (Value And &HF8000000) \ &H8000000
  Case 28&: ShiftRight08 = (Value And &HF0000000) \ &H10000000
  Case 29&: ShiftRight08 = (Value And &HE0000000) \ &H20000000
  Case 30&: ShiftRight08 = (Value And &HC0000000) \ &H40000000
  Case 31&: ShiftRight08 = CBool(Value And &H80000000)
  End Select
End Function


Public Function getLastDay(ByVal pYear As Integer, ByVal pMonth As Integer) As Integer

Dim dt As Date

If pMonth = 12 Then
    pMonth = 1
    pYear = pYear + 1
Else
    pMonth = pMonth + 1
End If

getLastDay = CInt(Format(CDate(pYear & "-" & pMonth & "-01") - 1, "dd"))


End Function

Public Function MakeDbTxt(str1 As String) As String
MakeDbTxt = Replace(str1, "'", "''")

End Function


'==============================================================================
'min ~ max 사이의 랜덤한 값을 가져온다.
'==============================================================================
Public Function rndVal(Min As Double, Max As Double) As Double
    Randomize
    rndVal = (Max - Min) * Rnd + Min
End Function

'==============================================================================
'현재시간의 1000초의 값을 가지고 guid를 생성한다.(주의/ 빨리 수행되면 guid가 아니다.
'==============================================================================
Public Function getGUID(n진수) As String

Dim val As Double
Dim nVal As String
Dim stack1 As String
Dim Order As Integer

Dim vArray  As Variant
Dim 몫 As Long
Dim 나머지 As Double
Dim stack2 As String

val = CDbl(Now) * 10000000000#

Order = 0

Do While (n진수 ^ Order) < val
    Order = Order + 1
Loop


Do
    몫 = Fix(Trim(val / (n진수 ^ Order)))
    나머지 = val - 몫 * (n진수 ^ Order)
    Order = Order - 1
    stack1 = stack1 & 몫 & "_"
    val = 나머지
Loop While Order >= 0


vArray = Split(stack1, "_")

stack2 = ""

Dim i As Integer

Do While vArray(i) <> ""
    stack2 = stack2 & num2antilog_char(vArray(i))
    i = i + 1
Loop
getGUID = stack2
End Function

'==============================================================================
'guid 함수에서 콜한다.
'==============================================================================
Function num2antilog_char(ByVal C As Integer) As String
    Dim val As String
    Select Case C
        Case 0: val = "0"
        Case 1: val = "1"
        Case 2: val = "2"
        Case 3: val = "3"
        Case 4: val = "4"
        Case 5: val = "5"
        Case 6: val = "6"
        Case 7: val = "7"
        Case 8: val = "8"
        Case 9: val = "9"
        Case 10: val = "A"
        Case 11: val = "B"
        Case 12: val = "C"
        Case 13: val = "D"
        Case 14: val = "E"
        Case 15: val = "F"
        Case 16: val = "G"
        Case 17: val = "H"
        Case 18: val = "I"
        Case 19: val = "J"
        Case 20: val = "K"
        Case 21: val = "L"
        Case 22: val = "M"
        Case 23: val = "N"
        Case 24: val = "O"
        Case 25: val = "P"
        Case 26: val = "Q"
        Case 27: val = "R"
        Case 28: val = "S"
        Case 29: val = "T"
        Case 30: val = "U"
        Case 31: val = "V"
        Case 32: val = "W"
        Case 33: val = "X"
        Case 34: val = "Y"
        Case 35: val = "Z"
        Case 36: val = "a"
        Case 37: val = "b"
        Case 38: val = "c"
        Case 39: val = "d"
        Case 40: val = "e"
        Case 41: val = "f"
        Case 42: val = "g"
        Case 43: val = "h"
        Case 44: val = "i"
        Case 45: val = "j"
        Case 46: val = "k"
        Case 47: val = "l"
        Case 48: val = "m"
        Case 49: val = "n"
        Case 50: val = "o"
        Case 51: val = "p"
        Case 52: val = "q"
        Case 53: val = "r"
        Case 54: val = "s"
        Case 55: val = "t"
        Case 56: val = "u"
        Case 57: val = "v"
        Case 58: val = "w"
        Case 59: val = "x"
        Case 60: val = "y"
        Case 61: val = "z"
        Case Else
            Debug.Assert False
    End Select
num2antilog_char = val
End Function

'==============================================================================
' 완전한 이름의 파일경로와 파일명을 리턴한다.
'==============================================================================
Public Function getPerfectFile(pth As String, fnm As String) As String
    If Right(pth, 1) = "\" Then
        getPerfectFile = pth & fnm
    Else
        getPerfectFile = pth & "\" & fnm
    End If
End Function
Public Function LeftH(ByVal strString As String, ByVal lngLength As Long) As String
'용도: 한글과 영문이 혼합된 경우에 한글은 2바이트, 영문은 1바이트로 계산
    
    LeftH = StrConv(LeftB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
    
End Function

Public Function RightH(ByVal strString As String, lngLength As Long) As String
'용도: 한글과 영문이 혼합된 경우에 한글은 2바이트, 영문은 1바이트로 계산

    RightH = StrConv(RightB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
    
End Function

Public Function MidH(ByVal strString As String, ByVal lngStart As Long, _
        Optional ByVal lngLength) As String
'용도: 한글과 영문이 혼합된 경우에 한글은 2바이트, 영문은 1바이트로 계산
    
    If IsMissing(lngLength) Then
        MidH = StrConv(MidB(StrConv(strString, vbFromUnicode), lngStart), _
            vbUnicode)
    Else
        MidH = StrConv(MidB(StrConv(strString, vbFromUnicode), lngStart, _
            lngLength), vbUnicode)
    End If
            
End Function

Public Function StripNulls(OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = LeftH(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = Trim(OriginalStr)
End Function

'-----------------------------------------------------------
' SUB: UpdateStatus
'
' "Fill" (by percentage) inside the PictureBox and also
' display the percentage filled
'
' IN: [pic] - PictureBox used to bound "fill" region
'     [sngPercent] - Percentage of the shape to fill
'     [fBorderCase] - Indicates whether the percentage
'        specified is a "border case", i.e. exactly 0%
'        or exactly 100%.  Unless fBorderCase is True,
'        the values 0% and 100% will be assumed to be
'        "close" to these values, and 1% and 99% will
'        be used instead.
'
' Notes: Set AutoRedraw property of the PictureBox to True
'        so that the status bar and percentage can be auto-
'        matically repainted if necessary
'-----------------------------------------------------------
'
Public Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase As Boolean = False)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer

    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &H800000 ' dark blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    Dim intPercent
    intPercent = Int(100 * sngPercent + 0.5)
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If

    pic.Refresh
End Sub

Function getParamValue(strCmd As String, strParam As String, strDefault As String) As String
    Dim tmpArr As Variant
    
    Dim strCmdTmp As String
    
    strCmdTmp = Replace(strCmd, "/", "-")
        
    strCmdTmp = Replace(strCmdTmp, "--host", "-h ") 'Sever
    strCmdTmp = Replace(strCmdTmp, "--password", "-p ") 'Password
    strCmdTmp = Replace(strCmdTmp, "--uid", "-u ") 'Uid
    strCmdTmp = Replace(strCmdTmp, "--option", "-o ") 'Option
    strCmdTmp = Replace(strCmdTmp, "--port", "-P ") 'Port
    strCmdTmp = Replace(strCmdTmp, "--database", "-d ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "--ftpaddr", "-x ") 'ftp addr
    strCmdTmp = Replace(strCmdTmp, "--ftpuser", "-y ") 'ftp user
    strCmdTmp = Replace(strCmdTmp, "--ftppass", "-z ") 'ftp pass

    strCmdTmp = Replace(strCmdTmp, "-h", "-h ") 'Sever
    strCmdTmp = Replace(strCmdTmp, "-p", "-p ") 'Password
    strCmdTmp = Replace(strCmdTmp, "-u", "-u ") 'Uid
    strCmdTmp = Replace(strCmdTmp, "-o", "-o ") 'Option
    strCmdTmp = Replace(strCmdTmp, "-P", "-P ") 'Port
    strCmdTmp = Replace(strCmdTmp, "-d", "-d ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "-x", "-x ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "-y", "-y ") 'DataBase
    strCmdTmp = Replace(strCmdTmp, "-z", "-z ") 'DataBase
    
    strCmdTmp = Replace(strCmdTmp, "  ", " ") 'space handling
    strCmdTmp = Replace(strCmdTmp, "  ", " ") 'space handling
    
    tmpArr = Split(strCmdTmp, " ")
    getParamValue = strDefault
    
    Dim strParamTmp As String
    
    strParamTmp = "-" & strParam
    
    Dim i As Long
    For i = 0 To UBound(tmpArr)
        If tmpArr(i) = strParamTmp Then
            getParamValue = tmpArr(i + 1)
            Exit For
        End If
        
    Next
    
End Function

