VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

'
' Defaults:
' DelayInitial   = 500  (1/2 sec)
' DelayAutoPopup = 5000  (5 secs)
' DelayReshow    = 100 (1/10 sec)
' MaxTipWidth    = 0
' all Margins    = 0

Private mnlgHwndTT As Long
Private mnlgMaxTip As Long
Private m_lStyle As Long
Private mvarAlpha As Long

' -- Enums --
Enum TooltipStyle
   ttNormal = &H0
   ttBaloon = &H1
End Enum


Public Function Create(ByRef frm As Form) As Boolean
  
    If (mnlgHwndTT = 0) Then
        Call InitCommonControls
        
        '
        ' The hwndParent param lets the tooltip window
        ' be owned by the specified form and be destroyed
        ' along with it. We'll cleanup in Class_Terminate anyway.
        ' No WS_EX_TOPMOST or TTS_ALWAYSTIP per Win95 UI rules.
        '
        
        mnlgHwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP, 0, 0, 0, 0, frm.hwnd, 0, App.hInstance, ByVal 0)
'        mnlgHwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_NOFADE Or TTS_NOANIMATE, 0, 0, 0, 0, frm.hWnd, 0, App.hInstance, ByVal 0)
        
    End If
    
    Create = CBool(mnlgHwndTT)

End Function

Private Sub Class_Terminate()

    If mnlgHwndTT > 0 Then Call DestroyWindow(mnlgHwndTT)
    
End Sub
Public Function AddTool(ByRef ctrl As Control, Optional ByVal strText As String) As Boolean
Dim ti As TOOLINFO
  
    If (mnlgHwndTT = 0) Then Exit Function
    
    
    
    If (fGetToolInfo(ctrl.hwnd, ti) = False) Then
        'With ti
            ti.cbSize = Len(ti)
            '
            ' TTF_IDISHWND must be set to tell the tooltip
            ' control to retrieve the control's rect from
            ' it's hWnd specified in uId.
            '
            ti.uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            ti.hwnd = ctrl.Container.hwnd
            ti.uId = ctrl.hwnd
            
            If Len(strText) > 0 Then
                ti.lpszText = strText
'            Else
'                .lpszText = "Tool" & ToolCount + 1
            End If
            
            '
            ' Maintain the maximun tip text
            ' length for fGetToolInfo.
            '
            mnlgMaxTip = fMax(mnlgMaxTip, Len(ti.lpszText) + 1)
        'End With
        
        '
        ' Returns 1 on success, 0 on failure
        '
        AddTool = SendMessageT(mnlgHwndTT, TTM_ADDTOOL, 0, ti)
    End If
  
End Function

Private Function fMax(ByVal lngParm1 As Long, ByVal lngParm2 As Long) As Long
    '
    ' Returns the larger of the two values.
    '
    If lngParm1 > lngParm2 Then
        fMax = lngParm1
    Else
        fMax = lngParm2
    End If
    
End Function
Public Function RemoveTool(ByRef ctrl As Control) As Boolean
Dim ti As TOOLINFO

    If (mnlgHwndTT = 0) Then Exit Function
    
    If fGetToolInfo(ctrl.hwnd, ti) Then
        Call SendMessageT(mnlgHwndTT, TTM_DELTOOL, 0, ti)
        RemoveTool = True
    End If

End Function
Public Property Get BackColor() As OLE_COLOR
    
    If (mnlgHwndTT = 0) Then Exit Property
    
    '
    ' OLE_COLOR is defined in stdole2.tlb
    '
    BackColor = SendMessageT(mnlgHwndTT, TTM_GETTIPBKCOLOR, 0, 0)
    
End Property
Public Property Let BackColor(clr As OLE_COLOR)
    
    If (mnlgHwndTT = 0) Then Exit Property
    Call SendMessageT(mnlgHwndTT, TTM_SETTIPBKCOLOR, clr, 0)

End Property
Public Property Get DelayTime(dwType As ttDelayTimeConstants) As Long
    
    If (mnlgHwndTT = 0) Then Exit Property
    DelayTime = SendMessageT(mnlgHwndTT, TTM_GETDELAYTIME, (dwType And ttDelayMask), 0&)
    
End Property
Public Property Let DelayTime(dwType As ttDelayTimeConstants, dwMilliSecs As Long)
    
    If (mnlgHwndTT = 0) Then Exit Property
    Call SendMessageT(mnlgHwndTT, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs)  ' no rtn val

End Property
Public Property Get ForeColor() As OLE_COLOR
    
    If (mnlgHwndTT = 0) Then Exit Property
    ForeColor = SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, 0, 0)

End Property
Public Property Let ForeColor(clr As OLE_COLOR)
    
    If (mnlgHwndTT = 0) Then Exit Property
    Call SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, clr, 0)   ' no rtn val

End Property
Public Property Get hwnd() As Long

    hwnd = mnlgHwndTT
    
End Property
Public Property Get Margin(dwType As ttMarginConstants) As Long
Dim rc As RECT

    If (mnlgHwndTT = 0) Then Exit Property
    
    Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
    
    Select Case dwType
        Case ttMarginLeft
            Margin = rc.Left
        
        Case ttMarginTop
            Margin = rc.Top
        
        Case ttMarginRight
            Margin = rc.Right
        
        Case ttMarginBottom
            Margin = rc.Bottom
    End Select

End Property
Public Property Let Margin(dwType As ttMarginConstants, cPixels As Long)
Dim rc As RECT
    
    If (mnlgHwndTT = 0) Then Exit Property
    
    Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
    
    Select Case dwType
        Case ttMarginLeft
            rc.Left = cPixels
            
        Case ttMarginTop
            rc.Top = cPixels
            
        Case ttMarginRight
            rc.Right = cPixels
            
        Case ttMarginBottom
            rc.Bottom = cPixels
    End Select
    
    Call SendMessageT(mnlgHwndTT, TTM_SETMARGIN, 0, rc)

End Property
Public Property Get MaxTipWidth() As Long

    '
    ' If MaxTipWidth is -1, there is no word wrapping and
    ' text control characters are printed and not
    ' evaluated (i.e. a vbCrLf shows up as "||")
    '
    If (mnlgHwndTT = 0) Then Exit Property
    MaxTipWidth = fLowWord(SendMessageT(mnlgHwndTT, TTM_GETMAXTIPWIDTH, 0, 0))

End Property
Private Function fLowWord(ByVal lngValue As Long) As Integer
    '
    ' Returns the low-order word from a 32-bit value.
    '
    Call MoveMemory(fLowWord, lngValue, 2)
    
End Function
Public Property Let MaxTipWidth(ByVal lngWidth As Long)
    
    '
    ' If MaxTipWidth is -1, there is no word wrapping and
    ' text control characters are printed and not
    ' evaluated (i.e. a vbCrLf shows up as "||")
    '
    If mnlgHwndTT = 0 Then Exit Property
    If lngWidth < 1 Then lngWidth = -1
    Call SendMessageT(mnlgHwndTT, TTM_SETMAXTIPWIDTH, 0, lngWidth)

End Property
Public Property Get ToolCount() As Long
    
    If (mnlgHwndTT = 0) Then Exit Property
    ToolCount = SendMessageT(mnlgHwndTT, TTM_GETTOOLCOUNT, 0, 0)

End Property
Public Property Get ToolTipHandle() As Long
    
    ToolTipHandle = mnlgHwndTT

End Property

Public Property Get ToolText(ByRef ctrl As Control) As String
Dim ti As TOOLINFO
  
    If (mnlgHwndTT = 0) Then Exit Property
    
    If fGetToolInfo(ctrl.hwnd, ti, True) Then
        ToolText = fGetStrFromBuffer(ti.lpszText)
    End If

End Property
Private Function fGetStrFromBuffer(ByVal strValue As String) As String
    
    If InStr(strValue, vbNullChar) Then
        fGetStrFromBuffer = Left$(strValue, InStr(strValue, vbNullChar) - 1)
    Else
        '
        ' If strValue had no null char, the Left$ function
        ' above would rtn a zero length string ("").
        '
        fGetStrFromBuffer = strValue
    End If
    
End Function
Public Property Let ToolText(ByRef ctrl As Control, ByVal strText As String)
Dim ti As TOOLINFO
  
    If (mnlgHwndTT = 0) Then Exit Property
    
    If fGetToolInfo(ctrl.hwnd, ti) Then
        ti.lpszText = strText
        mnlgMaxTip = fMax(mnlgMaxTip, Len(strText) + 1)
        '
        ' The tooltip won't appear for the control
        ' if lpszText is an empty string
        '
        Call SendMessageT(mnlgHwndTT, TTM_UPDATETIPTEXT, 0, ti)
    End If

End Property
Private Function fIsWindow(ByRef ctrl As Control) As Boolean
    
    On Error GoTo ErrorHandler
    fIsWindow = CBool(ctrl.hwnd)
    
ErrorHandler:
End Function
Private Function fGetToolInfo(ByVal lnghwndTool As Long, ti As TOOLINFO, _
    Optional fGetText As Boolean = False) As Boolean
  
Dim nItems As Long
Dim i      As Integer

    ti.cbSize = Len(ti)
    If fGetText Then ti.lpszText = String$(mnlgMaxTip, 0)
    
    nItems = ToolCount
    
    For i = 0 To nItems - 1
        '
        ' Returns 1 on success, 0 on failure.
        '
        If SendMessageT(mnlgHwndTT, TTM_ENUMTOOLS, (i), ti) Then
            If (lnghwndTool = ti.uId) Then
                fGetToolInfo = True
                Exit Function
            End If
        End If
    Next

End Function

Public Property Get Style() As TooltipStyle
   Style = m_lStyle
End Property

Public Property Let Style(ByVal Value As TooltipStyle)
Dim lStyle As Long

   m_lStyle = Value
   
   lStyle = GetWindowLong(mnlgHwndTT, GWL_STYLE)
   
   If Value = ttBaloon Then
      lStyle = (lStyle And Not WS_BORDER) Or TTS_BALLOON
   Else
      lStyle = (lStyle Or WS_BORDER) And Not TTS_BALLOON
   End If
   
   SetWindowLong mnlgHwndTT, GWL_STYLE, lStyle
   
   SendMessageT mnlgHwndTT, TTM_UPDATE, 0&, ByVal 0&

End Property


Public Property Let Transparency(Alpha As Long)

    Dim lOldStyle As Long
    Dim bTrans As Byte ' The level of transparency (0 - 255)

    bTrans = Alpha
    mvarAlpha = Alpha
    lOldStyle = GetWindowLong(mnlgHwndTT, GWL_EXSTYLE)
    SetWindowLong mnlgHwndTT, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes mnlgHwndTT, 0, bTrans, LWA_ALPHA
    
End Property

Public Property Get Transparency() As Long
    Transparency = mvarAlpha
End Property

Public Sub setTitle(sTitle As String, icon As Long)
    ' Set the tooltip title
    SendMessageT mnlgHwndTT, TTM_SETTITLE, icon, ByVal sTitle

End Sub






