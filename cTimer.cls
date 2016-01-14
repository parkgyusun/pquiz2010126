VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


'=============================================================================================================
'
' cTimer Class Module
' -------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : April 01, 2000
'
' VB Versions : 5.0 / 6.0
'
' Requires    : modTimer.bas  ( Helper Module )
'
' NOTE        : This module uses the Windows multimedia timer APIs to do a type of subclassing that requires
'               that you terminate this class module before exiting your program.  If you compile your program
'               and run it, you'll not have any problems and don't have to worry about this.  However, if
'               you're running in debug mode and click the STOP button, the class never gets properly
'               terminated.  This will not crash your program or cause any problems directly, but to avoid
'               possible problems with timer events being left open, please close all the open forms in your
'               project to end a debug run instead of clicking the STOP button in Visual Basic.  This will
'               process the class's terminate event.
'
' NOTE        : Use the Enabled property to start and stop this timer class
'
' Description : This class module was designed to take the place of VB's default Timer control.  Unlike the
'               VB standard Timer control, this class module does not require a form to run, allows for more
'               control of the timer,and is accurate down to 1 millisecond on most systems. ( the Min property
'               will return what your system's smallest possible time measurement is )
'
'               This timer class module also doubles as a stop watch as it reports back how many milliseconds
'               have passed since the timer was started through use of the TimeElapsed property.
'
' SEE ALSO    : cTimer_NoSC.cls
'               ( This version of the Timer class module is the subclassing version.  cTimer_NoSC.cls is the
'               No Subclassing version )
'
' Example Use :
'
'  Private Timer1 As New cTimer
'
'  Private Sub Form_Load()
'    Set Timer1 = New cTimer
'    Timer1.Interval = 2000
'    Timer1.Enabled = True
'  End Sub
'
'  Private Sub Form_Click()
'    Debug.Print Timer1.TimeElapsed
'  End Sub
'
'  * IMPORTANT - PUT TIMER RELATED FUNCTIONALITY IN THE TimeProc FUNCTION
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


Private Const TIME_ONESHOT = 0
Private Const TIME_PERIODIC = 1
Private Const TIME_CALLBACK_FUNCTION = 0

Private tTimerID As Long
Private tInterval As Long
Private tResolution As Long
Private tUserData As Long
Private tFlags As Long
Private tEnabled As Boolean
Private tTimerStart As Long
Private tMin As Long
Private tMax As Long

'=============================================================================================================
'                           CLASS INITILIZATIONS
'=============================================================================================================

' Start the class module
Private Sub Class_Initialize()
On Error Resume Next
  
  Dim Caps As TIMECAPS
  
  ' Initialize variables
  tInterval = 0
  tResolution = 0
  tUserData = 0
  tTimerStart = -1
  tFlags = TIME_CALLBACK_FUNCTION Or TIME_PERIODIC
  
  ' Set the smallest measurement of time to 1 millisecond
  timeBeginPeriod 1
  
  ' Get the minimum and maximum timer interval settings
  timeGetDevCaps Caps, Len(Caps)
  tMin = Caps.wPeriodMin
  tMax = Caps.wPeriodMax
  
End Sub

' End the class module
Private Sub Class_Terminate()
On Error Resume Next
  
  ' If a timer is active, kill it
  If tTimerID <> 0 Then
    timeKillEvent tTimerID
    tTimerID = 0
  End If
  
  ' Terminate the call to the "timeBeginPeriod" API
  timeEndPeriod 1
  
End Sub


'=============================================================================================================
'                            CLASS PROPERTIES
'=============================================================================================================


' Returns if the class is currently enabled or not
Public Property Get Enabled() As Boolean
On Error Resume Next
  
  Enabled = tEnabled
  
End Property

' Sets if the class is currently enabled or not
Public Property Let Enabled(ByVal NewValue As Boolean)
On Error Resume Next
  
  ' If enabled, create timer and start it, else destroy the current one
  If NewValue = True Then
    If tEnabled = False And tInterval > 0 And tTimerID = 0 Then
      tTimerID = timeSetEvent(tInterval, tResolution, AddressOf TimeProc, tUserData, tFlags)
      tTimerStart = timeGetTime
    End If
  ElseIf NewValue = False Then
    If tTimerID <> 0 Then
      timeKillEvent tTimerID
      tTimerID = 0
      tTimerStart = -1
    End If
  End If
  
  tEnabled = NewValue
  
End Property

' Returns how much time has past since the timer was started
Public Property Get TimeElapsed() As Long
On Error Resume Next
  
  If tTimerStart = -1 Or tEnabled = False Then TimeElapsed = 0 Else TimeElapsed = Abs(tTimerStart - timeGetTime)
  
End Property

' Returns the currently set interval
Public Property Get Interval() As Long
On Error Resume Next
  
  Interval = tInterval
  
End Property

' Sets the interval value
Public Property Let Interval(ByVal NewValue As Long)
On Error Resume Next
  
  ' Make sure interval isn't greater than the maximum
  If NewValue > tMax Then
    tInterval = tMax
    MsgBox "Interval can not be set greater than the maximum interval value of " & CStr(tMax), vbOKOnly + vbExclamation, "  Invalid Interval"
    
  ' If interval is set to a negative number, make it zero
  ElseIf NewValue < 0 Then
    tInterval = 0
    MsgBox "Interval can not be set to a negative number", vbOKOnly + vbExclamation, "  Invalid Interval"
    If tTimerID <> 0 Then
      timeKillEvent tTimerID
      tTimerID = 0
      tTimerStart = -1
    End If
    tEnabled = False
    
  ' If interval is set to 0, turn off the timer and set the Enabled property to FALSE
  ElseIf NewValue = 0 Then
    tInterval = 0
    If tTimerID <> 0 Then
      timeKillEvent tTimerID
      tTimerID = 0
      tTimerStart = -1
    End If
    tEnabled = False
    
  ' If the interval value is set to something smaller than the Minimum value, reset it
  ElseIf NewValue < tMin Then
    tInterval = tMin
    MsgBox "Interval can not be set less than the minimum interval value of " & CStr(tMin), vbOKOnly + vbExclamation, "  Invalid Interval"
    
  ' Interval is valid, set to new value
  Else
    tInterval = NewValue
    
  End If
  
End Property

' Returns the smallest interval value possible
Public Property Get IntervalMin() As Long
On Error Resume Next
  
  IntervalMin = tMin
  
End Property

' Returns the largest interval value possible
Public Property Get IntervalMax() As Long
On Error Resume Next
  
  IntervalMax = tMax
  
End Property

