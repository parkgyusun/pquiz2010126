VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmHint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "힌트 -더블클릭하여 닫기-"
   ClientHeight    =   5040
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   13365
   DrawMode        =   10  'Mask Pen
   ForeColor       =   &H0000FF00&
   Icon            =   "frmHint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   360
      ScaleHeight     =   2220
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   405
      Width           =   4785
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   45
      ExtentX         =   7938
      ExtentY         =   3969
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   30
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   53
   End
   Begin VB.Menu mnuPop1 
      Caption         =   "mnuPop1"
      Begin VB.Menu mnuCopy 
         Caption         =   "복사"
      End
      Begin VB.Menu mnuTTS 
         Caption         =   "TTS실행"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "새로고침"
      End
   End
End
Attribute VB_Name = "frmHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim bonDraw As Boolean
Dim clipText As String
Dim gSetLang As String '변수 따다 써서 쓸데없이 정의함.
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Public frmMain As frmMain

'Private Sub doPrint()
'    picMain.cls
'    Me.picMain.Print autoCRLF(refCquiz.hint, picMain.Width, picMain)
'End Sub


Private Sub cmdClose_Click()
Me.Hide
End Sub


Private Sub Form_DblClick()
Me.Hide
End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
If keycode = Asc(" ") Or keycode = vbKeyReturn Or keycode = vbKeyEscape Then
    Me.Hide
Else
    frmQuiz.Form_KeyDown keycode, Shift
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmQuiz.Form_KeyPress KeyAscii
End Sub

Private Sub Form_KeyUp(keycode As Integer, Shift As Integer)
    frmQuiz.Form_KeyPress keycode ', Shift
End Sub

Private Sub Form_Load()
'wb.Left = -50
    mnuPop1.Visible = False

    Me.ScaleMode = vbTwips
    
    picMain.Width = Screen.Width - picMain.Left
    picMain.Height = Screen.Height - picMain.Top

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmHint.Visible = False
    Call Me.Hide
End Sub

'Private Sub Form_Resize()
'    picMain.Width = Me.Width
'    doPrint
'End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText refCquiz.hint
End Sub

Private Sub mnuRefresh_Click()
frmQuiz.cmdHintRefresh
End Sub

Private Sub mnuTTS_Click()
Dim isHan As Boolean

Dim clipText2 As String
clipText2 = Trim(refCquiz.hint)
clipText2 = Replace(clipText2, vbCrLf, " ")
clipText2 = Replace(clipText2, "   ", " ")
clipText2 = Replace(clipText2, "  ", " ")
clipText2 = LeftH(clipText2, 200)

Dim nChkLenB As Long
Dim nChkLenH As Long

nChkLenB = LenB(StrConv(clipText2, vbFromUnicode))
nChkLenH = Len(clipText2)

If (nChkLenB <> nChkLenH) Then
    isHan = True
Else
    
    Dim char_first As String
    
    char_first = Left(clipText2, 1)
    
    Dim asc_first As Integer
    
    asc_first = Asc(char_first)
    
    If asc_first >= Asc("0") And asc_first <= Asc("9") Then
        isHan = True
    Else
        isHan = False
    End If
    
End If

Dim URL As String

clipText = Replace(clipText2, "@", "")

mnuTTS0_Click

'If isHan Then
'    'wb.Navigate2 "http://www.hcilab.co.kr/demo/pttsnet_proc.asp?language=0&voice=0&contents=" + clipText2
'
'    'Call make_html_tts_ko(clipText2, "tmp\$tts_ko" + setRndName() + ".htm")
'    'toUTF8 App.Path + "\tts_ko.htm"
'    'frmMain.wb3.Navigate2 App.Path + "\tmp\$tts_ko" + getRndName() + ".htm"
'    URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=ko-KR&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
'    frmMain.wb3.Navigate2 URL
'Else
'    'wb.Navigate2 "http://www.hcilab.co.kr/demo/pttsnet_proc.asp?language=1&voice=0&contents=" + clipText2
'
'    'Call make_html_tts_eng(clipText2, "tmp\$tts_eng" + setRndName() + ".htm")
'    'frmMain.wb3.Navigate2 App.Path + "\tmp\$tts_eng" + getRndName() + ".htm"
'
'    URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
'    frmMain.wb3.Navigate2 URL
'End If

frmMain.wb3.Tag = clipText2

End Sub

Public Sub mnuTTS0_Click()

Dim sLang As String

sLang = "ko-KR"

Dim clipText2 As String

clipText2 = Trim(clipText)
If clipText2 = "" Then Exit Sub

clipText2 = Replace(clipText2, "(A)", "")
clipText2 = Replace(clipText2, "(B)", "")
clipText2 = Replace(clipText2, "(C)", "")
clipText2 = Replace(clipText2, "(D)", "")
clipText2 = Replace(clipText2, vbCrLf, " ")
clipText2 = Replace(clipText2, "   ", " ")
clipText2 = Replace(clipText2, "  ", " ")

clipText2 = Replace(clipText2, "______", "_")
clipText2 = Replace(clipText2, "_____", "_")
clipText2 = Replace(clipText2, "____", "_")
clipText2 = Replace(clipText2, "___", "_")
clipText2 = Replace(clipText2, "__", "_")
'clipText2 = LeftH(clipText2, 200)

Dim nChkLenB As Long
Dim nChkLenH As Long

If gSetLang = "" Then
    nChkLenB = LenB(StrConv(clipText2, vbFromUnicode))
    nChkLenH = Len(clipText2)
    
    If (nChkLenB <> nChkLenH) Then
        'isHan = True
        sLang = "ko-KR" '?
    Else
        Dim char_first As String
        
        char_first = Left(clipText2, 1)
        
        Dim asc_first As Integer
        
        asc_first = Asc(char_first)
        
        If asc_first >= Asc("0") And asc_first <= Asc("9") Then
            sLang = "ko-KR" 'isHan = True
        Else
            'isHan = False
            sLang = "en-US"
        End If
        
    End If
Else
    sLang = gSetLang
End If

Dim tailSec As String
tailSec = Format(Now, "ss")

Static idx   As Long

idx = idx + 1
idx = idx Mod 3

Dim URL As String, url_hash As String

If sLang = "ko-KR" Then
    
    'Call make_html_tts_ko(clipText2, "tmp\$" + setRndName() & "tts_ko.htm")
    
    '일본어URL인지 확인한다.
    If IsJAJP(clipText2) Then
        'URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=ja&q=" + URLEncodeUTF8(clipText2)
        URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=ja-JP&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
        'URL = Module1.TTS_GLOBAL_URL
        'Call Module1.toclipboard(clipText2)
    Else
        'URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=ko&q=" + URLEncodeUTF8(clipText2)
        URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=ko-KR&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
        'URL = Module1.TTS_GLOBAL_URL
        'Call Module1.toclipboard(clipText2)
    End If
    Call TTS_PLAY_SMART3(URL)
    
'    Dim tmpPath As String
'    tmpPath = App.Path + "\tmp\$" & getRndName() & "tts_ko.htm"
        'Call toUTF8(tmpPath)
        ''frmMain.wb(idx).Navigate2 tmpPath
        'Call TTS_PLAY_SMART(tmpPath)
ElseIf sLang = "en-US" Then
    Dim bSuccessFromDaum As Boolean
    
    If 0 = InStr(clipText2, " ") Then
        If TTS_PLAY_SMART4(clipText2) Then
            bSuccessFromDaum = True
        End If
    End If
    
    If bSuccessFromDaum = False Then
        'URL = "http://translate.google.com/translate_tts?ie=UTF-8&tl=en&q=" + URLEncodeUTF8(clipText2)
        URL = "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" + URLEncodeUTF8(clipText2)
        'URL = Module1.TTS_GLOBAL_URL
        'Call Module1.toclipboard(clipText2)
        Call TTS_PLAY_SMART3(URL)
    End If
    
ElseIf sLang = "zh-CN" Then
    'frmMain.wb(idx).Navigate2 "http://translate.google.com/translate_tts?ie=UTF-8&tl=zh&q=" + clipText2
    'Call TTS_PLAY_SMART3("http://translate.google.com/translate_tts?ie=UTF-8&tl=zh&q=" + URLEncodeUTF8(clipText2))
    Call TTS_PLAY_SMART3("https://translate.google.com/translate_tts?ie=UTF-8&tl=zh-CN&client=tw-ob&q=" + URLEncodeUTF8(clipText2))
'    Call TTS_PLAY_SMART3(Module1.TTS_GLOBAL_URL)
'    Call Module1.toclipboard(clipText2)
ElseIf sLang = "ja-JP" Then
    'frmMain.wb(idx).Navigate2 "http://translate.google.com/translate_tts?ie=UTF-8&tl=ja&q=" + clipText2
    'Call TTS_PLAY_SMART3("http://translate.google.com/translate_tts?ie=UTF-8&tl=ja&q=" + URLEncodeUTF8(clipText2))
    Call TTS_PLAY_SMART3("https://translate.google.com/translate_tts?ie=UTF-8&tl=ja-JP&client=tw-ob&q=" + URLEncodeUTF8(clipText2))
'    Call TTS_PLAY_SMART3(Module1.TTS_GLOBAL_URL)
'    Call Module1.toclipboard(clipText2)
End If

frmMain.wb2.Tag = clipText2
#If MYAGENTUSE_ON Then
If Not frmMain.Character Is Nothing Then
    frmMain.Character.Balloon.Style = frmMain.Character.Balloon.Style And (Not 4) '말풍선뜬거 계속 떠있게 하기
End If
#End If


'0.5초 후에 포커싱 하기
'TmrAfterTTS_focus.Enabled = True
'TmrAfterTTS_focus.Interval = 100

End Sub

Public Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean

   DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS

End Function

Function searchLineInFile(ByVal f As String, ByVal search As String) As String

  Dim FileHandle As Integer
  FileHandle = FreeFile
  Open f For Input As #FileHandle

  Do While Not EOF(FileHandle)        ' Loop until end of file
    Line Input #FileHandle, searchLineInFile  ' Read line into variable
    If 0 < InStr(searchLineInFile, search) Then
        Exit Do
    End If
    ' Your code here
  Loop

  Close #FileHandle
  

End Function


Function searchFromUrlToContent(ByVal FileName_tmp As String, ByVal URL As String, ByVal firstPosChar As String, ByVal lastPosChar As String, ByVal search_key As String) As String

'Dim URL As String
        'URL = "http://small.dic.daum.net/search.do?dic=eng&q=" & str
        
        Call DownloadFile(URL, FileName_tmp)
        
        DoEvents
        
        Call SleepEx(100, True)
        
        If FileExists(FileName_tmp) Then
            'wordid=ekw000167646&q=terrify
            Dim file_content As String
            file_content = searchLineInFile(FileName_tmp, search_key)
            'file_content = StrConv(StrConv(file_content, vbUnicode), vbFromUnicode)
            
            Dim pos1 As Long
            'Dim firstPosChar As String
            'firstPosChar = "URL="
            Dim firstPostCharLeng As Long
            firstPostCharLeng = Len(firstPosChar)
            'Dim lastPosChar As String
            'lastPosChar = """"
            
            pos1 = InStr(file_content, firstPosChar)
            Dim pos2 As Long
            pos2 = InStr(pos1 + firstPostCharLeng, file_content, lastPosChar)
            
            Dim keyword As String
            
            If 0 < pos2 - pos1 - firstPostCharLeng Then
                searchFromUrlToContent = Mid(file_content, pos1 + firstPostCharLeng, pos2 - pos1 - firstPostCharLeng)
            Else
                Debug.Print "not found"
            End If
            
            'http://dic.daum.net/word/view.do?wordid=ekw000167646&q=terrify
        End If
        

End Function

'검색 성공후 앞에 30자리 이전은 모두 자른다.
Function searchLineInFileAbout30(ByVal f As String, ByVal search As String) As String

  Dim FileHandle As Integer
  Dim retVal As String
  Dim pos1 As Long
  
  FileHandle = FreeFile
  Open f For Input As #FileHandle

  Do While Not EOF(FileHandle)        ' Loop until end of file
    Line Input #FileHandle, retVal  ' Read line into variable
    
    pos1 = InStr(retVal, search)
    
    If 0 < pos1 Then
        
        If 30 < pos1 Then
            retVal = Mid(retVal, pos1 - 30)
        End If
        
        searchLineInFileAbout30 = retVal
        Exit Do
    End If
    ' Your code here
  Loop

  Close #FileHandle
  

End Function



Function searchFromFileToContent(ByVal FileName_tmp As String, ByVal firstPosChar As String, ByVal lastPosChar As String, ByVal search_key As String) As String

    If FileExists(FileName_tmp) Then
        'wordid=ekw000167646&q=terrify
        Dim file_content As String
        file_content = searchLineInFileAbout30(FileName_tmp, search_key)
        'file_content = StrConv(StrConv(file_content, vbUnicode), vbFromUnicode)
        
        Dim pos1 As Long
        'Dim firstPosChar As String
        'firstPosChar = "URL="
        Dim firstPostCharLeng As Long
        firstPostCharLeng = Len(firstPosChar)
        'Dim lastPosChar As String
        'lastPosChar = """"
        
        pos1 = InStr(file_content, firstPosChar)
        Dim pos2 As Long
        pos2 = InStr(pos1 + firstPostCharLeng, file_content, lastPosChar)
        
        Dim keyword As String
        
        If 0 < pos2 - pos1 - firstPostCharLeng Then
            searchFromFileToContent = Mid(file_content, pos1 + firstPostCharLeng, pos2 - pos1 - firstPostCharLeng)
        Else
            Debug.Print "not found"
        End If
        
        'http://dic.daum.net/word/view.do?wordid=ekw000167646&q=terrify
    End If
        

End Function



Function TTS_PLAY_SMART4(ByVal str As String) As Boolean
    Dim FileName_tmp As String
    Dim FileName_tmp2 As String
    Dim FileName_ok As String
    Dim FileName_error As String
    
    FileName_tmp = FileSystem.CurDir & "\cache\temp_daum" + str + ".htm"
    FileName_tmp2 = FileSystem.CurDir & "\cache\temp_daum" + str + "2.htm"
    
    FileName_ok = FileSystem.CurDir & "\cache\" & str & ".mp3"
    FileName_error = FileSystem.CurDir & "\cache\~" & str & ".mp3"
    
    Dim okMP3 As Boolean
    
    okMP3 = False
    If FileExists(FileName_error) Then
        TTS_PLAY_SMART4 = False
        Exit Function
    End If
    
    If FileExists(FileName_ok) Then
        WindowsMediaPlayer1.settings.mute = False
        WindowsMediaPlayer1.settings.volume = 100
        WindowsMediaPlayer1.URL = FileName_ok
        TTS_PLAY_SMART4 = True
    Else
    
        Dim URL As String
        URL = "http://small.dic.daum.net/search.do?dic=eng&q=" & str
        
        Dim str1 As String
        str1 = searchFromUrlToContent(FileName_tmp, URL, "URL=", """", "wordid=")
        
        If str1 <> "" And InStr(str1, "wordid=") = 0 Then
            str1 = searchFromFileToContent(FileName_tmp, "<a href=""", """", "wordid=")
        Else
            Debug.Print "test"
        End If
        
        If (str1 <> "") Then
            DeleteFile (FileName_tmp)
            
            Dim str2 As String
            str2 = searchFromUrlToContent(FileName_tmp2, "http://dic.daum.net" + str1, "event, '", "'", "mainPlayer.play")
            If (str2 <> "") Then
                DeleteFile (FileName_tmp2)
                
                Call DownloadFile(str2, FileName_ok)
                
                DoEvents
                
                Call SleepEx(100, True)
                
                If FileExists(FileName_ok) Then
                    TTS_PLAY_SMART4 = True
                    WindowsMediaPlayer1.settings.mute = False
                    WindowsMediaPlayer1.settings.volume = 100
                    WindowsMediaPlayer1.URL = FileName_ok
                Else
                    Debug.Print "file not exits1?"
                End If
            Else
                        Debug.Print "file not exits2?"
            End If
        Else
                    Debug.Print "file not exits3?"
        End If
    End If
    
End Function

Private Function FolderExists(sFullPath As String) As Boolean
    Dim myFSO As Object
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(sFullPath)
End Function




Sub TTS_PLAY_SMART3(ByVal URL As String)
    If TTS_PLAY_SMART2(URL) = False Then
        Call TTS_PLAY_SMART(URL)
    End If
End Sub

'mp3 파일에도 없고. URL을 호출하여도 응답이 없는 경우는 직접 해당 URL을 메인 웹브라우저에 표시한다.
Sub TTS_PLAY_SMART(ByVal URL As String)
'Call wbQuizMain.Navigate(URL)'그런데 의미 없어 주석 처리함
End Sub

Function TTS_PLAY_SMART2(ByVal URL As String) As Boolean '성공하면 true를 리턴

    Dim sha256 As New CSHA256
    Dim url_hash As String
    url_hash = sha256.sha256(URL) + ".mp3"

    Set sha256 = Nothing

    If FolderExists(FileSystem.CurDir & "\cache") = False Then
        On Error Resume Next
        FileSystem.MkDir (FileSystem.CurDir & "\cache") '//err.Number 75 가 발생되면서 Path/File access error 오류가 발생된다.
        If err.Number = 75 Then
            MsgBox "[" & FileSystem.CurDir & "\cache] 폴더를 수동으로 만드세요.", vbExclamation + vbOKOnly
        End If
    End If

    Dim FileName As String
    FileName = FileSystem.CurDir & "\cache\" & url_hash

    Dim okMP3 As Boolean
    
    okMP3 = False
    If FileExists(FileName) Then
        If 0 < FileSystem.FileLen(FileName) Then
            If isHTML(FileName) Then
                Call DeleteFile(FileName)
            Else
                okMP3 = True
            End If
        Else
            Call DeleteFile(FileName)
        End If
    End If
    

    If okMP3 Then
        'mp3 play
       'Call mciSendString(CommandString, vbNullString, 0, 0)
       WindowsMediaPlayer1.settings.mute = False
       WindowsMediaPlayer1.settings.volume = 100
       WindowsMediaPlayer1.URL = FileName

       TTS_PLAY_SMART2 = True
    Else
       'http://stackoverflow.com/questions/1976152/
       Call DownloadFile(URL, FileName)
       '"C:\Program Files\Internet Explorer\iexplore.exe"
       'Call Shell("C:\Program Files\Internet Explorer\iexplore.exe" & " " & url)

       DoEvents
       
       TTS_PLAY_SMART2 = False
       
       If FileExists(FileName) Then
            If 0 < FileSystem.FileLen(FileName) Then
                If isHTML(FileName) Then
                    Call DeleteFile(FileName)
                Else
                    TTS_PLAY_SMART2 = True
                End If
            Else
                Call DeleteFile(FileName)
            End If
       End If
       
       If TTS_PLAY_SMART2 Then
            WindowsMediaPlayer1.settings.mute = False
            WindowsMediaPlayer1.settings.volume = 100
            WindowsMediaPlayer1.URL = FileName
       End If
    End If
    
End Function

Function isHTML(ByVal filepath) As Boolean
    Dim filetext As String
    Dim handle As Integer
    Dim Pos As Long
    
    handle = FreeFile
    On Error Resume Next
    Open filepath For Input As #handle
               
        Do While Not EOF(handle)
          Line Input #handle, filetext
          If filetext Like "*<html*" Then
            Pos = 1
            Exit Do
          End If
        Loop
        
    Close #handle
    On Error GoTo 0
    '
    If 0 < Pos Then
     isHTML = True
    End If
    
End Function


Public Function FileExists(ByVal fname As String) As Boolean
    Dim TheFile As String
    Dim Results As String
    
    TheFile = fname
    Results = Dir$(TheFile)
    
    If Len(Results) = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function


Function IsJAJP(ByVal txt As String) As Boolean
    '%aa%a1~%abf6까지가 일본어
    Dim str1 As String
    str1 = "&H" & Replace(Encode(Left(txt, 1)), "%", "")
    
    Dim v As Long
    
    v = CLng(str1)
    
    If CLng("&Haaa1") <= v And v <= CLng("&Habf6") Then
        IsJAJP = True
    End If
    
End Function



Private Sub picMain_DblClick()
    Me.Hide
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button And vbLeftButton) Then
    bonDraw = True
    'Me.DrawMode = 12
    Me.DrawWidth = cFTP.Profile.FontSize / 5
    'Me.ForeColor = vbMagenta
    
    PicMainDrawModeForeColor
    
    Me.Line (x, y)-(x, y)
Else
    Me.PopupMenu mnuPop1
End If
End Sub

Sub PicMainDrawModeForeColor()

    If 128 < GrayScaleFromColor(cFTP.Profile.PenColor) Then
        Me.DrawMode = 12
        Me.ForeColor = frmQuiz.Shape2.FillColor
    Else
        Me.DrawMode = 13
        Me.ForeColor = frmQuiz.Shape2.FillColor Xor &HFFFFFF
    End If

End Sub

Private Function GrayScaleFromColor(ByVal lPenColor As Long) As Long

    Dim red As Long, green As Long, blue As Long, Sum As Long
    
    red = RedFromRGB(lPenColor)
    green = GreenFromRGB(lPenColor)
    blue = BlueFromRGB(lPenColor)

    GrayScaleFromColor = Fix(red * 0.3 + green * 0.6 + 0.1 * blue)

End Function



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static preX As Single
Static preY As Single

Dim Dist As Single


If bonDraw Then
    
    If preX = 0 And preY = 0 Then
        preX = x
        preY = y
        Exit Sub
    End If
    Dist = ((x - preX) ^ 2 + (y - preY) ^ 2) ^ 0.5 '[twip]
    
    If Dist > Me.DrawWidth * Screen.TwipsPerPixelX Then
    
        Me.Line -(x, y)
        preX = x
        preY = y
    End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
bonDraw = False
Me.DrawMode = 13
End Sub

'Private Sub wb_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'    If Len(frmMain.wb3.Tag) > 0 Then
'        If Not frmMain.Character Is Nothing Then
'            frmMain.Character.Speak frmMain.wb3.Tag
'        End If
'    End If
'End Sub

