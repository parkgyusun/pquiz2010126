VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "putool "
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox PB 
      Align           =   2  '아래 맞춤
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      FillColor       =   &H00FF0000&
      Height          =   384
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   5640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1380
      Width           =   5700
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소"
      Height          =   420
      Left            =   4500
      TabIndex        =   4
      Top             =   90
      Width           =   1050
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   90
      Top             =   765
   End
   Begin VB.Label Label10 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3330
      TabIndex        =   3
      Top             =   585
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   945
      Width           =   5895
   End
   Begin VB.Label Label11 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "Time Left: 00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   585
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "다운로드중..."
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4290
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents ctlFTP As clsFTP
Attribute ctlFTP.VB_VarHelpID = -1
'Dim WithEvents pZ As PSZLib

Dim WithEvents ZF As Cls_GetFileType
Attribute ZF.VB_VarHelpID = -1

Private TransferRate                    As Single
Private BeginTransfer                   As Single


Private Sub Command1_Click()
Call Show_ZipContents2(App.Path & "\" & "pquiz.zip", App.Path)
End Sub

Private Sub cmdCancel_Click()
gbZipStop = True
Unload Me
End Sub

Private Sub ctlFTP_FileTransferProgress(ByVal lCurrentBytes As Long, ByVal lTotalBytes As Long)


On Error Resume Next
Dim j As Long
Dim j2 As Long
Dim preTimer As Double


Static pretotalbytes As Long
Static prePercent As Long

If preTimer + 1 > Timer Then
    Exit Sub
Else
    preTimer = Timer
End If


If lTotalBytes <> pretotalbytes Then
    BeginTransfer = Timer
    pretotalbytes = lTotalBytes
    prePercent = 0
    UpdateStatus PB, 0, True
End If



TransferRate = Format(Int(lCurrentBytes / (Timer - BeginTransfer)) / 1000, "####.00")
'    lTotalBytes = lTotalBytes
'    PB.Min = 0
  
  
  
  DoEvents
        PB.Tag = lCurrentBytes
        
        If prePercent + 1 < lCurrentBytes / lTotalBytes * 100 Then
            UpdateStatus PB, lCurrentBytes / lTotalBytes, False
            prePercent = lCurrentBytes / lTotalBytes * 100
        End If
        
        
        
        
        j2 = PB.Tag \ 1024
        j = PB.Tag
        DoEvents
        PB.ToolTipText = PB.Tag & " Bytes of " & lTotalBytes & " Bytes Transfered"
        DoEvents
        Label7.Caption = PB.Tag \ 1024 & " KB of " & lTotalBytes \ 1024 & " KB Transfered"
        DoEvents
'        Label9.Caption = Format$(CLng((j / lTotalBytes) * 100)) + "%"
'        DoEvents
        Label10.Caption = Format(TransferRate, "#0.#0#") & " Kbps"
        Label11.Caption = "Time Left: " & ConvertTime(Int(((lTotalBytes - PB.Tag) / 1024) / TransferRate))


End Sub

Private Sub Form_Load()
Set ZF = New Cls_GetFileType

Me.Caption = "pupdate   ver." & App.Major & "." & App.Minor & "." & App.Revision
Set ctlFTP = New clsFTP
End Sub


Public Sub CheckFtp(ByVal cmd As String)

Dim addr As String
Dim usr As String
Dim pass As String

Dim vData As Variant
Dim ocmd As String '오리지날 cmd

ocmd = cmd

cmd = LCase(cmd)
'If InStr(cmd, "/ftp") > 0 Then
'    cmd = Trim(Replace(cmd, "/ftp ", "", 1, 1))
'
'    vData = Split(cmd, " ")
'
'    If UBound(vData) >= 2 Then
'        addr = vData(0)
'        usr = vData(1)
'        pass = vData(2)
'    ElseIf UBound(vData) = 1 Then
'        addr = vData(0)
'        usr = vData(1)
'        pass = "user"
'
'    ElseIf UBound(vData) = 0 Then
'        addr = vData(0)
'        usr = "pquiz"
'        pass = "user"
'
'    End If
'
'Else
        addr = getParamValue(cmd, "x", getParamValue(cmd, "h", "59.150.9.39"))
        usr = getParamValue(cmd, "y", "pquiz")
        pass = getParamValue(cmd, "z", "user")

'End If


Label1.Caption = "준비"

'------------------------------------------------------------------------------
'FTP를 통해서 업데이트기초정보(pupdate.zip) 파일을 다운로드한다.(서버:pupdate.ds 로컬:pupdate.dl 압축:pupdate.zip==>pupdate.ds)
'------------------------------------------------------------------------------
If ctlFTP.FTPDownloadFile(addr, usr, pass, "pupdate.zip", getPerfectFile(App.Path, "pupdate.zip")) = FL_SUCCESS Then
    Call DeleteFile(getPerfectFile(App.Path, "pupdate.ds"))
    Call UnZipfile("pupdate.zip")
Else
    Debug.Assert False
    MsgBox "FTP 연결 실패!", vbCritical
    Exit Sub '오류
End If


'------------------------------------------------------------------------------
'버젼을 비교한다.
'------------------------------------------------------------------------------
'1. 업데이트 프로그램을 다운로드할 필요가 있는지를 비교한다.(pUpdate.zip: pUpdate.exe)
'2. 공지사항 파일을 다운로드할 필요가 있는지를 비교한다.(pNotice.zip: notice.txt)
'3. 응용프로그램을 다운로드할 필요가 있는지를 비교한다.(pocketquiz.zip: pocketquiz.exe)

Dim RetStr As String * 255
Dim Ret As Long

Dim LOCALUPDATEFILE As String
Dim REMOTEUPDATEFILE As String

LOCALUPDATEFILE = "pupdate.dl"

REMOTEUPDATEFILE = "pupdate.ds"


Dim retVal_LU As String
Dim retVal_LN As String
Dim retVal_LA As String

Dim retVal_RU As String
Dim retVal_RN As String
Dim retVal_RA As String

If isFile(getPerfectFile(App.Path, "pupdate.dl")) Then
    
    Ret = GetPrivateProfileString("VERSION", "ver_update", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LU = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_notice", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LN = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_app", "", RetStr, 255, getPerfectFile(App.Path, LOCALUPDATEFILE))
    retVal_LA = LeftH(RetStr, Ret)
 
    Ret = GetPrivateProfileString("VERSION", "ver_update", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RU = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_notice", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RN = LeftH(RetStr, Ret)
    
    Ret = GetPrivateProfileString("VERSION", "ver_app", "", RetStr, 255, getPerfectFile(App.Path, REMOTEUPDATEFILE))
    retVal_RA = LeftH(RetStr, Ret)

    If retVal_LA < retVal_RA Then
        
        '------------------------------------------------------------------------------
        'FTP를 통해서 pquiz.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
Label1.Caption = " pquiz.zip 다운로드중 ..."
        If ctlFTP.FTPDownloadFile(addr, usr, pass, "pquiz.zip", getPerfectFile(App.Path, "pquiz.zip")) = FL_SUCCESS Then
'            Call Kill("putool.exe")
            Call DeleteFile(getPerfectFile(App.Path, "pocketquiz.exe"))
            Label1.Caption = "pquiz.zip 파일을 압축해제중..."
            If UnZipfile("pquiz.zip") Then
                WritePrivateProfileString "VERSION", "ver_app", retVal_RA, getPerfectFile(App.Path, LOCALUPDATEFILE)
            End If
            Label1.Caption = "pquiz.zip 파일을 압축해제완료!"
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
    
    End If

Label1.Caption = "준비"

    If retVal_LN < retVal_RN Or (Not isFile(getPerfectFile(App.Path, "pnotice.txt"))) Then
        
        '------------------------------------------------------------------------------
        'FTP를 통해서 pnotice.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
        If ctlFTP.FTPDownloadFile(addr, usr, pass, "pnotice.zip", getPerfectFile(App.Path, "pnotice.zip")) = FL_SUCCESS Then
'            Call Kill("pnotice.txt")
            Call DeleteFile(getPerfectFile(App.Path, "pnotice.txt"))
            Call UnZipfile("pnotice.zip")
            Call Shell("notepad " & getPerfectFile(App.Path, "pnotice.txt"), vbNormalFocus)
            WritePrivateProfileString "VERSION", "ver_notice", retVal_RN, getPerfectFile(App.Path, LOCALUPDATEFILE)
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
        
    End If

'    If retVal_LA < retVal_RA Then

        If Not isFile(getPerfectFile(App.Path, "pocketquiz.exe")) Then
            Call UnZipfile("pquiz.zip")
        End If
        On Error Resume Next
        
        Call Shell(getPerfectFile(App.Path, "pocketquiz.exe /noupdate " & ocmd), vbNormalFocus)
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbExclamation
        End If
        End
'    End If

Else
    
        '------------------------------------------------------------------------------
        'FTP를 통해서 pquiz.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
Label1.Caption = " pquiz.zip 다운로드중...."
        If ctlFTP.FTPDownloadFile(addr, usr, pass, "pquiz.zip", getPerfectFile(App.Path, "pquiz.zip")) = FL_SUCCESS Then
'            Call Kill("putool.exe")
            Call DeleteFile(getPerfectFile(App.Path, "pocketquiz.exe"))
            
            Label1.Caption = "pquiz.zip 파일을 압축해제중..."
            If UnZipfile("pquiz.zip") Then
                WritePrivateProfileString "VERSION", "ver_app", retVal_RA, getPerfectFile(App.Path, LOCALUPDATEFILE)
            End If
            Label1.Caption = "pquiz.zip 파일을 압축완료!"
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
        
        
        '------------------------------------------------------------------------------
        'FTP를 통해서 pnotice.zip 파일을 다운로드한다.
        '------------------------------------------------------------------------------
        If ctlFTP.FTPDownloadFile(addr, usr, pass, "pnotice.zip", getPerfectFile(App.Path, "pnotice.zip")) = FL_SUCCESS Then
'            Call Kill("pnotice.txt")
            Call DeleteFile(getPerfectFile(App.Path, "pnotice.txt"))
            Call UnZipfile("pnotice.zip")
            Call Shell("notepad " & getPerfectFile(App.Path, "pnotice.txt"), vbNormalFocus)
            WritePrivateProfileString "VERSION", "ver_notice", retVal_RN, getPerfectFile(App.Path, LOCALUPDATEFILE)
        Else
            Debug.Assert False
            Exit Sub '오류
        End If
        
        Call Shell(getPerfectFile(App.Path, "pocketquiz.exe /ftp " & addr & " " & usr & " " & pass & " " & ocmd & " /noupdate"), vbNormalFocus)
        End

End If


'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

End Sub

Function UnZipfile(fname As String) As Boolean
'    Dim pzFileCol As PSZipFileCol
'    Dim PZFile As PSZipFile
    
    UnZipfile = Show_ZipContents2(getPerfectFile(App.Path, fname), App.Path)
    
'
'    pZ.InFileName = getPerfectFile(App.Path, fname)
'    pZ.RootBasePath = App.Path
'
'    pZ.GetZipFileInfos
'
'    Set pzFileCol = pZ.Item
'    Dim i As Long
'    For i = 1 To pzFileCol.Count
'        UnZipfile = True
'        Set PZFile = pzFileCol.Item(i)
'
'
'        pZ.InFileName = getPerfectFile(App.Path, fname)
'        pZ.RootBasePath = App.Path
'
'        pZ.ResetFileList
'        pZ.AddFileList PZFile.FileName
'
'
'        Call pZ.Decompress(0&)
'    Next

End Function

'Private Sub pZ_OverWrite(ByVal FileName As String, pVal As Integer)
'
'pVal = 1
'
'End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
Me.Show
CheckFtp cmd
End
End Sub

Private Function Show_ZipContents2(fnamefull As String, ToDir As String) As Boolean
    Dim X As Long
    Dim Enc As String
    Dim DirCnt As Long
    Dim FileCnt As Long
    Dim Temp As Long
    ZF.Get_Contents (fnamefull)
'    For X = 1 To lstInZip.ListItems.Count
'        lstInZip.ListItems(X).Selected = False
'    Next
    For X = 1 To 1
'        With lstInZip
            Enc = " "
            If ZF.Encrypted(X) Then Enc = "+"
            If Not ZF.IsDir(X) Then
                FileCnt = FileCnt + 1
'                .ListItems.Add X, , Enc & ZF.FileName(X)
'                .ListItems(X).SubItems(1) = ZF.Method(X)
                Temp = ZF.CRC32(X)
                If Temp = 0 Then
'                    .ListItems(X).SubItems(2) = "?"
                Else
'                    .ListItems(X).SubItems(2) = Hex(Temp)
                End If
                Temp = ZF.Compressed_Size(X)
'                If Temp = 0 Then
'                    .ListItems(X).SubItems(3) = "?"
'                Else
'                    .ListItems(X).SubItems(3) = Temp
'                End If
'                .ListItems(X).SubItems(4) = ZF.UnCompressed_Size(X)
'                .ListItems(X).SubItems(5) = ZF.FileDateTime(X)
            Else
                DirCnt = DirCnt + 1
'                .ListItems.Add X, , Enc & ZF.FileName(X)
'                .ListItems(X).SubItems(1) = ZF.Method(X)
'                .ListItems(X).SubItems(2) = "Directory Entry"
'                .ListItems(X).SubItems(3) = "Directory Entry"
'                .ListItems(X).SubItems(4) = "Directory Entry"
'                .ListItems(X).SubItems(5) = ZF.FileDateTime(X)
            End If
'        End With
    Next
    If ZF.FileCount > 0 Then
'        lblHeadLine.Caption = "Contents of " & Filetype(PackFileType) & " file " & _
'                              FileList.FileName & " -> " & _
'                              DirCnt & " directories and " & _
'                              FileCnt & " files"
    Else
'        lblHeadLine.Caption = "Not supported format"
    End If
'    If ZF.CanUnpack Then
'        btnUnzip.Enabled = True
'    Else
'        btnUnzip.Enabled = False
'    End If
    
    Call UnZipAll(ZF.FileCount, ToDir)
    Show_ZipContents2 = True
End Function

Private Sub UnZipAll(ByVal Cnt As Long, ToDir As String)
    Dim FileUnzip() As Boolean
    Dim Sel As Boolean
    Dim X As Long
    Dim RetVal As Boolean
        ReDim FileUnzip(Cnt)
        For X = 1 To Cnt
                FileUnzip(X) = True
        Next
    If ToDir = "" Then
        MsgBox "No path to store files"
        Exit Sub
    End If
    MousePointer = vbHourglass
    RetVal = ZF.UnPack(FileUnzip, ToDir)
    MousePointer = vbNormal
End Sub


Private Sub ZF_Percent(ByVal p100 As Integer)
UpdateStatus PB, p100 / 100, True
End Sub
