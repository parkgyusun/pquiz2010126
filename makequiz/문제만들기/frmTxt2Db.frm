VERSION 5.00
Begin VB.Form frmTxt2Db 
   Caption         =   "txt->db"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdDBCopy 
      Caption         =   "mdb2mysql"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdToeic990 
      Caption         =   "토익990-20071103"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1005
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   780
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1275
   End
End
Attribute VB_Name = "frmTxt2Db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ObjRegExp As New RegExp

Private Sub cmd_Click()

End Sub

Private Sub cmdDBCopy_Click()
Dim con1 As New ADODB.Connection
Dim rs1 As ADODB.Recordset

Dim con2 As New ADODB.Connection
Dim pstmt As New ADODB.Command
Dim rs2 As ADODB.Recordset

Dim STRCON As String

STRCON = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mysql;DESC=;DATABASE=" + "pocket_znlwm_000001" + ";SERVER=" + "59.150.9.39" + ";UID=" + "changeit" + ";PASSWORD=" + "" + ";PORT=" + "80" + ";SOCKET=;OPTION=" + "2097155" + ";STMT=;"""

con1.Open "dsn=db;uid=admin;pwd=zkxldk"
'con1.Open "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" + App.Path + "\db3.mdb"
con2.Open STRCON

'Call rs1.Open("select * from tq14", con1, adOpenForwardOnly)
Set rs1 = con1.Execute("select * from tq14")
Dim strSql As String
Do Until rs1.EOF
    strSql = "insert into tq14(subj,seq,cat,quiz,a,b,c,d,e,hint,ans,resid,mode,updateymd,updatechasu) values("
    strSql = strSql + "'" + chkString(rs1(0)) + "', "
    strSql = strSql + "" + chkString(rs1(1)) + ", "
    strSql = strSql + "'" + chkString(rs1(2)) + "', "
    strSql = strSql + "'" + chkString(rs1(3)) + "', "
    strSql = strSql + "'" + chkString(rs1(4)) + "', "
    strSql = strSql + "'" + chkString(rs1(5)) + "', "
    strSql = strSql + "'" + chkString(rs1(6)) + "', "
    strSql = strSql + "'" + chkString(rs1(7)) + "', "
    strSql = strSql + "'" + chkString(rs1(8)) + "', "
    strSql = strSql + "'" + chkString(rs1(9)) + "', "
    strSql = strSql + "'" + chkString(rs1(10)) + "', "
    strSql = strSql + "'" + chkString(rs1(11)) + "', "
    strSql = strSql + "'4', "
    strSql = strSql + "'20071212', "
    strSql = strSql + "0)"
    con1.Execute strSql
    rs1.MoveNext
Loop


End Sub
Public Function chkString(val As Variant) As Variant
On Error GoTo ErrTrap1
chkString = ""
If IsNull(val) Then
    chkString = ""
Else
    chkString = Replace(val, "'", "''")
End If
Exit Function
ErrTrap1:
End Function

Private Sub cmdToeic990_Click()
On Error GoTo Err1

Dim txt As String
Dim txt1 As String
Dim qz As Variant
Dim pos1 As Long
Dim pos2, pos3, pos4, pos6, pos7, pos8, pos9
Dim pos5
Dim cat As String
Dim hint As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim ans As String
Dim quiz As String
Dim ssql As String
Dim rcnt As Integer
Dim c100 As Integer

Dim testMode As Boolean
Dim startBit As Boolean

testMode = False 'true 면 DB에 넣지 않음
startBit = False '스캔한 라인을 점검한다.

Dim v_num As Integer
Dim v_bef_num As Integer

Open App.Path & "/../문제뱅크/토익문제만들기990/토익_문제_은행_990제.txt" For Input As #1
    Do Until EOF(1)
        
        Line Input #1, txt1
        If Not startBit Then
            '시작조건
            v_bef_num = 1
            If InStr(txt1, "[[" & v_bef_num & ".") > 0 Then
                
                startBit = True
                txt = txt & vbCrLf & txt1
            End If
        Else
            
            '종료조건
            If InStr(txt1, "[[954.") > 0 Then
                Exit Do
            End If
            
            txt = txt & vbCrLf & txt1
        
        End If
    Loop
Close #1
txt = Replace(txt, vbLf & "[[", Chr(2))
qz = Split(txt, Chr(2))
Dim con As New ADODB.Connection
'Dim rs As New ADODB.Recordset

Dim STRCON As String

If Not testMode Then

STRCON = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=mysql;DESC=;DATABASE=" + "pocket_znlwm_000001" + ";SERVER=" + "59.150.9.39" + ";UID=" + "changeit" + ";PASSWORD=" + "" + ";PORT=" + "80" + ";SOCKET=;OPTION=" + "2097155" + ";STMT=;"""

con.Open STRCON ' "dsn=db;uid=admin;pwd=zkxldk"
End If

Dim i As Integer


For i = 1 To UBound(qz)
    pos1 = 0
    pos2 = 0
    pos3 = 0
    pos4 = 0
    pos5 = 0
    pos6 = 0
    
    txt1 = qz(i)
    
    pos1 = InStr(txt1, ".")
    If pos1 > 6 Then
        MsgBox "문제번호 다음에 식별자.기호 없음" & vbCrLf & txt1
        Debug.Assert False
        
    End If
    
    
    pos2 = InStr(pos1, txt1, "" & "(A)")
    If pos2 = 0 Then
        pos2 = InStr(pos1, txt1, " " & "a)")
        If pos2 = 0 Then
            pos2 = InStr(pos1, txt1, "(" & "a)")
            pos3 = InStr(pos2, txt1, "(" & "b)")
            pos4 = InStr(pos3, txt1, "(" & "c)")
            pos5 = InStr(pos4, txt1, "(" & "d)")
            pos6 = InStr(pos5, txt1, " " & "[정답]")

        Else
            pos3 = InStr(pos2, txt1, " " & "b)")
            pos4 = InStr(pos3, txt1, " " & "c)")
            pos5 = InStr(pos4, txt1, " " & "d)")
            pos6 = InStr(pos5, txt1, " " & "[정답]")
        End If
        
        pos8 = InStr(pos5, txt1, vbLf & "어휘:")
        If pos8 = 0 Then
            pos8 = InStr(pos5, txt1, vbLf & "해설:")
        End If
        If pos8 = 0 Then
            pos8 = InStr(pos5, txt1, vbLf & "번역:")
        End If
        
        If pos8 = 0 Then
            pos8 = InStr(pos5, txt1, vbCrLf)
        End If
    
    Else
        pos3 = InStr(pos2, txt1, "" & "(B)", vbTextCompare)
        If pos3 = 0 Then
            pos2 = InStr(pos1, txt1, "" & "(A)", vbTextCompare)
            pos3 = InStr(pos2, txt1, "" & "(B)", vbTextCompare)
        End If
        pos4 = InStr(pos3, txt1, "" & "(C)", vbTextCompare)
        If pos4 = 0 Then
            pos2 = InStr(pos1, txt1, "" & "(A)", vbTextCompare)
            pos3 = InStr(pos2, txt1, "" & "(B)", vbTextCompare)
            pos4 = InStr(pos3, txt1, "" & "(C)", vbTextCompare)
        End If
        pos5 = InStr(pos4, txt1, "" & "(D)", vbTextCompare)
        pos6 = InStr(pos5, txt1, vbLf & "정답:")
        If pos6 = 0 Then
            pos6 = InStr(pos5, txt1, vbLf & "[정답")
        End If
        
        pos8 = InStr(pos5, txt1, vbLf & "어휘:")
        If pos8 = 0 Then
            pos8 = InStr(pos5, txt1, vbLf & "해설:")
        End If
        If pos8 = 0 Then
            pos8 = InStr(pos5, txt1, vbLf & "번역:")
        End If

    
    End If
    pos9 = InStr(txt1, vbCr)
     
        
    Debug.Assert Not (pos5 = 0 Or pos4 = 0 Or pos3 = 0 Or pos2 = 0 Or pos6 = 0 Or pos5 = 0)
    
    
    cat = Mid(txt1, 1, pos1 - 1)
    
    If pos8 = 0 Then
        MsgBox "힌트구분 없음 어휘" & vbCrLf & cat
    End If
    
    If pos6 > 0 Then
        pos7 = InStr(pos6 + 1, txt1, vbLf & "정답:")
    Else
        MsgBox "정답이 없음" & vbCrLf & cat
        GoTo skip1
    End If
    
    
    If pos7 > 0 Then
        If InStr(txt1, "[[") > 0 Then
            MsgBox cat & vbCrLf & "문제가 2개"
            Debug.Assert False
        Else
            MsgBox cat & vbCrLf & "정답이 2개"
            Debug.Assert False
        End If
    End If
    
    v_num = CInt(cat)
    
    If v_bef_num + 1 <> v_num Then
        MsgBox "문제증가 번호 불일치" & cat
        Debug.Assert False
    End If
    
    v_bef_num = v_num
    
    If pos9 < pos2 Then
        quiz = regTrim(Mid(txt1, pos1 + 2, pos2 - pos1 - 3))
    Else
        If False Then
            quiz = regTrim(Mid(txt1, pos2, pos8 - pos2))
        Else
            quiz = regTrim(Mid(txt1, pos1 + 1, pos8 - (pos1 + 1)))
        End If
    End If
    
    a = regTrim(Mid(txt1, pos2 + 4, pos3 - pos2 - 4))
    b = regTrim(Mid(txt1, pos3 + 4, pos4 - pos3 - 4))
    c = regTrim(Mid(txt1, pos4 + 4, pos5 - pos4 - 4))
    d = regTrim(Mid(txt1, pos5 + 4, pos8 - pos5 - 4))
    
    ans = UCase(Trim(Mid(txt1, pos6 + 7, 1)))
    
    hint = "@ " & Trim(Mid(txt1, pos8))
    hint = Replace(hint, "정답:  (A)", "")
    hint = Replace(hint, "정답:  (B)", "")
    hint = Replace(hint, "정답:  (C)", "")
    hint = Replace(hint, "정답:  (D)", "")
    
    hint = Replace(hint, "[정답] (a)", "")
    hint = Replace(hint, "[정답] (b)", "")
    hint = Replace(hint, "[정답] (c)", "")
    hint = Replace(hint, "[정답] (d)", "")
    
'    Debug.Print "[cat]" & cat
'    Debug.Print "[quiz]" & quiz
'    Debug.Print "[a]" & a
'    Debug.Print "[b]" & b
'    Debug.Print "[c]" & c
'    Debug.Print "[d]" & d
'    Debug.Print "[ans]" & ans
'    Debug.Print "[hint]" & hint
    
    Debug.Assert (ans = "A" Or ans = "B" Or ans = "C" Or ans = "D")
'    Debug.Assert i <> 76
    If Not testMode Then
        ssql = "insert into tq14(subj,seq,cat,quiz,a,b,c,d,e,hint,ans,resid) values('토익990'," & i & ",'" & cat & "','" & Replace(quiz, "'", "''") & "','" & Replace(a, "'", "''") & "','" & Replace(b, "'", "''") & "','" & Replace(c, "'", "''") & "','" & Replace(d, "'", "''") & "','','" & Replace(hint, "'", "''") & "','" & ans & "',0)"
        
        Call con.Execute(ssql, rcnt)
    End If

'    Debug.Assert rcnt = 1
'    Debug.Print rcnt
skip1:
Next

If Not testMode Then
    con.Close
End If

MsgBox "work end"
Exit Sub
Err1:
MsgBox Err.Description + "멈춤 클릭후 디버그 할 것", vbCritical
Debug.Assert False
Resume

End Sub

Private Sub Command1_Click()
Dim txt As String
Dim qz As Variant
Open App.Path & "/../모의고사.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, txt1
        txt = txt & vbCrLf & txt1
    Loop
Close #1
txt = Replace(txt, vbLf & "[", Chr(2))
qz = Split(txt, Chr(2))
Dim con As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'카티아
con.Open "dsn=db;uid=admin;pwd=zkxldk"
Dim i As Integer
On Error Resume Next
    
For i = 1 To UBound(qz)
    pos1 = 0
    pos2 = 0
    pos3 = 0
    pos4 = 0
    pos5 = 0
    pos6 = 0
    
    txt1 = qz(i)
    
    pos1 = InStr(txt1, "]")
    pos2 = InStr(pos1, txt1, vbLf & "A. ")
    pos3 = InStr(pos2, txt1, vbLf & "B. ")
    pos4 = InStr(pos3, txt1, vbLf & "C. ")
    pos5 = InStr(pos4, txt1, vbLf & "D. ")
    pos6 = InStr(pos5, txt1, vbLf & "@ ")
    pos7 = InStrRev(txt1, "답")
    
    
    Debug.Assert Not (pos6 = 0 Or pos5 = 0)
    
    cat = Mid(txt1, 1, pos1 - 1)
    quiz = Mid(txt1, pos1 + 2, pos2 - pos1 - 3)
    a = Mid(txt1, pos2 + 4, pos3 - pos2 - 5)
    b = Mid(txt1, pos3 + 4, pos4 - pos3 - 5)
    c = Mid(txt1, pos4 + 4, pos5 - pos4 - 5)
    d = Mid(txt1, pos5 + 4, pos6 - pos5 - 5)
    hint = Mid(txt1, pos6 + 1, pos7 - pos6 - 1)
    ans = Mid(txt1, pos7 + 2, 1)
    
    Debug.Assert cat <> "D2-2"
    Debug.Assert (ans = "A" Or ans = "B" Or ans = "C" Or ans = "D")
    
    ssql = "insert into tq01 values('모의AA'," & i & ",'" & cat & "','" & Replace(quiz, "'", "`") & "','" & a & "','" & b & "','" & c & "','" & d & "','" & hint & "','" & ans & "',0)"
   con.Execute ssql
Next
con.Close

End Sub

Private Sub Command2_Click()
Open App.Path & "/../2003.txt" For Input As #1
Open App.Path & "/../2003_2.txt" For Output As #2
    Do Until EOF(1)
        Line Input #1, txt1
        If Not ((Len(txt1) <= 4 And IsNumeric(txt1)) Or Left(txt1, 1) = "ⓒ") Then
        If Mid(txt1, 3, 1) = "-" And IsNumeric(Mid(txt1, 2, 1)) Then
            txt1 = "[" & Replace(txt1, " ", "] ", 1, 1)
        ElseIf Left(txt1, 1) = "?" Then
            txt1 = Replace(txt1, "?", "답", 1, 1)
        End If
        Print #2, txt1
        '    txt = txt & vbCrLf & txt1
        End If
    Loop
Close #1
Close #2
End Sub

Private Sub Command3_Click()
Dim txt As String
Dim txt1 As String
Dim qz As Variant
Dim pos1 As Long
Dim pos2, pos3, pos4, pos6, pos7, pos8
Dim pos5
Dim cat As String
Dim hint As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim ans As String
Dim quiz As String
Dim ssql As String
Dim rcnt As Integer


Open App.Path & "/../모의고사.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, txt1
        txt = txt & vbCrLf & txt1
    Loop
Close #1
txt = Replace(txt, vbLf & "[[", Chr(2))
qz = Split(txt, Chr(2))
Dim con As New ADODB.Connection
'Dim rs As New ADODB.Recordset

con.Open "dsn=db;uid=admin;pwd=zkxldk"
Dim i As Integer
On Error Resume Next
    
For i = 1 To UBound(qz)
    pos1 = 0
    pos2 = 0
    pos3 = 0
    pos4 = 0
    pos5 = 0
    pos6 = 0
    
    txt1 = qz(i)
    
    pos1 = InStr(txt1, "]")
    pos2 = InStr(pos1, txt1, " " & "(A)")
    pos3 = InStr(pos2, txt1, " " & "(B)")
    pos4 = InStr(pos3, txt1, " " & "(C)")
    pos5 = InStr(pos4, txt1, " " & "(D)")
    pos6 = InStr(pos5, txt1, vbLf & "정답:")
    
    If pos6 > 0 Then
        pos7 = InStrRev(txt1, "답")
    End If
    
    Debug.Assert Not (pos5 = 0 Or pos4 = 0 Or pos3 = 0 Or pos2 = 0 Or pos6 = 0 Or pos5 = 0)
    
    cat = Mid(txt1, 1, pos1 - 1)
    quiz = regTrim(Mid(txt1, pos1 + 2, pos2 - pos1 - 3))
    a = regTrim(Mid(txt1, pos2 + 3, pos3 - pos2 - 4))
    b = regTrim(Mid(txt1, pos3 + 3, pos4 - pos3 - 4))
    c = regTrim(Mid(txt1, pos4 + 3, pos5 - pos4 - 4))
    d = regTrim(Mid(txt1, pos5 + 3, pos6 - pos5 - 4))
    ans = regTrim(Mid(txt1, pos6 + 3, 1))
    hint = "@ " & regTrim(Mid(txt1, pos6 + 5))
    
    Debug.Assert (ans = "A" Or ans = "B" Or ans = "C" Or ans = "D")
    Debug.Assert i <> 76
    ssql = "insert into tq01 values('모의AA'," & i & ",'" & cat & "','" & quiz & "','" & a & "','" & b & "','" & c & "','" & d & "','" & Replace(hint, "'", "''") & "','" & ans & "',0)"
'    Call con.Execute(ssql, rcnt)
    Debug.Assert rcnt = 1
    Debug.Print rcnt
Next
con.Close
End Sub


Private Function regTrim(str) As String
    regTrim = ObjRegExp.Replace(str, "")
End Function



Private Sub Form_Load()

    ObjRegExp.Global = True
    ObjRegExp.IgnoreCase = True
    ObjRegExp.MultiLine = False
    ObjRegExp.Pattern = "^\s*|\s*$"

End Sub
