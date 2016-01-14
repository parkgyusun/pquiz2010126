VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

Public quiz As String
Public cat As String
Public a As String
Public b As String
Public c As String
Public d As String
Public e As String

Public a2 As String '줄바꿈된 보기
Public b2 As String
Public C2 As String
Public d2 As String
Public e2 As String
Public hint2 As String

Public o As Long
Public x As Long

Public update_ymd As String
Public reserve_ymd As String
Public gangyek As Long

Public hint As String
Public ans As String
Public mode As String '1,2,4,5
Public resType As String '@,ⓘ,ⓗ,ⓜ,ⓢ

Public memo As String
Public resid As String
Public isNew As Boolean
Public Correct As Boolean
Public Correct_chk As Boolean
Public forReview As Boolean
Public pocketnm As String
Public num As Long
Public seq As Long
Public subj As String
Public chk As Integer
Public overFlow As Boolean
Public isBefore As Boolean
Public dLastSec As Double
Public bPass As Boolean
Public bNext As Boolean

Public A_seq As String
Public B_seq As String
Public C_seq As String
Public D_seq As String
Public E_seq As String

Public A_ans As String
Public B_ans As String
Public C_ans As String
Public D_ans As String
Public E_ans As String

Public user_ans As Long

Public Sub swap()
Randomize
Dim K As Integer
Dim tr As String

Debug.Assert mode <> "2"

If mode = "4" Or mode = "1" Then
    K = Fix(Rnd * 1000) Mod 4
ElseIf mode = "5" Then
    K = Fix(Rnd * 1000) Mod 5
End If

If mode = "4" Or mode = "1" Then

    If K = 0 Then
        tr = "A"
    ElseIf K = 1 Then
        tr = "B"
    ElseIf K = 2 Then
        tr = "C"
    ElseIf K = 3 Then
        tr = "D"
    End If
ElseIf mode = "5" Then
    If K = 0 Then
        tr = "A"
    ElseIf K = 1 Then
        tr = "B"
    ElseIf K = 2 Then
        tr = "C"
    ElseIf K = 3 Then
        tr = "D"
    ElseIf K = 4 Then
        tr = "E"
    End If
End If

If mode = "1" Or mode = "4" Then
    If swapOK() Then
        If ans = "A" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(c, b)
                Call swap2_seq(C_seq, B_seq)
                Call swap2_ans(C_ans, B_ans)
                
            Case "B"
                ans = "B"
                Call swap2(a, b)
                Call swap2_seq(A_seq, B_seq)
                Call swap2_ans(A_ans, B_ans)
                
            Case "C"
                ans = "C"
                Call swap2(a, c)
                Call swap2_seq(A_seq, C_seq)
                Call swap2_ans(A_ans, C_ans)
            Case "D"
                ans = "D"
                Call swap2(a, d)
                Call swap2_seq(A_seq, D_seq)
                Call swap2_ans(A_ans, D_ans)
            End Select
        End If
        
        If ans = "B" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(b, a)
                Call swap2_seq(B_seq, A_seq)
                Call swap2_ans(B_ans, A_ans)
            Case "B"
                ans = "B"
                Call swap2(c, d)
                Call swap2_seq(C_seq, D_seq)
                Call swap2_ans(C_ans, D_ans)
            Case "C"
                ans = "C"
                Call swap2(b, c)
                Call swap2_seq(B_seq, C_seq)
                Call swap2_ans(B_ans, C_ans)
            Case "D"
                ans = "D"
                Call swap2(b, d)
                Call swap2_seq(B_seq, D_seq)
                Call swap2_ans(B_ans, D_ans)
            End Select
        End If
        
        If ans = "C" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(c, a)
                Call swap2_seq(C_seq, A_seq)
                Call swap2_ans(C_ans, A_ans)
            Case "B"
                ans = "B"
                Call swap2(c, b)
                Call swap2_seq(C_seq, B_seq)
                Call swap2_ans(C_ans, B_ans)
            Case "C"
                ans = "C"
                Call swap2(d, a)
                Call swap2_seq(D_seq, A_seq)
                Call swap2_ans(D_ans, A_ans)
            Case "D"
                ans = "D"
                Call swap2(c, d)
                Call swap2_seq(C_seq, D_seq)
                Call swap2_ans(C_ans, D_ans)
            End Select
        End If
        
        If ans = "D" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(d, a)
                Call swap2_seq(D_seq, A_seq)
                Call swap2_ans(D_ans, A_ans)
            Case "B"
                ans = "B"
                Call swap2(d, b)
                Call swap2_seq(D_seq, B_seq)
                Call swap2_ans(D_ans, B_ans)
            Case "C"
                ans = "C"
                Call swap2(d, c)
                Call swap2_seq(D_seq, C_seq)
                Call swap2_ans(D_ans, C_ans)
            Case "D"
                ans = "D"
                Call swap2(a, b)
                Call swap2_seq(A_seq, B_seq)
                Call swap2_ans(A_ans, B_ans)
            End Select
        End If
        
    End If
ElseIf mode = "5" Then
    If swapOK() Then
        If ans = "A" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(c, b)
                Call swap2_seq(C_seq, B_seq)
                Call swap2_ans(C_ans, B_ans)
            Case "B"
                ans = "B"
                Call swap2(a, b)
                Call swap2_seq(A_seq, B_seq)
                Call swap2_ans(A_ans, B_ans)
            Case "C"
                ans = "C"
                Call swap2(a, c)
                Call swap2_seq(A_seq, C_seq)
                Call swap2_ans(A_ans, C_ans)
            Case "D"
                ans = "D"
                Call swap2(a, d)
                Call swap2_seq(A_seq, D_seq)
                Call swap2_ans(A_ans, D_ans)
            Case "E"
                ans = "E"
                Call swap2(a, e)
                Call swap2_seq(A_seq, E_seq)
                Call swap2_ans(A_ans, E_ans)
            End Select
        End If
        
        If ans = "B" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(b, a)
                Call swap2_seq(B_seq, A_seq)
                Call swap2_ans(B_ans, A_ans)
            Case "B"
                ans = "B"
                Call swap2(c, d)
                Call swap2_seq(C_seq, D_seq)
                Call swap2_ans(C_ans, D_ans)
            Case "C"
                ans = "C"
                Call swap2(b, c)
                Call swap2_seq(B_seq, C_seq)
                Call swap2_ans(B_ans, C_ans)
            Case "D"
                ans = "D"
                Call swap2(b, d)
                Call swap2_seq(B_seq, D_seq)
                Call swap2_ans(B_ans, D_ans)
            Case "E"
                ans = "E"
                Call swap2(b, e)
                Call swap2_seq(B_seq, E_seq)
                Call swap2_ans(B_ans, E_ans)
            End Select
        End If
        
        If ans = "C" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(c, a)
                Call swap2_seq(C_seq, A_seq)
                Call swap2_ans(C_ans, A_ans)
            Case "B"
                ans = "B"
                Call swap2(c, b)
                Call swap2_seq(C_seq, B_seq)
                Call swap2_ans(C_ans, B_ans)
            Case "C"
                ans = "C"
                Call swap2(d, a)
                Call swap2_seq(D_seq, A_seq)
                Call swap2_ans(D_ans, A_ans)
            Case "D"
                ans = "D"
                Call swap2(c, d)
                Call swap2_seq(C_seq, D_seq)
                Call swap2_ans(C_ans, D_ans)
            Case "E"
                ans = "E"
                Call swap2(c, e)
                Call swap2_seq(C_seq, E_seq)
                Call swap2_ans(C_ans, E_ans)
            End Select
        End If
        
        If ans = "D" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(d, a)
                Call swap2_seq(D_seq, A_seq)
                Call swap2_ans(D_ans, A_ans)
            Case "B"
                ans = "B"
                Call swap2(d, b)
                Call swap2_seq(D_seq, B_seq)
                Call swap2_ans(D_ans, B_ans)
            Case "C"
                ans = "C"
                Call swap2(d, c)
                Call swap2_seq(D_seq, C_seq)
                Call swap2_ans(D_ans, C_ans)
            Case "D"
                ans = "D"
                Call swap2(a, b)
                Call swap2_seq(A_seq, B_seq)
                Call swap2_ans(A_ans, B_ans)
            Case "E"
                ans = "E"
                Call swap2(d, e)
                Call swap2_seq(D_seq, E_seq)
                Call swap2_ans(D_ans, E_ans)
            End Select
        End If
        
        If ans = "E" Then
            Select Case tr
            Case "A"
                ans = "A"
                Call swap2(e, a)
                Call swap2_seq(E_seq, A_seq)
                Call swap2_ans(E_ans, A_ans)
            Case "B"
                ans = "B"
                Call swap2(e, b)
                Call swap2_seq(E_seq, B_seq)
                Call swap2_ans(E_ans, B_ans)
            Case "C"
                ans = "C"
                Call swap2(e, c)
                Call swap2_seq(E_seq, C_seq)
                Call swap2_ans(E_ans, C_ans)
            Case "D"
                ans = "D"
                Call swap2(e, d)
                Call swap2_seq(E_seq, D_seq)
                Call swap2_ans(E_ans, D_ans)
            Case "E"
                ans = "E"
                Call swap2(a, b)
                Call swap2_seq(A_seq, B_seq)
                Call swap2_ans(A_ans, B_ans)
            End Select
        End If
        
    End If
Else
    Debug.Assert False
End If

End Sub

Private Function swapOK() As Boolean

swapOK = NoVisibleHintCmdCase()

If swapOK = False Then
    'swapOK = True
    swapOK = NoExistABCDinHint()
End If
End Function
Public Function NoVisibleHintCmdCase() As Boolean
NoVisibleHintCmdCase = Not (isNew Or forReview)
End Function
Public Function NoExistABCDinHint() As Boolean
    If InStr(hint, "A") > 0 Then
        NoExistABCDinHint = False
        Exit Function
    ElseIf InStr(hint, "B") > 0 Then
        NoExistABCDinHint = False
        Exit Function
    
    ElseIf InStr(hint, "C") > 0 Then
        NoExistABCDinHint = False
        Exit Function
    
    ElseIf InStr(hint, "D") > 0 Then
        NoExistABCDinHint = False
        Exit Function
    End If
NoExistABCDinHint = True
End Function
Private Sub swap2(ByRef s1 As String, ByRef s2 As String)
Dim tmp As String
tmp = s1
s1 = s2
s2 = tmp

End Sub

Private Sub swap2_seq(ByRef s1 As String, ByRef s2 As String)
Dim tmp As String
tmp = s1
s1 = s2
s2 = tmp

End Sub

Private Sub swap2_ans(ByRef s1 As String, ByRef s2 As String)
Dim tmp As String
tmp = s1
s1 = s2
s2 = tmp

End Sub


Public Sub cls()
ans = ""
forReview = False
End Sub

