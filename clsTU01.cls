VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTU01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub logout(userid As String)

sSql = "update tu01 set islogon=0 where userid='" & userid & "' and con_time='" & gSessionId & "' "
If Con.State = 0 Then
    Con_Open
    conflag = True
End If

affected = Fn_SQLExec(sSql).nrow

'-2147467259 '단독사용시 오류
'-2147217904 '컬럼 오류시

If conflag Then
Con_Close
End If

End Sub



Public Function BreakedLogOn(sid As String) As Boolean

sSql = "select con_time from tu01 where userid='" & gUserid & "' and con_time='" & sid & "'"
If Con.State = 0 Then
    Con_Open
    conflag = True
End If

Dim cnt As Long

cnt = Fn_SQLExec(sSql).nrow

If cnt = 0 Then
    BreakedLogOn = True
Else
    BreakedLogOn = False
End If

If conflag Then
Con_Close
End If

End Function



