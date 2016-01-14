Attribute VB_Name = "FileExists"
Public Function FileExists(ByVal Fname As String) As Boolean
    Dim TheFile As String
    Dim Results As String
    
    TheFile = Fname
    Results = Dir$(TheFile)
    
    If Len(Results) = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function


