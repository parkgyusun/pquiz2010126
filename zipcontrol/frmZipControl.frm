VERSION 5.00
Begin VB.Form frmZipControl 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   1185
   End
End
Attribute VB_Name = "frmZipControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents pz As PSZIPLib.PSZLib
Attribute pz.VB_VarHelpID = -1

Private Sub Command1_Click()


Dim pfilecol As PSZIPLib.PSZipFileCol
Dim pfile As PSZIPLib.PSZipFile

pz.InFileName = "c:\a.zip"
pz.RootBasePath = "c:\"

pz.GetZipFileInfos

Set pfilecol = pz.Item
Dim i
For i = 1 To pfilecol.Count
    Set pfile = pfilecol.Item(i)
    pz.ResetFileList
    pz.AddFileList pfile.FileName
    
    
    pz.InFileName = "c:\a.zip"
    pz.RootBasePath = "c:\"
    
    Call pz.Decompress(0)
Next

End Sub

Private Sub Form_Load()
Set pz = New PSZIPLib.PSZLib
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set pz = Nothing
End Sub

Private Sub pz_EndResult(ByVal bResult As Long)
Debug.Print "pz_EndResult", bResult
End Sub

Private Sub pz_OverWrite(ByVal FileName As String, pVal As Integer)
Debug.Print "pz_OverWrite", FileName, pVal
pVal = 1 '덮어쓰기함
End Sub

Private Sub pz_Progress(ByVal nPercent As Long)
Debug.Print nPercent

End Sub

Private Sub pz_ProgressCurState(ByVal State As String)

Debug.Print "pz_ProgressCurState", State

End Sub
