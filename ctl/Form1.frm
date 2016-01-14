VERSION 5.00
Object = "{E3BF4609-51E4-4D23-B755-609961F9B909}#1.0#0"; "prjNumText.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   270
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   225
      Width           =   960
   End
   Begin prjNumText.numText numText3 
      Height          =   465
      Left            =   1350
      TabIndex        =   2
      Top             =   2205
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "±¼¸²"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
   End
   Begin prjNumText.numText numText2 
      Height          =   690
      Left            =   1485
      TabIndex        =   1
      Top             =   1350
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1217
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "±¼¸²"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
   End
   Begin prjNumText.numText numText1 
      CausesValidation=   0   'False
      Height          =   825
      Left            =   1530
      TabIndex        =   0
      Top             =   225
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1455
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "±¼¸²"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Validate(Cancel As Boolean)
If Text1.Text = "" Then Cancel = True
End Sub
