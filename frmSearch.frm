VERSION 5.00
Object = "{D8D562C3-878C-11D2-943F-444553540000}#1.0#0"; "ctlist.ocx"
Begin VB.Form frmSearch 
   Caption         =   "검색"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewPop 
      Caption         =   "팝업보기"
      Height          =   465
      Left            =   7800
      TabIndex        =   10
      Top             =   6810
      Width           =   1605
   End
   Begin VB.ListBox lstColumn 
      Height          =   1620
      ItemData        =   "frmSearch.frx":0442
      Left            =   6030
      List            =   "frmSearch.frx":0461
      TabIndex        =   9
      Top             =   210
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "검색"
      Default         =   -1  'True
      Height          =   465
      Left            =   7860
      TabIndex        =   8
      Top             =   1110
      Width           =   1605
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   7860
      TabIndex        =   6
      Top             =   570
      Width           =   1575
   End
   Begin VB.CommandButton cmd2L2R 
      Caption         =   ">"
      Height          =   330
      Left            =   2490
      TabIndex        =   3
      Top             =   300
      Width           =   915
   End
   Begin VB.CommandButton cmd2L2Ra 
      Caption         =   ">>"
      Height          =   330
      Left            =   2490
      TabIndex        =   0
      Top             =   660
      Width           =   915
   End
   Begin VB.CommandButton cmd2R2L 
      Caption         =   "<"
      Height          =   330
      Left            =   2490
      TabIndex        =   2
      Top             =   1020
      Width           =   915
   End
   Begin VB.CommandButton cmd2R2La 
      Caption         =   "<<"
      Height          =   330
      Left            =   2490
      TabIndex        =   1
      Top             =   1380
      Width           =   915
   End
   Begin CTLISTLibCtl.ctList lst1_3 
      DragIcon        =   "frmSearch.frx":04A8
      Height          =   1725
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   3043
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleBackImage  =   "frmSearch.frx":0D72
      HeaderPicture   =   "frmSearch.frx":0D8E
      Picture         =   "frmSearch.frx":0DAA
      CheckPicDown    =   "frmSearch.frx":0DC6
      CheckPicUp      =   "frmSearch.frx":0DE2
      CheckPicDisabled=   "frmSearch.frx":0DFE
      BackImage       =   "frmSearch.frx":0E1A
      ShowHeader      =   -1  'True
      MultiSelect     =   -1  'True
      HeaderData      =   "frmSearch.frx":0E36
      PicArray0       =   "frmSearch.frx":0EB3
      PicArray1       =   "frmSearch.frx":0ECF
      PicArray2       =   "frmSearch.frx":0EEB
      PicArray3       =   "frmSearch.frx":0F07
      PicArray4       =   "frmSearch.frx":0F23
      PicArray5       =   "frmSearch.frx":0F3F
      PicArray6       =   "frmSearch.frx":0F5B
      PicArray7       =   "frmSearch.frx":0F77
      PicArray8       =   "frmSearch.frx":0F93
      PicArray9       =   "frmSearch.frx":0FAF
      PicArray10      =   "frmSearch.frx":0FCB
      PicArray11      =   "frmSearch.frx":0FE7
      PicArray12      =   "frmSearch.frx":1003
      PicArray13      =   "frmSearch.frx":101F
      PicArray14      =   "frmSearch.frx":103B
      PicArray15      =   "frmSearch.frx":1057
      PicArray16      =   "frmSearch.frx":1073
      PicArray17      =   "frmSearch.frx":108F
      PicArray18      =   "frmSearch.frx":10AB
      PicArray19      =   "frmSearch.frx":10C7
      PicArray20      =   "frmSearch.frx":10E3
      PicArray21      =   "frmSearch.frx":10FF
      PicArray22      =   "frmSearch.frx":111B
      PicArray23      =   "frmSearch.frx":1137
      PicArray24      =   "frmSearch.frx":1153
      PicArray25      =   "frmSearch.frx":116F
      PicArray26      =   "frmSearch.frx":118B
      PicArray27      =   "frmSearch.frx":11A7
      PicArray28      =   "frmSearch.frx":11C3
      PicArray29      =   "frmSearch.frx":11DF
      PicArray30      =   "frmSearch.frx":11FB
      PicArray31      =   "frmSearch.frx":1217
      PicArray32      =   "frmSearch.frx":1233
      PicArray33      =   "frmSearch.frx":124F
      PicArray34      =   "frmSearch.frx":126B
      PicArray35      =   "frmSearch.frx":1287
      PicArray36      =   "frmSearch.frx":12A3
      PicArray37      =   "frmSearch.frx":12BF
      PicArray38      =   "frmSearch.frx":12DB
      PicArray39      =   "frmSearch.frx":12F7
      PicArray40      =   "frmSearch.frx":1313
      PicArray41      =   "frmSearch.frx":132F
      PicArray42      =   "frmSearch.frx":134B
      PicArray43      =   "frmSearch.frx":1367
      PicArray44      =   "frmSearch.frx":1383
      PicArray45      =   "frmSearch.frx":139F
      PicArray46      =   "frmSearch.frx":13BB
      PicArray47      =   "frmSearch.frx":13D7
      PicArray48      =   "frmSearch.frx":13F3
      PicArray49      =   "frmSearch.frx":140F
      PicArray50      =   "frmSearch.frx":142B
      PicArray51      =   "frmSearch.frx":1447
      PicArray52      =   "frmSearch.frx":1463
      PicArray53      =   "frmSearch.frx":147F
      PicArray54      =   "frmSearch.frx":149B
      PicArray55      =   "frmSearch.frx":14B7
      PicArray56      =   "frmSearch.frx":14D3
      PicArray57      =   "frmSearch.frx":14EF
      PicArray58      =   "frmSearch.frx":150B
      PicArray59      =   "frmSearch.frx":1527
      PicArray60      =   "frmSearch.frx":1543
      PicArray61      =   "frmSearch.frx":155F
      PicArray62      =   "frmSearch.frx":157B
      PicArray63      =   "frmSearch.frx":1597
      PicArray64      =   "frmSearch.frx":15B3
      PicArray65      =   "frmSearch.frx":15CF
      PicArray66      =   "frmSearch.frx":15EB
      PicArray67      =   "frmSearch.frx":1607
      PicArray68      =   "frmSearch.frx":1623
      PicArray69      =   "frmSearch.frx":163F
      PicArray70      =   "frmSearch.frx":165B
      PicArray71      =   "frmSearch.frx":1677
      PicArray72      =   "frmSearch.frx":1693
      PicArray73      =   "frmSearch.frx":16AF
      PicArray74      =   "frmSearch.frx":16CB
      PicArray75      =   "frmSearch.frx":16E7
      PicArray76      =   "frmSearch.frx":1703
      PicArray77      =   "frmSearch.frx":171F
      PicArray78      =   "frmSearch.frx":173B
      PicArray79      =   "frmSearch.frx":1757
      PicArray80      =   "frmSearch.frx":1773
      PicArray81      =   "frmSearch.frx":178F
      PicArray82      =   "frmSearch.frx":17AB
      PicArray83      =   "frmSearch.frx":17C7
      PicArray84      =   "frmSearch.frx":17E3
      PicArray85      =   "frmSearch.frx":17FF
      PicArray86      =   "frmSearch.frx":181B
      PicArray87      =   "frmSearch.frx":1837
      PicArray88      =   "frmSearch.frx":1853
      PicArray89      =   "frmSearch.frx":186F
      PicArray90      =   "frmSearch.frx":188B
      PicArray91      =   "frmSearch.frx":18A7
      PicArray92      =   "frmSearch.frx":18C3
      PicArray93      =   "frmSearch.frx":18DF
      PicArray94      =   "frmSearch.frx":18FB
      PicArray95      =   "frmSearch.frx":1917
      PicArray96      =   "frmSearch.frx":1933
      PicArray97      =   "frmSearch.frx":194F
      PicArray98      =   "frmSearch.frx":196B
      PicArray99      =   "frmSearch.frx":1987
   End
   Begin CTLISTLibCtl.ctList lst1_4 
      DragIcon        =   "frmSearch.frx":19A3
      Height          =   1725
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   3043
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleBackImage  =   "frmSearch.frx":226D
      HeaderPicture   =   "frmSearch.frx":2289
      Picture         =   "frmSearch.frx":22A5
      CheckPicDown    =   "frmSearch.frx":22C1
      CheckPicUp      =   "frmSearch.frx":22DD
      CheckPicDisabled=   "frmSearch.frx":22F9
      BackImage       =   "frmSearch.frx":2315
      ShowHeader      =   -1  'True
      MultiSelect     =   -1  'True
      SortArrows      =   0   'False
      HeaderData      =   "frmSearch.frx":2331
      PicArray0       =   "frmSearch.frx":23AE
      PicArray1       =   "frmSearch.frx":23CA
      PicArray2       =   "frmSearch.frx":23E6
      PicArray3       =   "frmSearch.frx":2402
      PicArray4       =   "frmSearch.frx":241E
      PicArray5       =   "frmSearch.frx":243A
      PicArray6       =   "frmSearch.frx":2456
      PicArray7       =   "frmSearch.frx":2472
      PicArray8       =   "frmSearch.frx":248E
      PicArray9       =   "frmSearch.frx":24AA
      PicArray10      =   "frmSearch.frx":24C6
      PicArray11      =   "frmSearch.frx":24E2
      PicArray12      =   "frmSearch.frx":24FE
      PicArray13      =   "frmSearch.frx":251A
      PicArray14      =   "frmSearch.frx":2536
      PicArray15      =   "frmSearch.frx":2552
      PicArray16      =   "frmSearch.frx":256E
      PicArray17      =   "frmSearch.frx":258A
      PicArray18      =   "frmSearch.frx":25A6
      PicArray19      =   "frmSearch.frx":25C2
      PicArray20      =   "frmSearch.frx":25DE
      PicArray21      =   "frmSearch.frx":25FA
      PicArray22      =   "frmSearch.frx":2616
      PicArray23      =   "frmSearch.frx":2632
      PicArray24      =   "frmSearch.frx":264E
      PicArray25      =   "frmSearch.frx":266A
      PicArray26      =   "frmSearch.frx":2686
      PicArray27      =   "frmSearch.frx":26A2
      PicArray28      =   "frmSearch.frx":26BE
      PicArray29      =   "frmSearch.frx":26DA
      PicArray30      =   "frmSearch.frx":26F6
      PicArray31      =   "frmSearch.frx":2712
      PicArray32      =   "frmSearch.frx":272E
      PicArray33      =   "frmSearch.frx":274A
      PicArray34      =   "frmSearch.frx":2766
      PicArray35      =   "frmSearch.frx":2782
      PicArray36      =   "frmSearch.frx":279E
      PicArray37      =   "frmSearch.frx":27BA
      PicArray38      =   "frmSearch.frx":27D6
      PicArray39      =   "frmSearch.frx":27F2
      PicArray40      =   "frmSearch.frx":280E
      PicArray41      =   "frmSearch.frx":282A
      PicArray42      =   "frmSearch.frx":2846
      PicArray43      =   "frmSearch.frx":2862
      PicArray44      =   "frmSearch.frx":287E
      PicArray45      =   "frmSearch.frx":289A
      PicArray46      =   "frmSearch.frx":28B6
      PicArray47      =   "frmSearch.frx":28D2
      PicArray48      =   "frmSearch.frx":28EE
      PicArray49      =   "frmSearch.frx":290A
      PicArray50      =   "frmSearch.frx":2926
      PicArray51      =   "frmSearch.frx":2942
      PicArray52      =   "frmSearch.frx":295E
      PicArray53      =   "frmSearch.frx":297A
      PicArray54      =   "frmSearch.frx":2996
      PicArray55      =   "frmSearch.frx":29B2
      PicArray56      =   "frmSearch.frx":29CE
      PicArray57      =   "frmSearch.frx":29EA
      PicArray58      =   "frmSearch.frx":2A06
      PicArray59      =   "frmSearch.frx":2A22
      PicArray60      =   "frmSearch.frx":2A3E
      PicArray61      =   "frmSearch.frx":2A5A
      PicArray62      =   "frmSearch.frx":2A76
      PicArray63      =   "frmSearch.frx":2A92
      PicArray64      =   "frmSearch.frx":2AAE
      PicArray65      =   "frmSearch.frx":2ACA
      PicArray66      =   "frmSearch.frx":2AE6
      PicArray67      =   "frmSearch.frx":2B02
      PicArray68      =   "frmSearch.frx":2B1E
      PicArray69      =   "frmSearch.frx":2B3A
      PicArray70      =   "frmSearch.frx":2B56
      PicArray71      =   "frmSearch.frx":2B72
      PicArray72      =   "frmSearch.frx":2B8E
      PicArray73      =   "frmSearch.frx":2BAA
      PicArray74      =   "frmSearch.frx":2BC6
      PicArray75      =   "frmSearch.frx":2BE2
      PicArray76      =   "frmSearch.frx":2BFE
      PicArray77      =   "frmSearch.frx":2C1A
      PicArray78      =   "frmSearch.frx":2C36
      PicArray79      =   "frmSearch.frx":2C52
      PicArray80      =   "frmSearch.frx":2C6E
      PicArray81      =   "frmSearch.frx":2C8A
      PicArray82      =   "frmSearch.frx":2CA6
      PicArray83      =   "frmSearch.frx":2CC2
      PicArray84      =   "frmSearch.frx":2CDE
      PicArray85      =   "frmSearch.frx":2CFA
      PicArray86      =   "frmSearch.frx":2D16
      PicArray87      =   "frmSearch.frx":2D32
      PicArray88      =   "frmSearch.frx":2D4E
      PicArray89      =   "frmSearch.frx":2D6A
      PicArray90      =   "frmSearch.frx":2D86
      PicArray91      =   "frmSearch.frx":2DA2
      PicArray92      =   "frmSearch.frx":2DBE
      PicArray93      =   "frmSearch.frx":2DDA
      PicArray94      =   "frmSearch.frx":2DF6
      PicArray95      =   "frmSearch.frx":2E12
      PicArray96      =   "frmSearch.frx":2E2E
      PicArray97      =   "frmSearch.frx":2E4A
      PicArray98      =   "frmSearch.frx":2E66
      PicArray99      =   "frmSearch.frx":2E82
   End
   Begin CTLISTLibCtl.ctList lstResult 
      DragIcon        =   "frmSearch.frx":2E9E
      Height          =   4695
      Left            =   150
      TabIndex        =   12
      Top             =   2040
      Width           =   9345
      _Version        =   65536
      _ExtentX        =   16484
      _ExtentY        =   8281
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleBackImage  =   "frmSearch.frx":3768
      HeaderPicture   =   "frmSearch.frx":3784
      Picture         =   "frmSearch.frx":37A0
      CheckPicDown    =   "frmSearch.frx":37BC
      CheckPicUp      =   "frmSearch.frx":37D8
      CheckPicDisabled=   "frmSearch.frx":37F4
      BackImage       =   "frmSearch.frx":3810
      ShowHeader      =   -1  'True
      SortArrows      =   0   'False
      HeaderData      =   "frmSearch.frx":382C
      PicArray0       =   "frmSearch.frx":38F4
      PicArray1       =   "frmSearch.frx":3910
      PicArray2       =   "frmSearch.frx":392C
      PicArray3       =   "frmSearch.frx":3948
      PicArray4       =   "frmSearch.frx":3964
      PicArray5       =   "frmSearch.frx":3980
      PicArray6       =   "frmSearch.frx":399C
      PicArray7       =   "frmSearch.frx":39B8
      PicArray8       =   "frmSearch.frx":39D4
      PicArray9       =   "frmSearch.frx":39F0
      PicArray10      =   "frmSearch.frx":3A0C
      PicArray11      =   "frmSearch.frx":3A28
      PicArray12      =   "frmSearch.frx":3A44
      PicArray13      =   "frmSearch.frx":3A60
      PicArray14      =   "frmSearch.frx":3A7C
      PicArray15      =   "frmSearch.frx":3A98
      PicArray16      =   "frmSearch.frx":3AB4
      PicArray17      =   "frmSearch.frx":3AD0
      PicArray18      =   "frmSearch.frx":3AEC
      PicArray19      =   "frmSearch.frx":3B08
      PicArray20      =   "frmSearch.frx":3B24
      PicArray21      =   "frmSearch.frx":3B40
      PicArray22      =   "frmSearch.frx":3B5C
      PicArray23      =   "frmSearch.frx":3B78
      PicArray24      =   "frmSearch.frx":3B94
      PicArray25      =   "frmSearch.frx":3BB0
      PicArray26      =   "frmSearch.frx":3BCC
      PicArray27      =   "frmSearch.frx":3BE8
      PicArray28      =   "frmSearch.frx":3C04
      PicArray29      =   "frmSearch.frx":3C20
      PicArray30      =   "frmSearch.frx":3C3C
      PicArray31      =   "frmSearch.frx":3C58
      PicArray32      =   "frmSearch.frx":3C74
      PicArray33      =   "frmSearch.frx":3C90
      PicArray34      =   "frmSearch.frx":3CAC
      PicArray35      =   "frmSearch.frx":3CC8
      PicArray36      =   "frmSearch.frx":3CE4
      PicArray37      =   "frmSearch.frx":3D00
      PicArray38      =   "frmSearch.frx":3D1C
      PicArray39      =   "frmSearch.frx":3D38
      PicArray40      =   "frmSearch.frx":3D54
      PicArray41      =   "frmSearch.frx":3D70
      PicArray42      =   "frmSearch.frx":3D8C
      PicArray43      =   "frmSearch.frx":3DA8
      PicArray44      =   "frmSearch.frx":3DC4
      PicArray45      =   "frmSearch.frx":3DE0
      PicArray46      =   "frmSearch.frx":3DFC
      PicArray47      =   "frmSearch.frx":3E18
      PicArray48      =   "frmSearch.frx":3E34
      PicArray49      =   "frmSearch.frx":3E50
      PicArray50      =   "frmSearch.frx":3E6C
      PicArray51      =   "frmSearch.frx":3E88
      PicArray52      =   "frmSearch.frx":3EA4
      PicArray53      =   "frmSearch.frx":3EC0
      PicArray54      =   "frmSearch.frx":3EDC
      PicArray55      =   "frmSearch.frx":3EF8
      PicArray56      =   "frmSearch.frx":3F14
      PicArray57      =   "frmSearch.frx":3F30
      PicArray58      =   "frmSearch.frx":3F4C
      PicArray59      =   "frmSearch.frx":3F68
      PicArray60      =   "frmSearch.frx":3F84
      PicArray61      =   "frmSearch.frx":3FA0
      PicArray62      =   "frmSearch.frx":3FBC
      PicArray63      =   "frmSearch.frx":3FD8
      PicArray64      =   "frmSearch.frx":3FF4
      PicArray65      =   "frmSearch.frx":4010
      PicArray66      =   "frmSearch.frx":402C
      PicArray67      =   "frmSearch.frx":4048
      PicArray68      =   "frmSearch.frx":4064
      PicArray69      =   "frmSearch.frx":4080
      PicArray70      =   "frmSearch.frx":409C
      PicArray71      =   "frmSearch.frx":40B8
      PicArray72      =   "frmSearch.frx":40D4
      PicArray73      =   "frmSearch.frx":40F0
      PicArray74      =   "frmSearch.frx":410C
      PicArray75      =   "frmSearch.frx":4128
      PicArray76      =   "frmSearch.frx":4144
      PicArray77      =   "frmSearch.frx":4160
      PicArray78      =   "frmSearch.frx":417C
      PicArray79      =   "frmSearch.frx":4198
      PicArray80      =   "frmSearch.frx":41B4
      PicArray81      =   "frmSearch.frx":41D0
      PicArray82      =   "frmSearch.frx":41EC
      PicArray83      =   "frmSearch.frx":4208
      PicArray84      =   "frmSearch.frx":4224
      PicArray85      =   "frmSearch.frx":4240
      PicArray86      =   "frmSearch.frx":425C
      PicArray87      =   "frmSearch.frx":4278
      PicArray88      =   "frmSearch.frx":4294
      PicArray89      =   "frmSearch.frx":42B0
      PicArray90      =   "frmSearch.frx":42CC
      PicArray91      =   "frmSearch.frx":42E8
      PicArray92      =   "frmSearch.frx":4304
      PicArray93      =   "frmSearch.frx":4320
      PicArray94      =   "frmSearch.frx":433C
      PicArray95      =   "frmSearch.frx":4358
      PicArray96      =   "frmSearch.frx":4374
      PicArray97      =   "frmSearch.frx":4390
      PicArray98      =   "frmSearch.frx":43AC
      PicArray99      =   "frmSearch.frx":43C8
   End
   Begin VB.Label lblCnt 
      Height          =   315
      Left            =   5910
      TabIndex        =   11
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "검색어"
      Height          =   285
      Left            =   7920
      TabIndex        =   7
      Top             =   300
      Width           =   1485
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmVq01 As frmVq01


Sub setInitStatus()

Dim lRs As ADODB.Recordset

lst1_3.ClearList
lst1_4.ClearList

sSql = "SELECT a.subjnm,a.subj from ts01 a inner join ts02 b on a.subj=b.subj where "
sSql = sSql & " b.userid='" & gUserid & "' and b.startymd<=date_format(current_date,'%Y%m%d') and b.endymd>=date_format(current_date,'%Y%m%d') "
sSql = sSql & "  and a.subj not in (select pocketnm from tp01 where cond like '%select%')"

Set lRs = Fn_SQLExec(sSql).rs

Do Until lRs.EOF
    Call lst1_3.AddItem(lRs(0) & Chr(10) & lRs(1))
    lRs.MoveNext
Loop
lRs.Close

End Sub

Private Sub cmdSearch_Click()

Dim lRs As ADODB.Recordset
Dim lcnt As Long
Dim lcntWhere As Long
Dim strListSubj As String

Dim strListSubjNoWhere As String
Dim strListSubjWhere As String

Dim Pos As Long, pos2 As Long
Dim strCond As String
Dim strListCond As String

If txtSearch.Text = "" Then
    MsgBox "검색어를 입력하세요.", vbExclamation
    txtSearch.SetFocus
    
    Exit Sub
ElseIf txtSearch.Text = "다" Then
    MsgBox "검색어가 너무 짧습니다.", vbExclamation
    txtSearch.SetFocus
    
    Exit Sub
ElseIf lst1_4.ListCount() = 0 Then
    MsgBox "과목을 선택해 주세요.", vbExclamation
    cmd2R2La.SetFocus
    
    Exit Sub
End If

lblCnt.Caption = ""

lstResult.ClearList

'strListSubj = itemSeries1(lst1_4, Chr(10))
strListSubj = itemSeries2(lst1_4, Chr(10))

'sSql = "select cond,pocketnm from tp01 where pocketnm in (" & itemSeries1(lst1_4, Chr(10)) & ") "
'
'lcnt = Fn_SQLExec(sSql).nrow
'
'Set lRs = Fn_SQLExec(sSql).rs
'
'strListSubjNoWhere = "''"
'
'Do Until lRs.EOF
'    strCond = lRs(0).Value
'    If InStr(strCond, "where") > 0 Then
'        lcntWhere = lcntWhere + 1
'        Pos = InStrRev(strCond, "where")
'        pos2 = InStrRev(strCond, ")")
'        strCond = Mid(strCond, Pos, pos2 - Pos)
'        strCond = Replace(strCond, "where", "")
'        strListCond = strListCond & " or " & strCond
'    Else
'        strListSubjNoWhere = strListSubjNoWhere & ",'" & lRs(1).Value & "'"
'    End If
'
'    lRs.MoveNext
'Loop
'lRs.Close

strListSubjNoWhere = strListSubj

sSql = "SELECT ta.subj,ta.seq,ta.quiz,ta.a from vq01 ta left outer join th01 tb on (ta.subj=tb.subj and ta.seq=tb.seq and tb.userid='" & gUserid & "' ) "
If lcntWhere = 0 Then
    sSql = sSql & " where ta.subj in (" & strListSubjNoWhere & ") "
Else
    sSql = sSql & " where (ta.subj in (" & strListSubjNoWhere & ") " & Replace(strListCond, " subj=", " ta.subj=") & ")"
End If

Select Case lstColumn.ListIndex

    Case 0 '문제+보기a
        sSql = sSql & " and match(ta.quiz,ta.a) against('" + txtSearch.Text + "') "
    Case 1
        sSql = sSql & " and ta.quiz like '%" + txtSearch.Text + "%' "
    Case 2
        sSql = sSql & " and ta.a like '%" + txtSearch.Text + "%' "
    Case 3
        sSql = sSql & " and ta.b like '%" + txtSearch.Text + "%' "
    Case 4
        sSql = sSql & " and ta.c like '%" + txtSearch.Text + "%' "
    Case 5
        sSql = sSql & " and ta.d like '%" + txtSearch.Text + "%' "
    Case 6
        sSql = sSql & " and ta.e like '%" + txtSearch.Text + "%' "
    Case 7
        sSql = sSql & " and ta.hint like '%" + txtSearch.Text + "%' "
    Case 8
        sSql = sSql & " and tb.hint like '%" + txtSearch.Text + "%' "
    Case Else
        'MsgBox "기본조회 조건으로 처리합니다.", vbExclamation
        sSql = sSql & " and match(ta.quiz,ta.a) against('" + txtSearch.Text + "') "
End Select

Set lRs = Fn_SQLExec(sSql).rs

lcnt = Fn_SQLExec(sSql).nrow

lblCnt.Caption = lcnt

Do Until lRs.EOF
    Call lstResult.AddItem(lRs(0) & Chr(10) & lRs(1) & Chr(10) & lRs(2) & Chr(10) & lRs(3))
    lRs.MoveNext
Loop
lRs.Close

End Sub

Private Sub ctList1_Click()

End Sub

Private Sub cmdViewPop_Click()

Dim formVq01 As New frmVq01
Dim retVal As VbMsgBoxResult

Load formVq01

Dim sSubj As String, sSeq As String

sSubj = lstResult.ListColumnText(lstResult.ListIndex, 1)
sSeq = lstResult.ListColumnText(lstResult.ListIndex, 2)

formVq01.txtSubj = sSubj
formVq01.txtSeq = sSeq

formVq01.cmdTH01.Enabled = False
If InStr(gUserid, "@") > 0 Then
    formVq01.cmdSave.Enabled = False
End If

Call formVq01.fillAll

'Set formVq01.parent = Me
TmrAfterTTS_exit = True
formVq01.Show vbModal
TmrAfterTTS_exit = False
Unload formVq01
Set formVq01 = Nothing

End Sub

Private Sub Form_Load()

    setInitStatus
    
    lstColumn.Text = "문제+보기a"
    
End Sub

Private Sub cmd2L2R_Click()
lst1_4_DropList lst1_4.ListCount, 0
End Sub

Private Sub cmd2L2Ra_Click()
Dim i As Integer
For i = 0 To lst1_3.ListCount - 1
    lst1_3.ListSelect(i) = True
Next
cmd2L2R_Click

txtSearch.SetFocus

End Sub

Private Sub cmd2R2L_Click()
lst1_3_DropList lst1_3.ListCount, 0
End Sub

Private Sub cmd2R2La_Click()
Dim i As Integer
For i = 0 To lst1_4.ListCount - 1
    lst1_4.ListSelect(i) = True
Next
cmd2R2L_Click

End Sub

Private Sub lst1_3_DblClick()
cmd2L2R_Click
End Sub

Private Sub lst1_3_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "lst1_4" Then
        lst1_3.DragDrop (X / Screen.TwipsPerPixelX), (Y / Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub lst1_3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If Source.Name = "lst1_4" Then
        lst1_3.DragOver (X / Screen.TwipsPerPixelX), (Y / Screen.TwipsPerPixelY), State
    End If
End Sub

Private Sub lst1_3_DropList(ByVal nIndex As Long, ByVal nColumn As Integer)
Dim strText As String
Dim i As Integer
    
    
    For i = lst1_4.ListCount - 1 To 0 Step -1
        If lst1_4.ListSelect(i) Then
            strText = lst1_4.ListText(i)
            lst1_3.InsertItem strText, nIndex
            lst1_4.removeItem (i)
        End If
    Next
    


End Sub

Private Sub lst1_3_StartDragOut()
lst1_3.Drag vbBeginDrag
End Sub

Private Sub lst1_4_DblClick()
cmd2R2L_Click
End Sub

Private Sub lst1_4_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "lst1_3" Then
        lst1_4.DragDrop (X / Screen.TwipsPerPixelX), (Y / Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub lst1_4_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If Source.Name = "lst1_3" Then
        lst1_4.DragOver (X / Screen.TwipsPerPixelX), (Y / Screen.TwipsPerPixelY), State
    End If
End Sub

Private Sub lst1_4_DropList(ByVal nIndex As Long, ByVal nColumn As Integer)
Dim strText As String
Dim i As Integer
    
    
    For i = lst1_3.ListCount - 1 To 0 Step -1
        If lst1_3.ListSelect(i) Then
            strText = lst1_3.ListText(i)
            lst1_4.InsertItem strText, nIndex
            lst1_3.removeItem (i)
        End If
    Next
    


End Sub

Private Sub lstResult_DblClick()
    cmdViewPop_Click
End Sub
