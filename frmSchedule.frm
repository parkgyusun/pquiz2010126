VERSION 5.00
Object = "{D8D562C3-878C-11D2-943F-444553540000}#1.0#0"; "ctlist.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSchedule 
   BorderStyle     =   0  'None
   Caption         =   "계획"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "FORM"
   Begin VB.Frame fra 
      Caption         =   "1/4"
      Height          =   6270
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   10455
      Begin VB.CheckBox chk1_6 
         Caption         =   "안푼문제"
         Height          =   375
         Left            =   6795
         TabIndex        =   84
         Top             =   4905
         Width           =   3255
      End
      Begin CTLISTLibCtl.ctList lst1_1 
         DragIcon        =   "frmSchedule.frx":0442
         Height          =   1725
         Left            =   405
         TabIndex        =   60
         Top             =   1215
         Width           =   2168
         _Version        =   65536
         _ExtentX        =   3824
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
         TitleBackImage  =   "frmSchedule.frx":0D0C
         HeaderPicture   =   "frmSchedule.frx":0D28
         Picture         =   "frmSchedule.frx":0D44
         CheckPicDown    =   "frmSchedule.frx":0D60
         CheckPicUp      =   "frmSchedule.frx":0D7C
         CheckPicDisabled=   "frmSchedule.frx":0D98
         BackImage       =   "frmSchedule.frx":0DB4
         HeaderOffset    =   1
         ShowHeader      =   -1  'True
         MultiSelect     =   -1  'True
         HeaderData      =   "frmSchedule.frx":0DD0
         PicArray0       =   "frmSchedule.frx":0E55
         PicArray1       =   "frmSchedule.frx":0E71
         PicArray2       =   "frmSchedule.frx":0E8D
         PicArray3       =   "frmSchedule.frx":0EA9
         PicArray4       =   "frmSchedule.frx":0EC5
         PicArray5       =   "frmSchedule.frx":0EE1
         PicArray6       =   "frmSchedule.frx":0EFD
         PicArray7       =   "frmSchedule.frx":0F19
         PicArray8       =   "frmSchedule.frx":0F35
         PicArray9       =   "frmSchedule.frx":0F51
         PicArray10      =   "frmSchedule.frx":0F6D
         PicArray11      =   "frmSchedule.frx":0F89
         PicArray12      =   "frmSchedule.frx":0FA5
         PicArray13      =   "frmSchedule.frx":0FC1
         PicArray14      =   "frmSchedule.frx":0FDD
         PicArray15      =   "frmSchedule.frx":0FF9
         PicArray16      =   "frmSchedule.frx":1015
         PicArray17      =   "frmSchedule.frx":1031
         PicArray18      =   "frmSchedule.frx":104D
         PicArray19      =   "frmSchedule.frx":1069
         PicArray20      =   "frmSchedule.frx":1085
         PicArray21      =   "frmSchedule.frx":10A1
         PicArray22      =   "frmSchedule.frx":10BD
         PicArray23      =   "frmSchedule.frx":10D9
         PicArray24      =   "frmSchedule.frx":10F5
         PicArray25      =   "frmSchedule.frx":1111
         PicArray26      =   "frmSchedule.frx":112D
         PicArray27      =   "frmSchedule.frx":1149
         PicArray28      =   "frmSchedule.frx":1165
         PicArray29      =   "frmSchedule.frx":1181
         PicArray30      =   "frmSchedule.frx":119D
         PicArray31      =   "frmSchedule.frx":11B9
         PicArray32      =   "frmSchedule.frx":11D5
         PicArray33      =   "frmSchedule.frx":11F1
         PicArray34      =   "frmSchedule.frx":120D
         PicArray35      =   "frmSchedule.frx":1229
         PicArray36      =   "frmSchedule.frx":1245
         PicArray37      =   "frmSchedule.frx":1261
         PicArray38      =   "frmSchedule.frx":127D
         PicArray39      =   "frmSchedule.frx":1299
         PicArray40      =   "frmSchedule.frx":12B5
         PicArray41      =   "frmSchedule.frx":12D1
         PicArray42      =   "frmSchedule.frx":12ED
         PicArray43      =   "frmSchedule.frx":1309
         PicArray44      =   "frmSchedule.frx":1325
         PicArray45      =   "frmSchedule.frx":1341
         PicArray46      =   "frmSchedule.frx":135D
         PicArray47      =   "frmSchedule.frx":1379
         PicArray48      =   "frmSchedule.frx":1395
         PicArray49      =   "frmSchedule.frx":13B1
         PicArray50      =   "frmSchedule.frx":13CD
         PicArray51      =   "frmSchedule.frx":13E9
         PicArray52      =   "frmSchedule.frx":1405
         PicArray53      =   "frmSchedule.frx":1421
         PicArray54      =   "frmSchedule.frx":143D
         PicArray55      =   "frmSchedule.frx":1459
         PicArray56      =   "frmSchedule.frx":1475
         PicArray57      =   "frmSchedule.frx":1491
         PicArray58      =   "frmSchedule.frx":14AD
         PicArray59      =   "frmSchedule.frx":14C9
         PicArray60      =   "frmSchedule.frx":14E5
         PicArray61      =   "frmSchedule.frx":1501
         PicArray62      =   "frmSchedule.frx":151D
         PicArray63      =   "frmSchedule.frx":1539
         PicArray64      =   "frmSchedule.frx":1555
         PicArray65      =   "frmSchedule.frx":1571
         PicArray66      =   "frmSchedule.frx":158D
         PicArray67      =   "frmSchedule.frx":15A9
         PicArray68      =   "frmSchedule.frx":15C5
         PicArray69      =   "frmSchedule.frx":15E1
         PicArray70      =   "frmSchedule.frx":15FD
         PicArray71      =   "frmSchedule.frx":1619
         PicArray72      =   "frmSchedule.frx":1635
         PicArray73      =   "frmSchedule.frx":1651
         PicArray74      =   "frmSchedule.frx":166D
         PicArray75      =   "frmSchedule.frx":1689
         PicArray76      =   "frmSchedule.frx":16A5
         PicArray77      =   "frmSchedule.frx":16C1
         PicArray78      =   "frmSchedule.frx":16DD
         PicArray79      =   "frmSchedule.frx":16F9
         PicArray80      =   "frmSchedule.frx":1715
         PicArray81      =   "frmSchedule.frx":1731
         PicArray82      =   "frmSchedule.frx":174D
         PicArray83      =   "frmSchedule.frx":1769
         PicArray84      =   "frmSchedule.frx":1785
         PicArray85      =   "frmSchedule.frx":17A1
         PicArray86      =   "frmSchedule.frx":17BD
         PicArray87      =   "frmSchedule.frx":17D9
         PicArray88      =   "frmSchedule.frx":17F5
         PicArray89      =   "frmSchedule.frx":1811
         PicArray90      =   "frmSchedule.frx":182D
         PicArray91      =   "frmSchedule.frx":1849
         PicArray92      =   "frmSchedule.frx":1865
         PicArray93      =   "frmSchedule.frx":1881
         PicArray94      =   "frmSchedule.frx":189D
         PicArray95      =   "frmSchedule.frx":18B9
         PicArray96      =   "frmSchedule.frx":18D5
         PicArray97      =   "frmSchedule.frx":18F1
         PicArray98      =   "frmSchedule.frx":190D
         PicArray99      =   "frmSchedule.frx":1929
      End
      Begin VB.Frame fra1_1 
         Height          =   915
         Left            =   3645
         TabIndex        =   33
         Top             =   5265
         Width           =   2715
         Begin VB.CheckBox chk1_5 
            Caption         =   "완료일"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   37
            Top             =   495
            Width           =   915
         End
         Begin VB.CheckBox chk1_5 
            Caption         =   "시작일"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   36
            Top             =   225
            Width           =   1005
         End
         Begin POCKETQUIZ.numText nTxt1_3 
            Height          =   285
            Left            =   1215
            TabIndex        =   34
            Top             =   495
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FormatMask      =   ""
            MinVal          =   0
            MaxVal          =   99991231
            MaxLength       =   8
            FontSize        =   9
            FontName        =   "굴림"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   0
            EndProperty
            AllowNull       =   -1  'True
         End
         Begin POCKETQUIZ.numText nTxt1_2 
            Height          =   285
            Left            =   1215
            TabIndex        =   35
            Top             =   225
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FormatMask      =   ""
            MinVal          =   0
            MaxVal          =   99991231
            MaxLength       =   8
            FontSize        =   9
            FontName        =   "굴림"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   0
            EndProperty
            AllowNull       =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp1 
            Height          =   285
            Left            =   2295
            TabIndex        =   86
            Top             =   225
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            _Version        =   393216
            Format          =   88342529
            CurrentDate     =   38188
         End
         Begin MSComCtl2.DTPicker dtp2 
            Height          =   285
            Left            =   2295
            TabIndex        =   87
            Top             =   495
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            _Version        =   393216
            Format          =   88342529
            CurrentDate     =   38188
         End
      End
      Begin VB.CheckBox chk1_4 
         Caption         =   "복습 일정"
         Height          =   285
         Left            =   2295
         TabIndex        =   29
         Top             =   5580
         Width           =   1230
      End
      Begin POCKETQUIZ.numText nTxt1_1 
         Height          =   285
         Index           =   0
         Left            =   7245
         TabIndex        =   19
         Top             =   1575
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinVal          =   0
         FontSize        =   9
         FontName        =   "굴림"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Object.CausesValidation=   0   'False
      End
      Begin VB.CheckBox chk1_3 
         Caption         =   "틀린비율"
         Height          =   375
         Left            =   6795
         TabIndex        =   18
         Top             =   3735
         Width           =   3255
      End
      Begin VB.CheckBox chk1_2 
         Caption         =   "틀린수"
         Height          =   375
         Left            =   6795
         TabIndex        =   17
         Top             =   2452
         Width           =   3255
      End
      Begin VB.CheckBox chk1_1 
         Caption         =   "맞은수"
         Height          =   375
         Left            =   6795
         TabIndex        =   16
         Top             =   1170
         Width           =   3255
      End
      Begin VB.CommandButton cmd2R2La 
         Caption         =   "<<"
         Height          =   330
         Left            =   2745
         TabIndex        =   12
         Top             =   4680
         Width           =   915
      End
      Begin VB.CommandButton cmd2R2L 
         Caption         =   "<"
         Height          =   330
         Left            =   2745
         TabIndex        =   11
         Top             =   4320
         Width           =   915
      End
      Begin VB.CommandButton cmd2L2Ra 
         Caption         =   ">>"
         Height          =   330
         Left            =   2745
         TabIndex        =   10
         Top             =   3960
         Width           =   915
      End
      Begin VB.CommandButton cmd2L2R 
         Caption         =   ">"
         Height          =   330
         Left            =   2745
         TabIndex        =   9
         Top             =   3600
         Width           =   915
      End
      Begin VB.CommandButton cmd1R2La 
         Caption         =   "<<"
         Height          =   330
         Left            =   2745
         TabIndex        =   8
         Top             =   2520
         Width           =   915
      End
      Begin VB.CommandButton cmd1R2L 
         Caption         =   "<"
         Height          =   330
         Left            =   2745
         TabIndex        =   7
         Top             =   2160
         Width           =   915
      End
      Begin VB.CommandButton cmd1L2Ra 
         Caption         =   ">>"
         Height          =   330
         Left            =   2745
         TabIndex        =   6
         Top             =   1800
         Width           =   915
      End
      Begin VB.CommandButton cmd1L2R 
         Caption         =   ">"
         Height          =   330
         Left            =   2745
         TabIndex        =   5
         Top             =   1440
         Width           =   915
      End
      Begin POCKETQUIZ.numText nTxt1_1 
         Height          =   285
         Index           =   1
         Left            =   7245
         TabIndex        =   23
         Top             =   2835
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinVal          =   0
         FontSize        =   9
         FontName        =   "굴림"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Object.CausesValidation=   0   'False
      End
      Begin POCKETQUIZ.numText nTxt1_1 
         Height          =   285
         Index           =   2
         Left            =   7245
         TabIndex        =   24
         Top             =   4185
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinVal          =   0
         MaxVal          =   100
         FontSize        =   9
         FontName        =   "굴림"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Object.CausesValidation=   0   'False
      End
      Begin CTLISTLibCtl.ctList lst1_2 
         DragIcon        =   "frmSchedule.frx":1945
         Height          =   1725
         Left            =   3735
         TabIndex        =   61
         Top             =   1215
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
         TitleBackImage  =   "frmSchedule.frx":220F
         HeaderPicture   =   "frmSchedule.frx":222B
         Picture         =   "frmSchedule.frx":2247
         CheckPicDown    =   "frmSchedule.frx":2263
         CheckPicUp      =   "frmSchedule.frx":227F
         CheckPicDisabled=   "frmSchedule.frx":229B
         BackImage       =   "frmSchedule.frx":22B7
         ShowHeader      =   -1  'True
         MultiSelect     =   -1  'True
         HeaderData      =   "frmSchedule.frx":22D3
         PicArray0       =   "frmSchedule.frx":2358
         PicArray1       =   "frmSchedule.frx":2374
         PicArray2       =   "frmSchedule.frx":2390
         PicArray3       =   "frmSchedule.frx":23AC
         PicArray4       =   "frmSchedule.frx":23C8
         PicArray5       =   "frmSchedule.frx":23E4
         PicArray6       =   "frmSchedule.frx":2400
         PicArray7       =   "frmSchedule.frx":241C
         PicArray8       =   "frmSchedule.frx":2438
         PicArray9       =   "frmSchedule.frx":2454
         PicArray10      =   "frmSchedule.frx":2470
         PicArray11      =   "frmSchedule.frx":248C
         PicArray12      =   "frmSchedule.frx":24A8
         PicArray13      =   "frmSchedule.frx":24C4
         PicArray14      =   "frmSchedule.frx":24E0
         PicArray15      =   "frmSchedule.frx":24FC
         PicArray16      =   "frmSchedule.frx":2518
         PicArray17      =   "frmSchedule.frx":2534
         PicArray18      =   "frmSchedule.frx":2550
         PicArray19      =   "frmSchedule.frx":256C
         PicArray20      =   "frmSchedule.frx":2588
         PicArray21      =   "frmSchedule.frx":25A4
         PicArray22      =   "frmSchedule.frx":25C0
         PicArray23      =   "frmSchedule.frx":25DC
         PicArray24      =   "frmSchedule.frx":25F8
         PicArray25      =   "frmSchedule.frx":2614
         PicArray26      =   "frmSchedule.frx":2630
         PicArray27      =   "frmSchedule.frx":264C
         PicArray28      =   "frmSchedule.frx":2668
         PicArray29      =   "frmSchedule.frx":2684
         PicArray30      =   "frmSchedule.frx":26A0
         PicArray31      =   "frmSchedule.frx":26BC
         PicArray32      =   "frmSchedule.frx":26D8
         PicArray33      =   "frmSchedule.frx":26F4
         PicArray34      =   "frmSchedule.frx":2710
         PicArray35      =   "frmSchedule.frx":272C
         PicArray36      =   "frmSchedule.frx":2748
         PicArray37      =   "frmSchedule.frx":2764
         PicArray38      =   "frmSchedule.frx":2780
         PicArray39      =   "frmSchedule.frx":279C
         PicArray40      =   "frmSchedule.frx":27B8
         PicArray41      =   "frmSchedule.frx":27D4
         PicArray42      =   "frmSchedule.frx":27F0
         PicArray43      =   "frmSchedule.frx":280C
         PicArray44      =   "frmSchedule.frx":2828
         PicArray45      =   "frmSchedule.frx":2844
         PicArray46      =   "frmSchedule.frx":2860
         PicArray47      =   "frmSchedule.frx":287C
         PicArray48      =   "frmSchedule.frx":2898
         PicArray49      =   "frmSchedule.frx":28B4
         PicArray50      =   "frmSchedule.frx":28D0
         PicArray51      =   "frmSchedule.frx":28EC
         PicArray52      =   "frmSchedule.frx":2908
         PicArray53      =   "frmSchedule.frx":2924
         PicArray54      =   "frmSchedule.frx":2940
         PicArray55      =   "frmSchedule.frx":295C
         PicArray56      =   "frmSchedule.frx":2978
         PicArray57      =   "frmSchedule.frx":2994
         PicArray58      =   "frmSchedule.frx":29B0
         PicArray59      =   "frmSchedule.frx":29CC
         PicArray60      =   "frmSchedule.frx":29E8
         PicArray61      =   "frmSchedule.frx":2A04
         PicArray62      =   "frmSchedule.frx":2A20
         PicArray63      =   "frmSchedule.frx":2A3C
         PicArray64      =   "frmSchedule.frx":2A58
         PicArray65      =   "frmSchedule.frx":2A74
         PicArray66      =   "frmSchedule.frx":2A90
         PicArray67      =   "frmSchedule.frx":2AAC
         PicArray68      =   "frmSchedule.frx":2AC8
         PicArray69      =   "frmSchedule.frx":2AE4
         PicArray70      =   "frmSchedule.frx":2B00
         PicArray71      =   "frmSchedule.frx":2B1C
         PicArray72      =   "frmSchedule.frx":2B38
         PicArray73      =   "frmSchedule.frx":2B54
         PicArray74      =   "frmSchedule.frx":2B70
         PicArray75      =   "frmSchedule.frx":2B8C
         PicArray76      =   "frmSchedule.frx":2BA8
         PicArray77      =   "frmSchedule.frx":2BC4
         PicArray78      =   "frmSchedule.frx":2BE0
         PicArray79      =   "frmSchedule.frx":2BFC
         PicArray80      =   "frmSchedule.frx":2C18
         PicArray81      =   "frmSchedule.frx":2C34
         PicArray82      =   "frmSchedule.frx":2C50
         PicArray83      =   "frmSchedule.frx":2C6C
         PicArray84      =   "frmSchedule.frx":2C88
         PicArray85      =   "frmSchedule.frx":2CA4
         PicArray86      =   "frmSchedule.frx":2CC0
         PicArray87      =   "frmSchedule.frx":2CDC
         PicArray88      =   "frmSchedule.frx":2CF8
         PicArray89      =   "frmSchedule.frx":2D14
         PicArray90      =   "frmSchedule.frx":2D30
         PicArray91      =   "frmSchedule.frx":2D4C
         PicArray92      =   "frmSchedule.frx":2D68
         PicArray93      =   "frmSchedule.frx":2D84
         PicArray94      =   "frmSchedule.frx":2DA0
         PicArray95      =   "frmSchedule.frx":2DBC
         PicArray96      =   "frmSchedule.frx":2DD8
         PicArray97      =   "frmSchedule.frx":2DF4
         PicArray98      =   "frmSchedule.frx":2E10
         PicArray99      =   "frmSchedule.frx":2E2C
      End
      Begin CTLISTLibCtl.ctList lst1_3 
         DragIcon        =   "frmSchedule.frx":2E48
         Height          =   1725
         Left            =   405
         TabIndex        =   62
         Top             =   3420
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
         TitleBackImage  =   "frmSchedule.frx":3712
         HeaderPicture   =   "frmSchedule.frx":372E
         Picture         =   "frmSchedule.frx":374A
         CheckPicDown    =   "frmSchedule.frx":3766
         CheckPicUp      =   "frmSchedule.frx":3782
         CheckPicDisabled=   "frmSchedule.frx":379E
         BackImage       =   "frmSchedule.frx":37BA
         ShowHeader      =   -1  'True
         MultiSelect     =   -1  'True
         HeaderData      =   "frmSchedule.frx":37D6
         PicArray0       =   "frmSchedule.frx":3853
         PicArray1       =   "frmSchedule.frx":386F
         PicArray2       =   "frmSchedule.frx":388B
         PicArray3       =   "frmSchedule.frx":38A7
         PicArray4       =   "frmSchedule.frx":38C3
         PicArray5       =   "frmSchedule.frx":38DF
         PicArray6       =   "frmSchedule.frx":38FB
         PicArray7       =   "frmSchedule.frx":3917
         PicArray8       =   "frmSchedule.frx":3933
         PicArray9       =   "frmSchedule.frx":394F
         PicArray10      =   "frmSchedule.frx":396B
         PicArray11      =   "frmSchedule.frx":3987
         PicArray12      =   "frmSchedule.frx":39A3
         PicArray13      =   "frmSchedule.frx":39BF
         PicArray14      =   "frmSchedule.frx":39DB
         PicArray15      =   "frmSchedule.frx":39F7
         PicArray16      =   "frmSchedule.frx":3A13
         PicArray17      =   "frmSchedule.frx":3A2F
         PicArray18      =   "frmSchedule.frx":3A4B
         PicArray19      =   "frmSchedule.frx":3A67
         PicArray20      =   "frmSchedule.frx":3A83
         PicArray21      =   "frmSchedule.frx":3A9F
         PicArray22      =   "frmSchedule.frx":3ABB
         PicArray23      =   "frmSchedule.frx":3AD7
         PicArray24      =   "frmSchedule.frx":3AF3
         PicArray25      =   "frmSchedule.frx":3B0F
         PicArray26      =   "frmSchedule.frx":3B2B
         PicArray27      =   "frmSchedule.frx":3B47
         PicArray28      =   "frmSchedule.frx":3B63
         PicArray29      =   "frmSchedule.frx":3B7F
         PicArray30      =   "frmSchedule.frx":3B9B
         PicArray31      =   "frmSchedule.frx":3BB7
         PicArray32      =   "frmSchedule.frx":3BD3
         PicArray33      =   "frmSchedule.frx":3BEF
         PicArray34      =   "frmSchedule.frx":3C0B
         PicArray35      =   "frmSchedule.frx":3C27
         PicArray36      =   "frmSchedule.frx":3C43
         PicArray37      =   "frmSchedule.frx":3C5F
         PicArray38      =   "frmSchedule.frx":3C7B
         PicArray39      =   "frmSchedule.frx":3C97
         PicArray40      =   "frmSchedule.frx":3CB3
         PicArray41      =   "frmSchedule.frx":3CCF
         PicArray42      =   "frmSchedule.frx":3CEB
         PicArray43      =   "frmSchedule.frx":3D07
         PicArray44      =   "frmSchedule.frx":3D23
         PicArray45      =   "frmSchedule.frx":3D3F
         PicArray46      =   "frmSchedule.frx":3D5B
         PicArray47      =   "frmSchedule.frx":3D77
         PicArray48      =   "frmSchedule.frx":3D93
         PicArray49      =   "frmSchedule.frx":3DAF
         PicArray50      =   "frmSchedule.frx":3DCB
         PicArray51      =   "frmSchedule.frx":3DE7
         PicArray52      =   "frmSchedule.frx":3E03
         PicArray53      =   "frmSchedule.frx":3E1F
         PicArray54      =   "frmSchedule.frx":3E3B
         PicArray55      =   "frmSchedule.frx":3E57
         PicArray56      =   "frmSchedule.frx":3E73
         PicArray57      =   "frmSchedule.frx":3E8F
         PicArray58      =   "frmSchedule.frx":3EAB
         PicArray59      =   "frmSchedule.frx":3EC7
         PicArray60      =   "frmSchedule.frx":3EE3
         PicArray61      =   "frmSchedule.frx":3EFF
         PicArray62      =   "frmSchedule.frx":3F1B
         PicArray63      =   "frmSchedule.frx":3F37
         PicArray64      =   "frmSchedule.frx":3F53
         PicArray65      =   "frmSchedule.frx":3F6F
         PicArray66      =   "frmSchedule.frx":3F8B
         PicArray67      =   "frmSchedule.frx":3FA7
         PicArray68      =   "frmSchedule.frx":3FC3
         PicArray69      =   "frmSchedule.frx":3FDF
         PicArray70      =   "frmSchedule.frx":3FFB
         PicArray71      =   "frmSchedule.frx":4017
         PicArray72      =   "frmSchedule.frx":4033
         PicArray73      =   "frmSchedule.frx":404F
         PicArray74      =   "frmSchedule.frx":406B
         PicArray75      =   "frmSchedule.frx":4087
         PicArray76      =   "frmSchedule.frx":40A3
         PicArray77      =   "frmSchedule.frx":40BF
         PicArray78      =   "frmSchedule.frx":40DB
         PicArray79      =   "frmSchedule.frx":40F7
         PicArray80      =   "frmSchedule.frx":4113
         PicArray81      =   "frmSchedule.frx":412F
         PicArray82      =   "frmSchedule.frx":414B
         PicArray83      =   "frmSchedule.frx":4167
         PicArray84      =   "frmSchedule.frx":4183
         PicArray85      =   "frmSchedule.frx":419F
         PicArray86      =   "frmSchedule.frx":41BB
         PicArray87      =   "frmSchedule.frx":41D7
         PicArray88      =   "frmSchedule.frx":41F3
         PicArray89      =   "frmSchedule.frx":420F
         PicArray90      =   "frmSchedule.frx":422B
         PicArray91      =   "frmSchedule.frx":4247
         PicArray92      =   "frmSchedule.frx":4263
         PicArray93      =   "frmSchedule.frx":427F
         PicArray94      =   "frmSchedule.frx":429B
         PicArray95      =   "frmSchedule.frx":42B7
         PicArray96      =   "frmSchedule.frx":42D3
         PicArray97      =   "frmSchedule.frx":42EF
         PicArray98      =   "frmSchedule.frx":430B
         PicArray99      =   "frmSchedule.frx":4327
      End
      Begin CTLISTLibCtl.ctList lst1_4 
         DragIcon        =   "frmSchedule.frx":4343
         Height          =   1725
         Left            =   3735
         TabIndex        =   63
         Top             =   3420
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
         TitleBackImage  =   "frmSchedule.frx":4C0D
         HeaderPicture   =   "frmSchedule.frx":4C29
         Picture         =   "frmSchedule.frx":4C45
         CheckPicDown    =   "frmSchedule.frx":4C61
         CheckPicUp      =   "frmSchedule.frx":4C7D
         CheckPicDisabled=   "frmSchedule.frx":4C99
         BackImage       =   "frmSchedule.frx":4CB5
         ShowHeader      =   -1  'True
         MultiSelect     =   -1  'True
         SortArrows      =   0   'False
         HeaderData      =   "frmSchedule.frx":4CD1
         PicArray0       =   "frmSchedule.frx":4D4E
         PicArray1       =   "frmSchedule.frx":4D6A
         PicArray2       =   "frmSchedule.frx":4D86
         PicArray3       =   "frmSchedule.frx":4DA2
         PicArray4       =   "frmSchedule.frx":4DBE
         PicArray5       =   "frmSchedule.frx":4DDA
         PicArray6       =   "frmSchedule.frx":4DF6
         PicArray7       =   "frmSchedule.frx":4E12
         PicArray8       =   "frmSchedule.frx":4E2E
         PicArray9       =   "frmSchedule.frx":4E4A
         PicArray10      =   "frmSchedule.frx":4E66
         PicArray11      =   "frmSchedule.frx":4E82
         PicArray12      =   "frmSchedule.frx":4E9E
         PicArray13      =   "frmSchedule.frx":4EBA
         PicArray14      =   "frmSchedule.frx":4ED6
         PicArray15      =   "frmSchedule.frx":4EF2
         PicArray16      =   "frmSchedule.frx":4F0E
         PicArray17      =   "frmSchedule.frx":4F2A
         PicArray18      =   "frmSchedule.frx":4F46
         PicArray19      =   "frmSchedule.frx":4F62
         PicArray20      =   "frmSchedule.frx":4F7E
         PicArray21      =   "frmSchedule.frx":4F9A
         PicArray22      =   "frmSchedule.frx":4FB6
         PicArray23      =   "frmSchedule.frx":4FD2
         PicArray24      =   "frmSchedule.frx":4FEE
         PicArray25      =   "frmSchedule.frx":500A
         PicArray26      =   "frmSchedule.frx":5026
         PicArray27      =   "frmSchedule.frx":5042
         PicArray28      =   "frmSchedule.frx":505E
         PicArray29      =   "frmSchedule.frx":507A
         PicArray30      =   "frmSchedule.frx":5096
         PicArray31      =   "frmSchedule.frx":50B2
         PicArray32      =   "frmSchedule.frx":50CE
         PicArray33      =   "frmSchedule.frx":50EA
         PicArray34      =   "frmSchedule.frx":5106
         PicArray35      =   "frmSchedule.frx":5122
         PicArray36      =   "frmSchedule.frx":513E
         PicArray37      =   "frmSchedule.frx":515A
         PicArray38      =   "frmSchedule.frx":5176
         PicArray39      =   "frmSchedule.frx":5192
         PicArray40      =   "frmSchedule.frx":51AE
         PicArray41      =   "frmSchedule.frx":51CA
         PicArray42      =   "frmSchedule.frx":51E6
         PicArray43      =   "frmSchedule.frx":5202
         PicArray44      =   "frmSchedule.frx":521E
         PicArray45      =   "frmSchedule.frx":523A
         PicArray46      =   "frmSchedule.frx":5256
         PicArray47      =   "frmSchedule.frx":5272
         PicArray48      =   "frmSchedule.frx":528E
         PicArray49      =   "frmSchedule.frx":52AA
         PicArray50      =   "frmSchedule.frx":52C6
         PicArray51      =   "frmSchedule.frx":52E2
         PicArray52      =   "frmSchedule.frx":52FE
         PicArray53      =   "frmSchedule.frx":531A
         PicArray54      =   "frmSchedule.frx":5336
         PicArray55      =   "frmSchedule.frx":5352
         PicArray56      =   "frmSchedule.frx":536E
         PicArray57      =   "frmSchedule.frx":538A
         PicArray58      =   "frmSchedule.frx":53A6
         PicArray59      =   "frmSchedule.frx":53C2
         PicArray60      =   "frmSchedule.frx":53DE
         PicArray61      =   "frmSchedule.frx":53FA
         PicArray62      =   "frmSchedule.frx":5416
         PicArray63      =   "frmSchedule.frx":5432
         PicArray64      =   "frmSchedule.frx":544E
         PicArray65      =   "frmSchedule.frx":546A
         PicArray66      =   "frmSchedule.frx":5486
         PicArray67      =   "frmSchedule.frx":54A2
         PicArray68      =   "frmSchedule.frx":54BE
         PicArray69      =   "frmSchedule.frx":54DA
         PicArray70      =   "frmSchedule.frx":54F6
         PicArray71      =   "frmSchedule.frx":5512
         PicArray72      =   "frmSchedule.frx":552E
         PicArray73      =   "frmSchedule.frx":554A
         PicArray74      =   "frmSchedule.frx":5566
         PicArray75      =   "frmSchedule.frx":5582
         PicArray76      =   "frmSchedule.frx":559E
         PicArray77      =   "frmSchedule.frx":55BA
         PicArray78      =   "frmSchedule.frx":55D6
         PicArray79      =   "frmSchedule.frx":55F2
         PicArray80      =   "frmSchedule.frx":560E
         PicArray81      =   "frmSchedule.frx":562A
         PicArray82      =   "frmSchedule.frx":5646
         PicArray83      =   "frmSchedule.frx":5662
         PicArray84      =   "frmSchedule.frx":567E
         PicArray85      =   "frmSchedule.frx":569A
         PicArray86      =   "frmSchedule.frx":56B6
         PicArray87      =   "frmSchedule.frx":56D2
         PicArray88      =   "frmSchedule.frx":56EE
         PicArray89      =   "frmSchedule.frx":570A
         PicArray90      =   "frmSchedule.frx":5726
         PicArray91      =   "frmSchedule.frx":5742
         PicArray92      =   "frmSchedule.frx":575E
         PicArray93      =   "frmSchedule.frx":577A
         PicArray94      =   "frmSchedule.frx":5796
         PicArray95      =   "frmSchedule.frx":57B2
         PicArray96      =   "frmSchedule.frx":57CE
         PicArray97      =   "frmSchedule.frx":57EA
         PicArray98      =   "frmSchedule.frx":5806
         PicArray99      =   "frmSchedule.frx":5822
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   6300
         X2              =   6435
         Y1              =   2430
         Y2              =   2160
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   6300
         X2              =   6300
         Y1              =   1890
         Y2              =   2430
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   6300
         X2              =   6435
         Y1              =   1890
         Y2              =   2160
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00808000&
         BorderColor     =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   5985
         Top             =   2025
         Width           =   330
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   6300
         X2              =   6435
         Y1              =   4590
         Y2              =   4320
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   6300
         X2              =   6300
         Y1              =   4050
         Y2              =   4590
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   6300
         X2              =   6435
         Y1              =   4050
         Y2              =   4320
      End
      Begin VB.Shape Shape3 
         Height          =   285
         Index           =   0
         Left            =   5985
         Top             =   4185
         Width           =   330
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "[1/4단계]학습 대상 선정: 과목 및 기존 시험지를 선택합니다."
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   25
         Top             =   360
         Width           =   8970
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "% 이상"
         Height          =   375
         Index           =   2
         Left            =   8370
         TabIndex        =   22
         Top             =   4275
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "개 이상"
         Height          =   375
         Index           =   1
         Left            =   8370
         TabIndex        =   21
         Top             =   2880
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "개 이하"
         Height          =   375
         Index           =   0
         Left            =   8370
         TabIndex        =   20
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "선택과목명"
         Height          =   285
         Index           =   3
         Left            =   3825
         TabIndex        =   15
         Top             =   3150
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "과목명"
         Height          =   285
         Index           =   2
         Left            =   450
         TabIndex        =   14
         Top             =   3150
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "선택시험지명"
         Height          =   285
         Index           =   1
         Left            =   3825
         TabIndex        =   13
         Top             =   945
         Width           =   2265
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         X1              =   6435
         X2              =   6435
         Y1              =   990
         Y2              =   5505
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   6450
         X2              =   6450
         Y1              =   990
         Y2              =   5520
      End
      Begin VB.Label Label1 
         Caption         =   "시험지명"
         Height          =   285
         Index           =   0
         Left            =   450
         TabIndex        =   4
         Top             =   945
         Width           =   2265
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   465
         Index           =   0
         Left            =   450
         Top             =   225
         Width           =   9600
      End
   End
   Begin VB.Frame fra 
      Caption         =   "3/4"
      Height          =   6270
      Index           =   2
      Left            =   780
      TabIndex        =   2
      Top             =   240
      Width           =   10455
      Begin VB.Frame Frame5 
         Caption         =   "분포 방식"
         Height          =   1185
         Left            =   5280
         TabIndex        =   88
         Top             =   1050
         Width           =   4590
         Begin VB.OptionButton opt3_2 
            Caption         =   "첫날에 집중 분포"
            Height          =   315
            Index           =   1
            Left            =   540
            TabIndex        =   90
            Top             =   720
            Width           =   3195
         End
         Begin VB.OptionButton opt3_2 
            Caption         =   "골고루 분포"
            Height          =   315
            Index           =   0
            Left            =   540
            TabIndex        =   89
            Top             =   330
            Value           =   -1  'True
            Width           =   3195
         End
      End
      Begin VB.TextBox txt3_2 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2835
         TabIndex        =   83
         Top             =   3780
         Width           =   960
      End
      Begin VB.TextBox txt3_1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   50
         Text            =   "frmSchedule.frx":583E
         Top             =   5310
         Width           =   5415
      End
      Begin POCKETQUIZ.numText nTxt3_2 
         Height          =   330
         Left            =   6885
         TabIndex        =   48
         Top             =   3465
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "굴림"
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
      Begin POCKETQUIZ.numText nTxt3_1 
         Height          =   330
         Left            =   6885
         TabIndex        =   46
         Top             =   2880
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "굴림"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Object.CausesValidation=   0   'False
      End
      Begin MSComCtl2.MonthView mv3_2 
         Height          =   2370
         Left            =   3870
         TabIndex        =   45
         Top             =   2745
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   88342529
         CurrentDate     =   38207
      End
      Begin MSComCtl2.MonthView mv3_1 
         Height          =   2370
         Left            =   495
         TabIndex        =   44
         Top             =   2745
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   88342529
         CurrentDate     =   38207
      End
      Begin VB.Frame Frame1 
         Caption         =   "계획 형식"
         Height          =   1185
         Left            =   360
         TabIndex        =   38
         Top             =   1035
         Width           =   4770
         Begin VB.OptionButton opt3_1 
            Caption         =   "시작일, (X일) 마다, (Y문항) 씩"
            Height          =   240
            Index           =   2
            Left            =   450
            TabIndex        =   41
            Top             =   855
            Width           =   4050
         End
         Begin VB.OptionButton opt3_1 
            Caption         =   "시작일, 종료일, (Y문항) 씩"
            Height          =   240
            Index           =   1
            Left            =   450
            TabIndex        =   40
            Top             =   585
            Width           =   3570
         End
         Begin VB.OptionButton opt3_1 
            Caption         =   "시작일, 종료일, (X일) 마다"
            Height          =   240
            Index           =   0
            Left            =   450
            TabIndex        =   39
            Top             =   315
            Width           =   4200
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "기 간"
         Height          =   285
         Left            =   2790
         TabIndex        =   82
         Top             =   3510
         Width           =   1050
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000018&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000012&
         FillColor       =   &H80000018&
         Height          =   870
         Left            =   540
         Top             =   5220
         Width           =   5595
      End
      Begin VB.Label lb3_4 
         Caption         =   "문항씩"
         Height          =   285
         Left            =   8190
         TabIndex        =   49
         Top             =   3510
         Width           =   1185
      End
      Begin VB.Label lb3_3 
         Caption         =   "일마다"
         Height          =   285
         Left            =   8190
         TabIndex        =   47
         Top             =   2970
         Width           =   1185
      End
      Begin VB.Label lb3_2 
         Caption         =   "종료일"
         Height          =   285
         Left            =   4635
         TabIndex        =   43
         Top             =   2475
         Width           =   825
      End
      Begin VB.Label lb3_1 
         Caption         =   "시작일"
         Height          =   285
         Left            =   1170
         TabIndex        =   42
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "[3/4단계]일정 계획: 학습 일정을 수립합니다."
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   27
         Top             =   570
         Width           =   8970
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   465
         Index           =   2
         Left            =   270
         Top             =   450
         Width           =   9600
      End
   End
   Begin VB.Frame fra 
      Caption         =   "4/4"
      Height          =   6270
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   90
      Width           =   10455
      Begin VB.Frame Frame4 
         Caption         =   "시험지명"
         Height          =   1410
         Left            =   405
         TabIndex        =   72
         Top             =   810
         Width           =   9600
         Begin VB.CheckBox chk4_1 
            Caption         =   "시험지명 식별자"
            Height          =   420
            Left            =   450
            TabIndex        =   76
            Top             =   765
            Width           =   1770
         End
         Begin VB.ComboBox cbo4_1 
            Height          =   300
            ItemData        =   "frmSchedule.frx":5880
            Left            =   2250
            List            =   "frmSchedule.frx":60E5
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   855
            Width           =   915
         End
         Begin VB.TextBox txt4_3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3690
            MultiLine       =   -1  'True
            TabIndex        =   73
            Text            =   "frmSchedule.frx":6C15
            Top             =   765
            Width           =   3525
         End
         Begin POCKETQUIZ.numText nTxt4_1 
            Height          =   285
            Left            =   1305
            TabIndex        =   74
            Top             =   360
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FormatMask      =   ""
            MinVal          =   0
            MaxVal          =   0
            MaxLength       =   15
            FontSize        =   9
            FontName        =   "굴림"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   0
            EndProperty
            Alignment       =   0
         End
         Begin VB.Label Label7 
            Caption         =   "시험지명"
            Height          =   285
            Left            =   450
            TabIndex        =   78
            Top             =   405
            Width           =   870
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000018&
            Caption         =   "시험지명 예시"
            Height          =   285
            Left            =   3690
            TabIndex        =   77
            Top             =   450
            Width           =   1230
         End
         Begin VB.Shape txt4_1 
            BackColor       =   &H80000018&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000012&
            FillColor       =   &H80000018&
            Height          =   735
            Left            =   3285
            Top             =   360
            Width           =   4200
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "하위 시험지명 형식"
         Height          =   3255
         Left            =   405
         TabIndex        =   51
         Top             =   2295
         Width           =   9600
         Begin VB.TextBox txt4_7 
            Height          =   300
            Left            =   7560
            MaxLength       =   10
            TabIndex        =   81
            Top             =   1395
            Width           =   645
         End
         Begin VB.TextBox txt4_6 
            Height          =   300
            Left            =   4680
            MaxLength       =   10
            TabIndex        =   80
            Top             =   1395
            Width           =   645
         End
         Begin VB.TextBox txt4_5 
            Height          =   300
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   79
            Top             =   1395
            Width           =   645
         End
         Begin VB.TextBox txt4_4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   450
            MultiLine       =   -1  'True
            TabIndex        =   58
            Text            =   "frmSchedule.frx":6C25
            Top             =   2655
            Width           =   4200
         End
         Begin VB.ComboBox cbo4_4 
            Height          =   300
            ItemData        =   "frmSchedule.frx":6C3B
            Left            =   5670
            List            =   "frmSchedule.frx":6C5A
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1395
            Width           =   1500
         End
         Begin VB.ComboBox cbo4_3 
            Height          =   300
            ItemData        =   "frmSchedule.frx":6CAA
            Left            =   2790
            List            =   "frmSchedule.frx":6CC9
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1395
            Width           =   1455
         End
         Begin VB.ComboBox cbo4_2 
            Height          =   300
            ItemData        =   "frmSchedule.frx":6D19
            Left            =   405
            List            =   "frmSchedule.frx":6D26
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1395
            Width           =   1095
         End
         Begin VB.CheckBox chk4_4 
            Caption         =   "종료일 형식"
            Height          =   420
            Left            =   5670
            TabIndex        =   54
            Top             =   990
            Width           =   1455
         End
         Begin VB.CheckBox chk4_3 
            Caption         =   "시작일 형식"
            Height          =   420
            Left            =   2790
            TabIndex        =   53
            Top             =   990
            Width           =   1455
         End
         Begin VB.CheckBox chk4_2 
            Caption         =   "회차"
            Height          =   420
            Left            =   450
            TabIndex        =   52
            Top             =   990
            Width           =   1050
         End
         Begin VB.Shape Shape7 
            Height          =   1005
            Index           =   2
            Left            =   7380
            Top             =   945
            Width           =   1005
         End
         Begin VB.Shape Shape7 
            Height          =   1005
            Index           =   1
            Left            =   4500
            Top             =   945
            Width           =   1005
         End
         Begin VB.Label Label10 
            BackColor       =   &H80000018&
            Caption         =   "하위 시험지명 예시"
            Height          =   240
            Left            =   450
            TabIndex        =   59
            Top             =   2250
            Width           =   1815
         End
         Begin VB.Shape txt4_2 
            BackColor       =   &H80000018&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000012&
            FillColor       =   &H80000018&
            Height          =   1005
            Left            =   315
            Top             =   2115
            Width           =   4425
         End
         Begin VB.Shape Shape7 
            Height          =   1005
            Index           =   0
            Left            =   5490
            Top             =   945
            Width           =   1905
         End
         Begin VB.Shape Shape6 
            Height          =   1005
            Left            =   1620
            Top             =   945
            Width           =   1005
         End
         Begin VB.Shape Shape5 
            Height          =   1005
            Left            =   315
            Top             =   945
            Width           =   1320
         End
         Begin VB.Shape Shape4 
            Height          =   1005
            Left            =   2610
            Top             =   945
            Width           =   1905
         End
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "[4/4단계]시험지명 입력: 시험지명을 정합니다."
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   600
         TabIndex        =   28
         Top             =   375
         Width           =   8970
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   465
         Index           =   3
         Left            =   405
         Top             =   225
         Width           =   9600
      End
   End
   Begin VB.Frame fra 
      Caption         =   "2/4"
      Height          =   6270
      Index           =   1
      Left            =   495
      TabIndex        =   1
      Top             =   360
      Width           =   10455
      Begin CTLISTLibCtl.ctList lst2_1 
         Height          =   3975
         Left            =   630
         TabIndex        =   71
         Top             =   900
         Width           =   5370
         _Version        =   65536
         _ExtentX        =   9472
         _ExtentY        =   7011
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
         TitleBackImage  =   "frmSchedule.frx":6D45
         HeaderPicture   =   "frmSchedule.frx":6D61
         Picture         =   "frmSchedule.frx":6D7D
         CheckPicDown    =   "frmSchedule.frx":6D99
         CheckPicUp      =   "frmSchedule.frx":6DB5
         CheckPicDisabled=   "frmSchedule.frx":6DD1
         BackImage       =   "frmSchedule.frx":6DED
         ShowHeader      =   -1  'True
         SmoothScroll    =   -1  'True
         HeaderData      =   "frmSchedule.frx":6E09
         PicArray0       =   "frmSchedule.frx":6ED2
         PicArray1       =   "frmSchedule.frx":6EEE
         PicArray2       =   "frmSchedule.frx":6F0A
         PicArray3       =   "frmSchedule.frx":6F26
         PicArray4       =   "frmSchedule.frx":6F42
         PicArray5       =   "frmSchedule.frx":6F5E
         PicArray6       =   "frmSchedule.frx":6F7A
         PicArray7       =   "frmSchedule.frx":6F96
         PicArray8       =   "frmSchedule.frx":6FB2
         PicArray9       =   "frmSchedule.frx":6FCE
         PicArray10      =   "frmSchedule.frx":6FEA
         PicArray11      =   "frmSchedule.frx":7006
         PicArray12      =   "frmSchedule.frx":7022
         PicArray13      =   "frmSchedule.frx":703E
         PicArray14      =   "frmSchedule.frx":705A
         PicArray15      =   "frmSchedule.frx":7076
         PicArray16      =   "frmSchedule.frx":7092
         PicArray17      =   "frmSchedule.frx":70AE
         PicArray18      =   "frmSchedule.frx":70CA
         PicArray19      =   "frmSchedule.frx":70E6
         PicArray20      =   "frmSchedule.frx":7102
         PicArray21      =   "frmSchedule.frx":711E
         PicArray22      =   "frmSchedule.frx":713A
         PicArray23      =   "frmSchedule.frx":7156
         PicArray24      =   "frmSchedule.frx":7172
         PicArray25      =   "frmSchedule.frx":718E
         PicArray26      =   "frmSchedule.frx":71AA
         PicArray27      =   "frmSchedule.frx":71C6
         PicArray28      =   "frmSchedule.frx":71E2
         PicArray29      =   "frmSchedule.frx":71FE
         PicArray30      =   "frmSchedule.frx":721A
         PicArray31      =   "frmSchedule.frx":7236
         PicArray32      =   "frmSchedule.frx":7252
         PicArray33      =   "frmSchedule.frx":726E
         PicArray34      =   "frmSchedule.frx":728A
         PicArray35      =   "frmSchedule.frx":72A6
         PicArray36      =   "frmSchedule.frx":72C2
         PicArray37      =   "frmSchedule.frx":72DE
         PicArray38      =   "frmSchedule.frx":72FA
         PicArray39      =   "frmSchedule.frx":7316
         PicArray40      =   "frmSchedule.frx":7332
         PicArray41      =   "frmSchedule.frx":734E
         PicArray42      =   "frmSchedule.frx":736A
         PicArray43      =   "frmSchedule.frx":7386
         PicArray44      =   "frmSchedule.frx":73A2
         PicArray45      =   "frmSchedule.frx":73BE
         PicArray46      =   "frmSchedule.frx":73DA
         PicArray47      =   "frmSchedule.frx":73F6
         PicArray48      =   "frmSchedule.frx":7412
         PicArray49      =   "frmSchedule.frx":742E
         PicArray50      =   "frmSchedule.frx":744A
         PicArray51      =   "frmSchedule.frx":7466
         PicArray52      =   "frmSchedule.frx":7482
         PicArray53      =   "frmSchedule.frx":749E
         PicArray54      =   "frmSchedule.frx":74BA
         PicArray55      =   "frmSchedule.frx":74D6
         PicArray56      =   "frmSchedule.frx":74F2
         PicArray57      =   "frmSchedule.frx":750E
         PicArray58      =   "frmSchedule.frx":752A
         PicArray59      =   "frmSchedule.frx":7546
         PicArray60      =   "frmSchedule.frx":7562
         PicArray61      =   "frmSchedule.frx":757E
         PicArray62      =   "frmSchedule.frx":759A
         PicArray63      =   "frmSchedule.frx":75B6
         PicArray64      =   "frmSchedule.frx":75D2
         PicArray65      =   "frmSchedule.frx":75EE
         PicArray66      =   "frmSchedule.frx":760A
         PicArray67      =   "frmSchedule.frx":7626
         PicArray68      =   "frmSchedule.frx":7642
         PicArray69      =   "frmSchedule.frx":765E
         PicArray70      =   "frmSchedule.frx":767A
         PicArray71      =   "frmSchedule.frx":7696
         PicArray72      =   "frmSchedule.frx":76B2
         PicArray73      =   "frmSchedule.frx":76CE
         PicArray74      =   "frmSchedule.frx":76EA
         PicArray75      =   "frmSchedule.frx":7706
         PicArray76      =   "frmSchedule.frx":7722
         PicArray77      =   "frmSchedule.frx":773E
         PicArray78      =   "frmSchedule.frx":775A
         PicArray79      =   "frmSchedule.frx":7776
         PicArray80      =   "frmSchedule.frx":7792
         PicArray81      =   "frmSchedule.frx":77AE
         PicArray82      =   "frmSchedule.frx":77CA
         PicArray83      =   "frmSchedule.frx":77E6
         PicArray84      =   "frmSchedule.frx":7802
         PicArray85      =   "frmSchedule.frx":781E
         PicArray86      =   "frmSchedule.frx":783A
         PicArray87      =   "frmSchedule.frx":7856
         PicArray88      =   "frmSchedule.frx":7872
         PicArray89      =   "frmSchedule.frx":788E
         PicArray90      =   "frmSchedule.frx":78AA
         PicArray91      =   "frmSchedule.frx":78C6
         PicArray92      =   "frmSchedule.frx":78E2
         PicArray93      =   "frmSchedule.frx":78FE
         PicArray94      =   "frmSchedule.frx":791A
         PicArray95      =   "frmSchedule.frx":7936
         PicArray96      =   "frmSchedule.frx":7952
         PicArray97      =   "frmSchedule.frx":796E
         PicArray98      =   "frmSchedule.frx":798A
         PicArray99      =   "frmSchedule.frx":79A6
      End
      Begin VB.Frame Frame3 
         Caption         =   "문제 순서"
         Height          =   1950
         Left            =   6210
         TabIndex        =   66
         Top             =   2565
         Width           =   2130
         Begin VB.OptionButton opt2_1 
            Caption         =   "회차內"
            Height          =   195
            Index           =   2
            Left            =   495
            TabIndex        =   70
            Top             =   1305
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.OptionButton opt2_1 
            Caption         =   "과목內"
            Height          =   195
            Index           =   1
            Left            =   495
            TabIndex        =   69
            Top             =   1012
            Width           =   1320
         End
         Begin VB.OptionButton opt2_1 
            Caption         =   "전    체"
            Height          =   195
            Index           =   0
            Left            =   495
            TabIndex        =   68
            Top             =   765
            Width           =   1320
         End
         Begin VB.CheckBox chk2_1 
            Caption         =   "문제 섞기"
            Height          =   240
            Left            =   270
            TabIndex        =   67
            Top             =   405
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "아래로"
         Height          =   375
         Left            =   6345
         TabIndex        =   65
         Top             =   1710
         Width           =   1230
      End
      Begin VB.CommandButton cmdUP 
         Caption         =   "위로"
         Height          =   375
         Left            =   6345
         TabIndex        =   64
         Top             =   1305
         Width           =   1230
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "[2/4단계]과목선정: 과목을 선정합니다. 제외시킬 수 있습니다."
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   810
         TabIndex        =   26
         Top             =   450
         Width           =   8970
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   465
         Index           =   1
         Left            =   540
         Top             =   315
         Width           =   9600
      End
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "이전"
      Height          =   420
      Left            =   7335
      TabIndex        =   30
      Top             =   6660
      Width           =   1050
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "다음"
      Height          =   420
      Left            =   8460
      TabIndex        =   31
      Top             =   6615
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   420
      Left            =   9675
      TabIndex        =   32
      Top             =   6615
      Width           =   1050
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   85
      Top             =   6885
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public parent As frmMain
Dim WithEvents lclsSchedule As clsSchedule
Attribute lclsSchedule.VB_VarHelpID = -1
'Attribute lclsSchedule.VB_VarHelpID = -1
Dim savedSql As String


Private Sub changeButtonStatus()

Dim idx As Integer
idx = getSelectedFrame()

cmdPre.Visible = True

Select Case idx
Case 0
    cmdNext.Caption = "다음"
    cmdPre.Visible = False
    cmdPre.Enabled = False
Case 1
    cmdNext.Caption = "다음"
    cmdPre.Enabled = True
Case 2
    cmdNext.Caption = "다음"
    cmdPre.Enabled = True
Case 3
    cmdNext.Caption = "완료"
    cmdPre.Enabled = True
Case 4
    Debug.Assert False
End Select
cmdNext.ZOrder 0
cmdPre.ZOrder 0
cmdClose.ZOrder 0

End Sub

Private Function getSelectedFrame() As Integer
Dim idx As Integer
Dim lFra As Frame
For Each lFra In fra
    If lFra.Visible Then
        idx = lFra.Index
        Exit For
    End If
Next
getSelectedFrame = idx
End Function

Private Sub Check9_Click()

End Sub

Private Sub cbo4_1_Click()
    lclsSchedule.sInitial = cbo4_1.Text
    txt4_3.Text = lclsSchedule.getPocketNm
End Sub

Private Sub cbo4_2_Click()
lclsSchedule.sSubNm1Format = cbo4_2.Text
txt4_4.Text = lclsSchedule.getSubPocketNm()
End Sub

Private Sub cbo4_3_Click()
lclsSchedule.sSubNm2Format = cbo4_3.Text
txt4_4.Text = lclsSchedule.getSubPocketNm()

End Sub


Private Sub cbo4_4_Click()
lclsSchedule.sSubNm3Format = cbo4_4.Text
txt4_4.Text = lclsSchedule.getSubPocketNm

End Sub

Private Sub chk1_1_Click()
lclsSchedule.bCorrect = CBool(chk1_1)

    If chk1_1.Value Then
        If CBool(chk1_6.Value) Then
            chk1_1.Value = vbGrayed
        End If
    End If
    

End Sub

Private Sub chk1_2_Click()
    lclsSchedule.bIncorrect = CBool(chk1_2.Value)
    
    If chk1_2.Value Then
        If CBool(chk1_6.Value) Then
            chk1_2.Value = vbGrayed
        End If
    End If
    
End Sub

Private Sub chk1_3_Click()
    lclsSchedule.bIncorrectRate = CBool(chk1_3.Value)
    
    If chk1_3.Value Then
        If CBool(chk1_6.Value) Then
            chk1_3.Value = vbGrayed
        End If
    End If

End Sub

Private Sub chk1_4_Click()
    lclsSchedule.bReserveChk = CBool(chk1_4.Value)
End Sub

Private Sub chk1_5_Click(Index As Integer)
    Static oneTime As Boolean
    If oneTime = True Then Exit Sub
    oneTime = True
    If Index = 0 Then
        lclsSchedule.bReserveFromChk = CBool(chk1_5(Index).Value)
    Else
        lclsSchedule.bReserveToChk = CBool(chk1_5(Index).Value)
    End If
    
    If chk1_5(Index).Value Then
        If Not CBool(chk1_4.Value) Then
            chk1_5(Index).Value = vbGrayed
        End If
    End If
    oneTime = False
End Sub

Private Sub chk1_6_Click()
If chk1_6.Value = vbChecked Then
    If chk1_1.Value <> vbUnchecked Then
        chk1_1.Value = vbGrayed
    End If
    
    If chk1_2.Value <> vbUnchecked Then
        chk1_2.Value = vbGrayed
    End If
    
    If chk1_3.Value <> vbUnchecked Then
        chk1_3.Value = vbGrayed
    End If
Else
    If chk1_1.Value <> vbUnchecked Then
        chk1_1.Value = vbChecked
    End If
    
    If chk1_2.Value <> vbUnchecked Then
        chk1_2.Value = vbChecked
    End If
    
    If chk1_3.Value <> vbUnchecked Then
        chk1_3.Value = vbChecked
    End If
    

End If
End Sub

Private Sub chk2_1_Click()
lclsSchedule.bQuizSwap = CBool(chk2_1.Value)

lclsSchedule.bQuizSwap_all = opt2_1(0).Value
lclsSchedule.bQuizSwap_subj = opt2_1(1).Value
lclsSchedule.bQuizSwap_chasu = opt2_1(2).Value

If lclsSchedule.bQuizSwap_all = False And lclsSchedule.bQuizSwap_subj = False And lclsSchedule.bQuizSwap_chasu = False Then
    opt2_1(0).Value = True
End If

End Sub

Private Sub chk4_1_Click()
    lclsSchedule.sInitial = cbo4_1.Text
    lclsSchedule.bInitial = CBool(chk4_1.Value)
    txt4_3.Text = lclsSchedule.getPocketNm
End Sub

Private Sub chk4_2_Click()
lclsSchedule.bSubNm1Format = CBool(chk4_2.Value)
txt4_4.Text = lclsSchedule.getSubPocketNm
End Sub

Private Sub chk4_3_Click()
lclsSchedule.bSubNm2Format = CBool(chk4_3.Value)
txt4_4.Text = lclsSchedule.getSubPocketNm
End Sub

Private Sub chk4_4_Click()
lclsSchedule.bSubNm3Format = CBool(chk4_4.Value)
txt4_4.Text = lclsSchedule.getSubPocketNm

End Sub

Private Sub cmd1L2R_Click()
lst1_2_DropList lst1_2.ListCount, 0
End Sub

Private Sub cmd1L2Ra_Click()
Dim i As Integer
For i = 0 To lst1_1.ListCount - 1
    lst1_1.ListSelect(i) = True
Next
cmd1L2R_Click
End Sub

Private Sub cmd1R2L_Click()
lst1_1_DropList lst1_1.ListCount, 0
End Sub

Private Sub cmd1R2La_Click()
Dim i As Integer
For i = 0 To lst1_2.ListCount - 1
    lst1_2.ListSelect(i) = True
Next
cmd1R2L_Click

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

Private Sub cmdClose_Click()
If parent.imgSplitter.Left < 1000 Then
    parent.SizeControls 1000
End If
Unload Me
End Sub

Private Sub cmdDown_Click()
Dim nIdx As Integer

Dim srcStr As String
Dim srcChk As Integer

Dim tarStr As String
Dim tarChk As Integer

nIdx = lst2_1.Selected

If nIdx < 0 Then
    MsgBox "선택된 항목이 없습니다", vbExclamation
    Exit Sub
End If

If nIdx = lst2_1.ListCount - 1 Then Exit Sub

srcStr = lst2_1.ListText(nIdx)
srcChk = lst2_1.ListColumnCheck(nIdx, 1)

tarStr = lst2_1.ListText(nIdx + 1)
tarChk = lst2_1.ListColumnCheck(nIdx + 1, 1)

lst2_1.ListText(nIdx) = tarStr
lst2_1.ListColumnCheck(nIdx, 1) = tarChk

lst2_1.ListText(nIdx + 1) = srcStr
lst2_1.ListColumnCheck(nIdx + 1, 1) = srcChk

lst2_1.ListSelect(nIdx + 1) = True
End Sub

Private Sub cmdNext_Click()
Dim idx As Integer

idx = getSelectedFrame()
'Call MsgBox(CStr(idx + 1) & "단계", vbExclamation)
If Not fraPorcess(idx) Then Exit Sub
cmdNext.Enabled = False
If idx = 0 Then

    Dim lRs As ADODB.Recordset
   
    sSql = "select count(*) cnt from tm01 where userid='" + gUserid + "' and ymd = date_format(current_date,'%Y%m%d')"
    Set lRs = Fn_SQLExec(sSql).rs

    If 4 < lRs(0) Then
        MsgBox "하루에 계획을 최대 5개 까지만 만들 수 있습니다.", vbOKOnly + vbExclamation, Me.Caption
    End If

    fra(0).Visible = False
    fra(1).Visible = True
    changeButtonStatus
ElseIf idx = 1 Then
    fra(1).Visible = False
    fra(2).Visible = True
    changeButtonStatus
ElseIf idx = 2 Then
    fra(2).Visible = False
    fra(3).Visible = True
    changeButtonStatus
ElseIf idx = 3 Then
    makeprocess
End If

cmdNext.Enabled = True
End Sub


Private Sub cmdPre_Click()
Dim idx As Integer
idx = getSelectedFrame()

If idx = 0 Then

ElseIf idx = 1 Then
    fra(0).Visible = True
    fra(1).Visible = False
ElseIf idx = 2 Then
    fra(1).Visible = True
    fra(2).Visible = False
ElseIf idx = 3 Then
    fra(2).Visible = True
    fra(3).Visible = False
ElseIf idx = 4 Then
    fra(3).Visible = True
    fra(4).Visible = False
End If

changeButtonStatus

End Sub

Private Sub cmdUP_Click()

Debug.Print lst2_1.Selected

Dim nIdx As Integer

Dim srcStr As String
Dim srcChk As Integer

Dim tarStr As String
Dim tarChk As Integer

nIdx = lst2_1.Selected

If nIdx < 0 Then
    MsgBox "선택된 항목이 없습니다", vbExclamation
    Exit Sub
End If

If nIdx = 0 Then Exit Sub

srcStr = lst2_1.ListText(nIdx)
srcChk = lst2_1.ListColumnCheck(nIdx, 1)

tarStr = lst2_1.ListText(nIdx - 1)
tarChk = lst2_1.ListColumnCheck(nIdx - 1, 1)

lst2_1.ListText(nIdx) = tarStr
lst2_1.ListColumnCheck(nIdx, 1) = tarChk

lst2_1.ListText(nIdx - 1) = srcStr
lst2_1.ListColumnCheck(nIdx - 1, 1) = srcChk

lst2_1.ListSelect(nIdx - 1) = True

End Sub

'Private Sub dtp1_CloseUp()
'nTxt1_2.Text = date2Str(dtp1.Value)
'End Sub
'
'Private Sub dtp1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'dtp1.Value = str2Date(nTxt1_2.Text)
'End Sub
'Private Sub dtp2_CloseUp()
'nTxt1_3.Text = date2Str(dtp2.Value)
'End Sub
'
'Private Sub dtp2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'dtp2.Value = str2Date(nTxt1_3.Text)
'End Sub


Private Sub Form_Load()
    Dim lFra As Frame
    
    Set lclsSchedule = New clsSchedule
    
    '##1
    Dim conbef As Integer
    If Con.State = 1 Then
        conbef = 1
    Else
        Con_Open
        conbef = 0
    End If
    
    setInitStatus
    For Each lFra In fra
        lFra.Caption = ""
        lFra.Visible = False
        lFra.Left = fra(0).Left
        lFra.Top = fra(0).Top
        lFra.Width = fra(0).Width
        lFra.Height = fra(0).Height
    Next
    fra(0).Visible = True
    
    changeButtonStatus
    
    '##2
    If conbef = 1 Then
    Else
        Con_Close
    End If
        
End Sub
Sub setInitStatus()

Dim lRs As ADODB.Recordset

lst1_1.ClearList
lst1_2.ClearList
lst1_3.ClearList
lst1_4.ClearList

sSql = "SELECT TP02.code, TP02.pocketnm, TP02.chasu"
sSql = sSql & " FROM TP02 where userid='" & gUserid & "' and pcode=0 and hidden=0 order by code"

Set lRs = Fn_SQLExec(sSql).rs
If lRs Is Nothing Then Exit Sub
Do Until lRs.EOF
    Call lst1_1.AddItem(lRs(1) & Chr(10) & lRs(0))
    lRs.MoveNext
Loop
lRs.Close

'회원가입시 디폴트 과목에 넣을 것 20060422
'insert into ts02(subj,userid,startymd,endymd)
'select a.subj,b.userid,'20000101','21001231' from ts01 a ,tu01 b where
'a.subj in ('영단어1','영속담1','영숙어(중)','영숙어1','중1단어','한자','한자01','토익Voca','운전면허표지안전판','운전01','운전02','운전03','운전04','운전05','운전06','운전07','운전08','운전09','운전10','운전11','운전12','운전A','운전B','운전C','운전D','일본어800') and b.userid='영어';

'sSql = "SELECT subjnm,subj from ts01"
sSql = "SELECT a.subjnm,a.subj from ts01 a, ts02 b where a.subj=b.subj and "
sSql = sSql & " b.userid='" & gUserid & "' and b.startymd<=date_format(current_date,'%Y%m%d') and b.endymd>=date_format(current_date,'%Y%m%d')"

Set lRs = Fn_SQLExec(sSql).rs

Do Until lRs.EOF
    Call lst1_3.AddItem(lRs(0) & Chr(10) & lRs(1))
    lRs.MoveNext
Loop
lRs.Close

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Con_Close
End Sub

Private Sub Form_Resize()
Dim lLeft As Single
Dim lTop As Single
Dim lFra As Frame

lLeft = (Me.Width - fra(0).Width) / 2
lTop = (Me.Height - fra(0).Height) / 2

For Each lFra In fra
   lFra.Left = lLeft
   lFra.Top = lTop
Next

cmdPre.Left = fra(0).Left + fra(0).Width - cmdPre.Width - cmdPre.Width - cmdPre.Width - cmdPre.Width
cmdNext.Left = fra(0).Left + fra(0).Width - cmdPre.Width - cmdPre.Width - cmdPre.Width * 2 / 3
cmdClose.Left = fra(0).Left + fra(0).Width - cmdPre.Width - cmdPre.Width / 3

cmdPre.Top = fra(0).Top + fra(0).Height - cmdPre.Height * 1.5
cmdNext.Top = cmdPre.Top
cmdClose.Top = cmdPre.Top

End Sub

Private Sub lclsSchedule_ReserveChkChanged(ByVal bChk As Boolean)
If bChk Then
    If chk1_5(0).Value <> vbUnchecked Then
        chk1_5(0).Value = vbChecked
    End If
    If chk1_5(1).Value <> vbUnchecked Then
        chk1_5(1).Value = vbChecked
    End If

Else

    If chk1_5(0).Value <> vbUnchecked Then
        chk1_5(0).Value = vbGrayed
    End If
    If chk1_5(1).Value <> vbUnchecked Then
        chk1_5(1).Value = vbGrayed
    End If

End If
End Sub

Private Sub lst1_1_DblClick()
cmd1L2R_Click
End Sub

Private Sub lst1_1_DragDrop(Source As Control, x As Single, y As Single)
    If Source.Name = "lst1_2" Then
        lst1_1.DragDrop (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub lst1_1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If Source.Name = "lst1_2" Then
        lst1_1.DragOver (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY), State
    End If

End Sub

Private Sub lst1_1_DropList(ByVal nIndex As Long, ByVal nColumn As Integer)
Dim strText As String
Dim i As Integer
    
    
    For i = lst1_2.ListCount - 1 To 0 Step -1
        If lst1_2.ListSelect(i) Then
            strText = lst1_2.ListText(i)
            lst1_1.InsertItem strText, nIndex
            lst1_2.removeItem (i)
        End If
    Next
    


End Sub

Private Sub lst1_1_StartDragOut()
lst1_1.Drag vbBeginDrag
End Sub

Private Sub lst1_2_DblClick()
cmd1R2L_Click
End Sub

Private Sub lst1_2_DragDrop(Source As Control, x As Single, y As Single)
    If Source.Name = "lst1_1" Then
        lst1_2.DragDrop (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub lst1_2_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If Source.Name = "lst1_1" Then
        lst1_2.DragOver (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY), State
    End If
End Sub

Private Sub lst1_2_DropList(ByVal nIndex As Long, ByVal nColumn As Integer)
Dim strText As String
Dim i As Integer
    
    
    For i = lst1_1.ListCount - 1 To 0 Step -1
        If lst1_1.ListSelect(i) Then
            strText = lst1_1.ListText(i)
            lst1_2.InsertItem strText, nIndex
            lst1_1.removeItem (i)
        End If
    Next
    

End Sub

Private Sub lst1_2_StartDragOut()
lst1_2.Drag vbBeginDrag
End Sub

Private Sub lst1_3_DblClick()
cmd2L2R_Click
End Sub

Private Sub lst1_3_DragDrop(Source As Control, x As Single, y As Single)
    If Source.Name = "lst1_4" Then
        lst1_3.DragDrop (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub lst1_3_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If Source.Name = "lst1_4" Then
        lst1_3.DragOver (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY), State
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

Private Sub lst1_4_DragDrop(Source As Control, x As Single, y As Single)
    If Source.Name = "lst1_3" Then
        lst1_4.DragDrop (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub lst1_4_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If Source.Name = "lst1_3" Then
        lst1_4.DragOver (x / Screen.TwipsPerPixelX), (y / Screen.TwipsPerPixelY), State
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

Private Sub lst1_4_StartDragOut()
lst1_4.Drag vbBeginDrag
End Sub



Private Sub mv3_1_SelChange(ByVal startDate As Date, ByVal endDate As Date, Cancel As Boolean)
On Error Resume Next
    mv3_1.Value = str2Date(Format(mv3_1.Value, "yyyyMMdd"))
    lclsSchedule.sReserveFrom = Format(mv3_1.Value, "yyyymmdd")
    lclsSchedule.nDayTotal = CLng(mv3_2.Value - mv3_1.Value) + 1
    txt3_2.Text = lclsSchedule.nDayTotal 'CLng(mv3_2.Value - mv3_1.Value) + 1
    cmdNext.SetFocus
    
End Sub


Private Sub mv3_2_SelChange(ByVal startDate As Date, ByVal endDate As Date, Cancel As Boolean)
On Error Resume Next
mv3_2.Value = str2Date(Format(mv3_2.Value, "yyyyMMdd"))
    lclsSchedule.nDayTotal = CLng(mv3_2.Value - mv3_1.Value) + 1
    txt3_2.Text = lclsSchedule.nDayTotal 'CLng(mv3_2.Value - mv3_1.Value) + 1
    lclsSchedule.sReserveTo = Format(mv3_2.Value, "yyyymmdd")
    cmdNext.SetFocus
End Sub

Private Sub nTxt1_1_Change(Index As Integer)
    
    Select Case Index
    Case 0:
        lclsSchedule.CorrectCnt = CLng(IIf(nTxt1_1(Index).Text = "", 0, nTxt1_1(Index).Text))
    Case 1:
        lclsSchedule.IncorrectCnt = CLng(IIf(nTxt1_1(Index).Text = "", 0, nTxt1_1(Index).Text))
    Case 2:
        lclsSchedule.IncorrectRate = CLng(IIf(nTxt1_1(Index).Text = "", 0, nTxt1_1(Index).Text))
        End Select
      
End Sub

Private Function fraPorcess(idx As Integer) As Boolean

 On Error GoTo errorTrap

    Dim ssql2b As String
    Dim ssql3c As String
    
    Dim lRs As ADODB.Recordset
    Dim i As Integer
    Dim cnt As Long
    Dim str As String
    Dim strFirstSubj As String
    Dim subjcnt As Long
    
    If Not fraValidation(idx) Then Exit Function
    
    Select Case idx
    '[1/4]단계에서 진행 설정
    Case 0
    
        '====================================================================
        '최적화 프로세스로 선택과목 목록만 어카운트 추가 한다.
        '단! 등록일을 2틀전으로 한다. 왜냐하면 오늘인경우에는 틀린카운트가 반
        '영되지 않기 때문이다.20060514
        '====================================================================
        sSql = "INSERT INTO TU02(subj, seq, userid, o, x, chk, update_ymd, reserve_ymd, gangyek) "
        sSql = sSql & vbCrLf & " (select a.subj,a.seq,'" & gUserid & "',0,0,0,'" & date2Str(DateAdd("d", -2, Now)) & "','99999999',0 "
        sSql = sSql & vbCrLf & "FROM vq01 a , ts02 d where a.subj=d.subj and d.userid='" & gUserid & "' and d.subj in (" & itemSeries2(lst1_4, Chr(10)) & ")"
        sSql = sSql & vbCrLf & "and not exists (select a.subj,a.seq from tu02  b,ts02 c where"
        sSql = sSql & vbCrLf & "b.userid = '" & gUserid & "' and a.subj=b.subj and a.seq=b.seq and b.userid=c.userid"
        sSql = sSql & vbCrLf & "and a.subj=c.subj and a.subj=b.subj))"
    
        cnt = Fn_SQLExec(sSql).nrow
        Debug.Print cnt & "건의 어카운트가 추가되었습니다."
    
        '_/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/
        '
        '왼쪽은 union으로 묶고 오른쪽조건은 and로 묶어야 하는데 그게 안돼어 있는것같다.
        '그래서 tp02 에 있는 결과만 뿌리도록 해야 옳다. 20050424
        '
        '_/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/ _/
    
        If chk1_6.Value = vbChecked Then '안푼문제
        
            '//아이디와 시험지코드 맞은수 틀린수로 해당문제지마스타(tu02)의 문제 조회
            sSql = "select c.subj,c.seq  from tp02 a, tp03 b, tu02 c where 1=1 and 3=3 and a.userid='" & gUserid & "'  and a.code in (" & Replace(Replace(itemSeries2(lst1_2, Chr(10)), "''", "-1"), "'", "") & ")"
            sSql = sSql + vbCrLf + "and a.pocketnm=b.pocketnm and a.chasu=b.chasu and a.userid=b.userid"
            sSql = sSql + vbCrLf + "and b.subj=c.subj and b.seq=c.seq and b.userid=c.userid"
            sSql = sSql + vbCrLf + "and c.o=0"
            sSql = sSql + vbCrLf + "and c.x=0"
        
        Else
            '//아이디와 시험지코드 맞은수 틀린수로 해당문제지마스타(tu02)의 문제 조회
            
            If chk1_3.Value <> vbChecked Then
            
                sSql = "select c.subj,c.seq  from tp02 a, tp03 b, tu02 c where 1=1 and 3=3 and a.userid='" & gUserid & "'  and a.code in (" & Replace(Replace(itemSeries2(lst1_2, Chr(10)), "''", "-1"), "'", "") & ")"
                sSql = sSql + vbCrLf + "and a.pocketnm=b.pocketnm and a.chasu=b.chasu and a.userid=b.userid"
                sSql = sSql + vbCrLf + "and b.subj=c.subj and b.seq=c.seq and b.userid=c.userid"
            
                If chk1_1.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and c.o<=" & nTxt1_1(0).Text & " and c.o+c.x>0" '푼 문제로 국한시킨다 '20051230
                End If
                If chk1_2.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and c.x>=" & nTxt1_1(1).Text & " and c.o+c.x>0" '푼 문제로 국한시킨다 '20051230
                End If
            End If
            If chk1_3.Value = vbChecked Then
'                sSql = sSql + vbCrLf + "Union"
                '//위와 마찬가지로 하여 틀린 비율에 해당하는 문제지마스타의 문제 조회
                sSql = "select subj,seq From ( select (c.x+0.00001)/(c.x+c.o+0.00001)*100 as rt,c.*  from tp02 a, tp03 b, tu02 c where 1=1 and 3=3 and a.userid='" & gUserid & "'  and a.code in (" & Replace(Replace(itemSeries2(lst1_2, Chr(10)), "''", "-1"), "'", "") & ")"
                sSql = sSql + vbCrLf + "and a.pocketnm=b.pocketnm and a.chasu=b.chasu and a.userid=b.userid"
                sSql = sSql + vbCrLf + "and b.subj=c.subj and b.seq=c.seq and b.userid=c.userid"
                sSql = sSql + vbCrLf + "and c.o+c.x>0"
                
                If chk1_1.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and c.o<=" & nTxt1_1(0).Text & " and c.o+c.x>0" '푼 문제로 국한시킨다 '20051230
                End If
                If chk1_2.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and c.x>=" & nTxt1_1(1).Text & " and c.o+c.x>0" '푼 문제로 국한시킨다 '20051230
                End If
                
                sSql = sSql + vbCrLf + ") d"
                sSql = sSql + vbCrLf + "Where d.rt > " & nTxt1_1(2).Text & " "
                
    
            End If
        End If

        sSql = sSql + vbCrLf + "Union"
        
        If chk1_6.Value = vbChecked Then

            '//사용자아이디와 과목명으로 문제지 마스타의 문제 조회
            sSql = sSql + vbCrLf + "select b.subj,b.seq from ts01 a , tu02 b"
            sSql = sSql + vbCrLf + "where 1=1 and 2=2 and a.subj in (" & itemSeries2(lst1_4, Chr(10)) & ")"
            sSql = sSql + vbCrLf + "and a.subj=b.subj"
            sSql = sSql + vbCrLf + "and b.userid='" & gUserid & "'"
            sSql = sSql + vbCrLf + "and b.o=0"
            sSql = sSql + vbCrLf + "and b.x=0"
                
        Else
        
            If chk1_3.Value <> vbChecked Then
                '//사용자아이디와 과목명으로 문제지 마스타의 문제 조회
                sSql = sSql + vbCrLf + "select b.subj,b.seq from ts01 a , tu02 b"
                sSql = sSql + vbCrLf + "where 1=1 and 2=2 and a.subj in (" & itemSeries2(lst1_4, Chr(10)) & ")"
                sSql = sSql + vbCrLf + "and a.subj=b.subj"
                sSql = sSql + vbCrLf + "and b.userid='" & gUserid & "'"
                If chk1_1.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and b.o<=" & nTxt1_1(0).Text & " and b.o+b.x>1"
                End If
                If chk1_2.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and b.x>=" & nTxt1_1(1).Text & " and b.o+b.x>1"
                End If
                
            End If
            
            If chk1_3.Value = vbChecked Then
'                sSql = sSql + vbCrLf + "Union"
                '//위와 마찬가지로 하여 틀린비율에 해당하는 문제지 마스타 조회
                sSql = sSql + vbCrLf + "select subj,seq from (select (b.x+0.00001)/(b.o+b.x+0.00001)*100 as rt , b.*  from ts01 a , tu02 b"
                sSql = sSql + vbCrLf + "where 1=1 and 2=2 and a.subj in (" & itemSeries2(lst1_4, Chr(10)) & ")"
                sSql = sSql + vbCrLf + "and a.subj=b.subj"
                sSql = sSql + vbCrLf + "and b.userid='" & gUserid & "'"
                sSql = sSql + vbCrLf + "and b.o+b.x>0"
                If chk1_1.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and b.o<=" & nTxt1_1(0).Text & " and b.o+b.x>1"
                End If
                If chk1_2.Value = vbChecked Then
                    sSql = sSql + vbCrLf + "and b.x>=" & nTxt1_1(1).Text & " and b.o+b.x>1"
                End If
                sSql = sSql + vbCrLf + ") cc where cc.rt>=" & nTxt1_1(2).Text & ""
            '
            End If
                    
        End If
        
        
'------------------------------------깊이과목으로 조회한다.
        Dim ssql_selectes As String
        ssql_selectes = selectSeriesTP01(itemSeries2(lst1_4, Chr(10)))
        
        If ssql_selectes <> "" Then
            sSql = sSql + vbCrLf + "Union"
            
            If chk1_6.Value = vbChecked Then '안푼문제 선택
    
                '//사용자아이디와 깊이있는과목명으로 문제지 마스타의 문제 조회
                sSql = sSql + vbCrLf + "select b.subj,b.seq from ($$$) a , tu02 b"
                sSql = sSql + vbCrLf + "where 1=1 and 2=2 and a.subj=b.subj"
                sSql = sSql + vbCrLf + "and a.seq=b.seq"
                sSql = sSql + vbCrLf + "and b.userid='" & gUserid & "'"
                sSql = sSql + vbCrLf + "and b.o=0"
                sSql = sSql + vbCrLf + "and b.x=0"
            Else
                If chk1_3.Value <> vbChecked Then
                    '//사용자아이디와 깊이있는과목명으로 문제지 마스타의 문제 조회
                    sSql = sSql + vbCrLf + "select b.subj,b.seq from ($$$) a , tu02 b"
                    sSql = sSql + vbCrLf + "where 1=1 and 2=2 and a.subj=b.subj and a.seq=b.seq"
                    sSql = sSql + vbCrLf + "and b.userid='" & gUserid & "'"
                    If chk1_1.Value = vbChecked Then
                        sSql = sSql + vbCrLf + "and b.o<=" & nTxt1_1(0).Text & " and b.o+b.x>0"
                    End If
                    If chk1_2.Value = vbChecked Then
                        sSql = sSql + vbCrLf + "and b.x>=" & nTxt1_1(1).Text & " and b.o+b.x>0"
                    End If
                    
                End If
                
                If chk1_3.Value = vbChecked Then
    '                sSql = sSql + vbCrLf + "Union"
                    '//위와 마찬가지로 하여 틀린비율에 해당하는 문제지 마스타 조회
                    sSql = sSql + vbCrLf + "select subj,seq from (select (b.x+0.00001)/(b.o+b.x+0.00001)*100 as rt , b.*  from ($$$) a , tu02 b"
                    sSql = sSql + vbCrLf + "where 1=1 and 2=2 and a.subj=b.subj and a.seq=b.seq"
                    sSql = sSql + vbCrLf + "and b.userid='" & gUserid & "'"
                    sSql = sSql + vbCrLf + "and b.o+b.x>0"
                    
                    If chk1_1.Value = vbChecked Then
                        sSql = sSql + vbCrLf + "and b.o<=" & nTxt1_1(0).Text & " and b.o+b.x>0"
                    End If
                    If chk1_2.Value = vbChecked Then
                        sSql = sSql + vbCrLf + "and b.x>=" & nTxt1_1(1).Text & " and b.o+b.x>0"
                    End If
                    
                    sSql = sSql + vbCrLf + ") cc where cc.rt>=" & nTxt1_1(2).Text & ""
                '
                End If
                        
            End If
            
            sSql = Replace(sSql, "$$$", ssql_selectes)
            
        End If

'------------------------------------깊이과목으로 조회한다.

        If chk1_4.Value = vbChecked Then
'            sSql = sSql + vbCrLf + "Union"
'            '//복습예정일로 문제지마스타(tu02) 선택
'            sSql = sSql + vbCrLf + "select subj,seq from tu02"
'            sSql = sSql + vbCrLf + "where userid =  '" & gUserid & "'"
'            If chk1_5(0).Value <> vbUnchecked Then
'                sSql = sSql + vbCrLf + "and reserve_ymd >='" & date2Str(str2Date(nTxt1_2.Text) - 1) & "'"
'                sSql = sSql + vbCrLf + "and reserve_ymd<'99999999'"
'            End If
'            If chk1_5(1).Value <> vbUnchecked Then
'                sSql = sSql + vbCrLf + "and reserve_ymd <='" & date2Str(str2Date(nTxt1_3.Text) - 1) & "'"
'            End If

            If chk1_5(0).Value <> vbUnchecked Then
                ssql2b = ssql2b + vbCrLf + "and reserve_ymd >='" & date2Str(str2Date(nTxt1_2.Text) - 1) & "'"
                ssql2b = ssql2b + vbCrLf + "and reserve_ymd<'99999999'"
                
                ssql3c = ssql3c + vbCrLf + "and reserve_ymd >='" & date2Str(str2Date(nTxt1_2.Text) - 1) & "'"
                ssql3c = ssql3c + vbCrLf + "and reserve_ymd<'99999999'"
                
            End If
            If chk1_5(1).Value <> vbUnchecked Then
                ssql2b = ssql2b + vbCrLf + "and reserve_ymd <='" & date2Str(str2Date(nTxt1_3.Text) - 1) & "'"
                ssql3c = ssql3c + vbCrLf + "and reserve_ymd <='" & date2Str(str2Date(nTxt1_3.Text) - 1) & "'"
            End If
            
            sSql = Replace(sSql, "and 2=2", ssql2b)
            sSql = Replace(sSql, "and 3=3", ssql3c)
            
        End If
        
        savedSql = sSql
        
        sSql = "select bb.subjnm,bb.subj,count(*) as cnt from (" & savedSql & ") aa , ts01 bb where aa.subj=bb.subj group by  bb.subjnm,bb.subj"
        
        Dim URS As ut_bRecordSet
        
        URS = Fn_SQLExec(sSql)
        Set lRs = URS.rs
        
        If URS.nrow = 0 Then
            If lst1_4.ListCount + lst1_2.ListCount > 0 Then
               MsgBox "시험지 만들 자료가 부족합니다.(어카운트 갱신을 시도하세요.-F7 버튼-)", vbExclamation
            Else
               MsgBox "시험지 만들 자료가 부족합니다.", vbExclamation
            End If
            Exit Function
        End If
        
        lst2_1.ClearList
        i = 0
        Do Until lRs.EOF
            Call lst2_1.AddItem("" & Chr(10) & lRs("subjnm") & Chr(10) & lRs("cnt") & Chr(10) & lRs("subj"))
            lst2_1.ListColumnCheck(i, 1) = 1
            i = i + 1
            lRs.MoveNext
        Loop
        
    Case 1
    '[2/4]
        opt3_1(0).Value = True
        mv3_1.Value = str2Date(Format(Now, "yyyyMMdd"))
'        mv3_2.Value = mv3_1.Value + 30
        nTxt3_1.Text = "1"
        nTxt3_2.Text = "10"
        
        cnt = 0
        subjcnt = 0
        For i = 0 To lst2_1.ListCount - 1
            If lst2_1.ListColumnCheck(i, 1) = 1 Then
                lst2_1.ListIndex = i
                cnt = cnt + CLng(lst2_1.ListColumnText(i, 3))
                If Len(str) = 0 Then
                    str = lst2_1.ListColumnText(i, 2)
                    lclsSchedule.sPocketNm = str
                    strFirstSubj = str
                Else
                    str = str & "," & lst2_1.ListColumnText(i, 2)
                    lclsSchedule.sPocketNm = strFirstSubj & " 외" & subjcnt
                End If
                
                subjcnt = subjcnt + 1
            End If
        Next
        
        lclsSchedule.nTotalCnt = cnt
        
        mv3_2.Value = mv3_1.Value + Fix(cnt / 100)
        
        txt3_1.Text = " ● 총 [" & Format(cnt, "#,###") & "]문항" & vbCrLf & " ● 선택과목: " & str
        Call mv3_1_SelChange(mv3_1.Value, mv3_1.Value, False)
    Case 2
    '[3/4]단게에서 [4/4]초기 설정
        
        nTxt4_1.Text = lclsSchedule.sPocketNm
        
        chk4_1.Value = vbChecked
        cbo4_1.Text = "☆"
    
        chk4_2.Value = vbChecked
        chk4_3.Value = vbChecked
        chk4_4.Value = vbChecked
        
        cbo4_2.Text = cbo4_2.List(0)
        cbo4_3.Text = cbo4_3.List(0)
        cbo4_4.Text = cbo4_4.List(0)
         
        txt4_5.Text = " "
        txt4_6.Text = " ~ "

    Case 3
    '[4/4]단계 후 시험지 생성
        
    Case 4
        Debug.Assert False
    End Select
    
    fraPorcess = True
Exit Function
errorTrap:
End Function

Private Function fraValidation(idx As Integer) As Boolean
    Dim i As Integer
    Dim cnt As Integer
    
    Select Case idx
    Case 0
        
        If (lst1_2.ListCount + lst1_4.ListCount) = 0 Then 'And chk1_4.Value = vbUnchecked Then
            MsgBox "시험지 만들 대상을 선택하세요.", vbExclamation
            Exit Function
        End If
        
        If chk1_4.Value = vbChecked Then
            If chk1_5(0).Value = vbUnchecked And chk1_5(1).Value = vbUnchecked Then
                
                MsgBox "복습예정 스케쥴의 시작 끝 일정을 선택하세요.", vbExclamation
                
                Exit Function
                
            End If
            
            If chk1_5(0).Value Then
                If Len(nTxt1_2.Text) <> 8 Then
                    MsgBox "시작일 오류!", vbExclamation
                    nTxt1_2.SetFocus
                    Exit Function
                End If
            End If
            
            If chk1_5(1).Value Then
                If Len(nTxt1_3.Text) <> 8 Then
                    MsgBox "종료일 오류!", vbExclamation
                    nTxt1_3.SetFocus
                    Exit Function
                End If
            End If
            
            If chk1_5(0).Value And chk1_5(1).Value Then
                If nTxt1_2.Text > nTxt1_3.Text Then
                    MsgBox "시작일은 종료일보다 작아야 합니다.", vbExclamation
                    nTxt1_2.SetFocus
                    Exit Function
                End If
            End If
            
        End If
        
        
    Case 1
    '[2/4] validation check
        cnt = 0
        For i = 0 To lst2_1.ListCount - 1
            If lst2_1.ListColumnCheck(i, 1) = 1 Then
                cnt = cnt + 1
            End If
        Next
        
        
        If cnt = 0 Then
            MsgBox "하나 이상의 과목을 선택하세요", vbExclamation
            Exit Function
        End If
    
    
    Case 2
    
        
        If opt3_1(0).Value Then
        
            lclsSchedule.nDayTotal = Fix(mv3_2.Value - mv3_1.Value) + 1
            
            lclsSchedule.nPerX = CDbl(nTxt3_1.Text)
            
            '시작일 종료일 x체크
            If date2Str(mv3_2) < date2Str(mv3_1.Value) Then
                MsgBox "종료일은 시작일보다 커야합니다.", vbExclamation
                Exit Function
            End If
            
            If CLng(nTxt3_1.Text) < 1 Then
                MsgBox "문항씩의 값은 1 이상이어야 합니다.", vbExclamation
                nTxt3_1.SetFocus
                Exit Function
            End If
            
            If lclsSchedule.nDayTotal < CInt(nTxt3_1.Text) Then
                MsgBox "일마다의 값은 [" & lclsSchedule.nDayTotal & "]이하여야 합니다.", vbExclamation
                nTxt3_1.SetFocus
                Exit Function
            End If
            
            lclsSchedule.nPerY = lclsSchedule.nTotalCnt / lclsSchedule.nDayTotal * CDbl(nTxt3_1.Text)
            
            nTxt3_2.Text = CLng(lclsSchedule.nPerY)
            
        End If
        
        If opt3_1(1).Value Then
            '시작일 종료일 x체크
            
            lclsSchedule.nPerY = CDbl(nTxt3_2.Text)
            
            If date2Str(mv3_2) <= date2Str(mv3_1.Value) Then
                MsgBox "종료일은 시작일보다 커야합니다.", vbExclamation
                Exit Function
            End If
            
            If CLng(nTxt3_2.Text) < 2 Then
                MsgBox "문항씩의 값은 2 이상이어야 합니다.", vbExclamation
                nTxt3_2.SetFocus
                Exit Function
            End If
            
            lclsSchedule.nPerX = (lclsSchedule.nDayTotal * CDbl(nTxt3_2.Text)) / lclsSchedule.nTotalCnt
            
            If lclsSchedule.nPerX < 1 Then
                MsgBox "종료일이 너무 작습니다. " & vbNewLine + vbNewLine & "종료일은 [" & Format(mv3_1.Value + CLng(lclsSchedule.nTotalCnt / CDbl(nTxt3_2.Text)), "YYYY-mm-dd") & "]일 보다 커야합니다.", vbExclamation
                nTxt3_2.SetFocus
                Exit Function
            End If
            
            If lclsSchedule.nPerX >= lclsSchedule.nDayTotal Then
                MsgBox "문항씩의 값이 너무 큽니다. " & vbNewLine + vbNewLine & " 전체문제수를 초과하였습니다.", vbExclamation
                nTxt3_2.SetFocus
                Exit Function
            End If
            
            nTxt3_1.Text = CLng(lclsSchedule.nPerX)
            
        End If
        
        If opt3_1(2).Value Then
            '시작일 x , y 체크
            lclsSchedule.nPerX = CDbl(nTxt3_1.Text)
            lclsSchedule.nPerY = CDbl(nTxt3_2.Text)
            
            If CLng(nTxt3_2.Text) > lclsSchedule.nTotalCnt Then
                MsgBox "일마다의 값이 전체문제수를 초과하였습니다.", vbExclamation
                nTxt3_2.SetFocus
                Exit Function
            End If
            
            If CLng(nTxt3_1.Text) < 1 Then
                MsgBox "문항씩의 값은 1 이상이어야 합니다.", vbExclamation
                nTxt3_1.SetFocus
                Exit Function
            End If
            
            If CLng(nTxt3_2.Text) < 2 Then
                MsgBox "일마다의 값은 2 이상이어야 합니다.", vbExclamation
                nTxt3_2.SetFocus
                Exit Function
            End If
                        
            mv3_2.Value = mv3_1.Value + CLng(lclsSchedule.nTotalCnt / CDbl(nTxt3_2.Text) * (CDbl(nTxt3_1.Text)))
            
            '시작일 종료일 x체크
            If date2Str(mv3_2) <= date2Str(mv3_1.Value) Then
                MsgBox "종료일은 시작일보다 커야합니다.", vbExclamation
                Exit Function
            End If
            
        End If
        
    Case 3
        If Len(lclsSchedule.getPocketNm) = 0 Then
            MsgBox "시험지명을 입력하세요.", vbExclamation
            Exit Function
        End If
        
        If chk4_2.Value = vbUnchecked And chk4_3.Value = vbUnchecked And chk4_4.Value = vbUnchecked Then
            MsgBox "하위 시험지명의 출력형식을 하나이상 입력하세요.", vbExclamation
            Exit Function
        End If
    Case 4
    
    End Select
    
    fraValidation = True
End Function

Private Sub nTxt3_1_Change()
'lclsSchedule.nPerX = nTxt3_1.Text
End Sub

Private Sub nTxt3_2_Change()
'lclsSchedule.nPerY = nTxt3_2.Text
End Sub

Private Sub nTxt4_1_Change()
    lclsSchedule.sPocketNm = nTxt4_1.Text
    txt4_3.Text = lclsSchedule.getPocketNm
End Sub

Private Sub opt2_1_Click(Index As Integer)
    chk2_1.Value = vbChecked
    chk2_1_Click
End Sub

Private Sub opt3_1_Click(Index As Integer)
    
Call LockWindowUpdate(Me.hwnd)
    
lb3_4.Visible = True
nTxt3_2.Visible = True
lb3_3.Visible = True
nTxt3_1.Visible = True
lb3_2.Visible = True
mv3_2.Visible = True

lclsSchedule.selMethod = Index

Select Case Index

Case 0
    '시작일 종료일 x일마다
    lb3_4.Visible = False
    nTxt3_2.Visible = False
    
Case 1
    '시작일 종료일 y문항씩
    lb3_3.Visible = False
    nTxt3_1.Visible = False
    
Case 2
    '시작일 x일마다 y문항씩
    lb3_2.Visible = False
    mv3_2.Visible = False
    
End Select

Call LockWindowUpdate(0&)

End Sub

Private Sub txt4_5_Change()
lclsSchedule.sSubChar1 = txt4_5.Text
txt4_4.Text = lclsSchedule.getSubPocketNm
End Sub

Private Sub txt4_6_Change()
lclsSchedule.sSubChar2 = txt4_6.Text
txt4_4.Text = lclsSchedule.getSubPocketNm
End Sub

Private Sub txt4_7_Change()
lclsSchedule.sSubChar3 = txt4_7.Text
txt4_4.Text = lclsSchedule.getSubPocketNm
End Sub

Private Sub makeprocess()

Dim lRs As ADODB.Recordset

sSql = "select count(*) cnt from tm01 where userid='" + gUserid + "' and ymd = date_format(current_date,'%Y%m%d')"
Set lRs = Fn_SQLExec(sSql).rs

If 4 < lRs(0) Then
    MsgBox "하루에 계획을 최대 5개 까지만 만들 수 있습니다.", vbOKOnly + vbExclamation, Me.Caption
    Exit Sub
End If

sSql = "insert into tm01 values ('" + gUserid + "',sysdate(),date_format(current_date,'%Y%m%d'))"
Fn_SQLExec (sSql)


'1. 테이블 데이터 준비
If Not pro1() Then Exit Sub
'2. 시험지코드준비
If Not pro2() Then Exit Sub
'3. 시험지만들기
If Not pro3() Then Exit Sub
'4. 시험지 만들었던 임시 테이블 데이터 삭제
If Not pro4() Then Exit Sub
'5. 종료

parent.mnuRefresh_Click

Unload Me

End Sub

'==============================================================================
'1. 테이블 데이터 준비
'==============================================================================
Private Function pro1() As Boolean

ProgressBar1.ToolTipText = "데이터 준비중"
ProgressBar1.Value = 0
ProgressBar1.Max = 4
    
    sSql = "delete from tt01 where userid = '" & gUserid & "'"
    If Fn_SQLExec(sSql).Error Then Exit Function
ProgressBar1.Value = 2
    sSql = "delete from tt02 where userid = '" & gUserid & "'"
    If Fn_SQLExec(sSql).Error Then Exit Function
ProgressBar1.Value = 3
    sSql = "delete from tt03 where userid = '" & gUserid & "'"
    If Fn_SQLExec(sSql).Error Then Exit Function
ProgressBar1.Value = 4
    pro1 = True
End Function
'==============================================================================
'2. 시험지코드준비
'==============================================================================
Private Function pro2() As Boolean
    
    Dim i As Integer
    Dim sSubj As String
    Dim sortkey As Integer
    Dim lRs As ADODB.Recordset
    Dim nansu As Long
    
    Dim utb As ut_bRecordSet
    
    sortkey = 1
    
    For i = 0 To lst2_1.ListCount - 1
        If lst2_1.ListColumnCheck(i, 1) = 1 Then
            
            sSubj = lst2_1.ListColumnText(i, 4)
            sSql = "insert into tt01(userid,subj,sortkey) values ('" & gUserid & "','" & sSubj & "'," & sortkey & ")"
            If Fn_SQLExec(sSql).Error Then Exit Function
            sortkey = sortkey + 1
        End If
    Next
    
    
    If InStr(LCase(STRCON), "mdb") > 0 Then
    
        'Set lRs = Fn_SQLExec(savedSql, , , True).rs
        utb = Fn_SQLExec(savedSql)
        Set lRs = utb.rs
ProgressBar1.ToolTipText = "문제섞는중..."
ProgressBar1.Value = 0
ProgressBar1.Value = utb.nrow
        Randomize
        Do Until lRs.EOF
        
ProgressBar1.Value = ProgressBar1.Value + 1
            nansu = CLng(Rnd * 1000000)
            sSql = "insert into tt02(userid,subj,seq,nansu) values('" & gUserid & "','"
            sSql = sSql & lRs("subj") & "'," & lRs("seq") & "," & nansu & ")"
            Fn_SQLExec (sSql)
            lRs.MoveNext
        Loop
        
        lRs.Close
    Else
    
            sSql = "insert into tt02(userid,subj,seq,nansu) (select '" & gUserid & "', a.subj,a.seq,rand()*1000000 from (" & savedSql & ") a)"
            
            Call Fn_SQLExec(sSql)
    
    End If
    
    sSql = "SELECT a.userid, a.subj, a.sortkey, b.seq, b.nansu FROM tt01 AS a, tt02 AS b WHERE a.userid=b.userid and a.subj=b.subj and a.userid='" & gUserid & "' "
    
    If lclsSchedule.bQuizSwap Then
        If lclsSchedule.bQuizSwap_all Then
            sSql = sSql & "order by nansu"
        ElseIf lclsSchedule.bQuizSwap_subj Then
            sSql = sSql & "order by sortkey,nansu"
        End If
    Else
        sSql = sSql & "order by sortkey, seq"
    End If
    
    utb = Fn_SQLExec(sSql)
    
    Set lRs = utb.rs ' Fn_SQLExec(sSql, , , True).rs
    
    i = 1
    Dim chasu As Long
    chasu = 1
    Dim tDt As Date
    Dim num As Long
    Dim cnt As Long
    Dim fromIlJa As String
    Dim toIlJa As String
    Dim affected As Integer
    
    fromIlJa = date2Str(mv3_1.Value)
    
    Debug.Assert lclsSchedule.nPerX >= 1
    
    toIlJa = date2Str(mv3_1.Value + CLng(chasu * lclsSchedule.nPerX) - 1)
    num = 1
    cnt = 0
    
ProgressBar1.ToolTipText = "[2단계]/[4단계]:데이터 수집중"
ProgressBar1.Value = 0
ProgressBar1.Max = utb.nrow

    Dim circleFunctionResultOfDay1 As Double
    Dim sum1 As Double
    Dim factor1 As Double
    
    If lclsSchedule.nDayTotal = 1 Then
        lclsSchedule.nDayTotal = 2 '1일이면 2일로 바꿈.
    End If
    
    circleFunctionResultOfDay1 = 4# * lclsSchedule.nTotalCnt / PI / (lclsSchedule.nDayTotal - 1)
    
    Do Until lRs.EOF
    
ProgressBar1.Value = ProgressBar1.Value + 1
status.Text = CStr(ProgressBar1.Value) & "/" & ProgressBar1.Max & " 진행중{[2단계]/[4단계] 기초자료수집중...}"
        sSql = "insert into tt03(userid,subj,seq,chasu,fromilja,toilja,num) "
        sSql = sSql & " values('" & gUserid & "','" & lRs("subj") & "'," & lRs("seq") & "," & chasu & ",'" & fromIlJa & "','" & toIlJa & "'," & num & " )  "
        
        affected = Fn_SQLExec(sSql).nrow
        Debug.Assert affected = 1
                
        num = num + 1
        cnt = cnt + 1
        
        Select Case opt3_2(0).Value
        
        Case True
        
            If chasu * lclsSchedule.nPerY <= cnt Then '다음날짜의 문항으로 항목을 변경될 조건 <선형학습>
                chasu = chasu + 1
                fromIlJa = date2Str(str2Date(toIlJa) + 1)
                toIlJa = date2Str(mv3_1.Value + CLng(chasu * lclsSchedule.nPerX) - 1)
                num = 1
            End If
            
        Case False
        
            '아래는 다음날짜의 문항으로 항목을 변경될 조건 <타원스케쥴학습>
            'http://blog.naver.com/iq_up?Redirect=Log&logNo=100058343263 에서 엑셀파일이 계산 근거인
            '파일명 serise11-iq_up.xls
            
            factor1 = (chasu - 1) ^ 2 / (lclsSchedule.nDayTotal / lclsSchedule.nPerX - 1) ^ 2
            If factor1 > 1 Then Exit Do
            If sum1 + ((1 - factor1) * (circleFunctionResultOfDay1) ^ 2) ^ 0.5 * lclsSchedule.nPerX <= cnt Then
                sum1 = sum1 + ((1 - factor1) * (circleFunctionResultOfDay1) ^ 2) ^ 0.5 * lclsSchedule.nPerX
                chasu = chasu + 1
                fromIlJa = date2Str(str2Date(toIlJa) + 1)
                toIlJa = date2Str(mv3_1.Value + CLng(chasu * lclsSchedule.nPerX) - 1)
                num = 1
            End If
        
        End Select
        
        lRs.MoveNext
    Loop
    
    lRs.Close
status.Text = ""
pro2 = True
End Function
'==============================================================================
'3. 시험지만들기
'==============================================================================
Private Function pro3() As Boolean
On Error GoTo ErrTrap
Dim RS2 As New ADODB.Recordset
Dim SSQL1 As String

Dim makecnt As Long

Dim ymd As String
Dim OBJTABLE As String
Dim i As Long
Dim affected  As Long
Dim maxCode As Long
Dim pMaxCode As Long
Dim URS As ut_bRecordSet

ymd = GETYMD()

Dim chasu As Long

OBJTABLE = "TT03" 'RS2(1)

SSQL1 = "SELECT " & OBJTABLE & ".* FROM " & OBJTABLE & " WHERE USERID='" & gUserid & "' order by chasu,num"

Con.BeginTrans
status.Text = "[3단계]/[4단계]처리중입니다..."
URS = Fn_SQLExec(SSQL1)

Set RS2 = URS.rs ' Fn_SQLExec(SSQL1, , , True).rs

i = 0
makecnt = 0
Dim makeOrder As Long
Dim preChasu As Long
Dim pn As String

preChasu = -1

If RS2.EOF = False Then
    pn = lclsSchedule.getPocketNm
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    chasu = getMaxTableVal("CHASU", "tp02", "WHERE USERID='" & gUserid & "' AND (POCKETNM='" & Replace(pn, "!", "") & "' or POCKETNM='!" & Replace(pn, "!", "") & "')")
    pMaxCode = maxCode
    
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "',0,0,0,1,2,'" & date2Str(mv3_1.Value) & "','" & date2Str(mv3_2.Value) & "')"
    Fn_SQLExec (SSQL1)
    
    SSQL1 = "select count(*) from tt03 where userid='" & gUserid & "'"

status.Text = "[3단계]/[4단계] 데이터 갯수를 세고있습니다...."

    makecnt = Fn_SQLExec(SSQL1).rs(0)
    
    makeOrder = 10
    If makecnt < 10 Then
        makeOrder = makecnt
    ElseIf makecnt >= 100 Then
        makeOrder = Fix(makecnt ^ 0.5) + 1 '너무 지루한 학습 방지.
    End If

    SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "'," & 1 & "," & 1 & "," & makeOrder & ")"
status.Text = "[3단계]/[4단계] 마스터 데이터 스케쥴을 입력하고 있습니다..."
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0
End If

i = 1


ProgressBar1.ToolTipText = "[3단계]/[4단계] 데이터 이동중..."
status.Text = ProgressBar1.ToolTipText
ProgressBar1.Value = 0
ProgressBar1.Max = URS.nrow

Do Until RS2.EOF
ProgressBar1.Value = ProgressBar1.Value + 1
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected = 1
    status.Text = ProgressBar1.Value & "/" & ProgressBar1.Max & " [3단계]/[4단계] 스케쥴 계산중..."

    RS2.MoveNext
    i = i + 1
Loop

RS2.MoveFirst

preChasu = -1
    
ProgressBar1.ToolTipText = "[3단계]/[4단계] 스케쥴 데이터 생성중..."
ProgressBar1.Value = 0
ProgressBar1.Max = URS.nrow

Do Until RS2.EOF
ProgressBar1.Value = ProgressBar1.Value + 1



    If RS2("chasu") <> preChasu Then
        'sub 시험지의 첫단에서 실행됨
        pn = lclsSchedule.getSubPocketNm(RS2("chasu"), str2Date(RS2("fromilja")), str2Date(RS2("toilja")))
        maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
        chasu = getMaxTableVal("CHASU", "tp02", "WHERE USERID='" & gUserid & "' AND (POCKETNM='" & Replace(pn, "!", "") & "' or POCKETNM='!" & Replace(pn, "!", "") & "')")
        
        SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "'," & pMaxCode & ",0,0,1,2,'" & RS2("fromilja") & "','" & RS2("toilja") & "')"
        Fn_SQLExec (SSQL1)
    
        SSQL1 = "select count(*) from tt03 where userid='" & gUserid & "' and  chasu=" & RS2("chasu")
        
        makecnt = Fn_SQLExec(SSQL1).rs(0)
        
        makeOrder = 10
        If makecnt < 10 Then
            makeOrder = makecnt
            If makeOrder = 1 Then
               makeOrder = 2
            End If
        End If
        
        If makecnt < 10 Then
        makeOrder = makecnt
    ElseIf makecnt >= 100 Then
        makeOrder = Fix(makecnt ^ 0.5) + 1 '너무 지루한 학습 방지.
    End If
    
        SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "'," & 1 & "," & 1 & "," & makeOrder & ")"
        affected = Fn_SQLExec(SSQL1).nrow
        Debug.Assert affected > 0
        preChasu = RS2("chasu")
    End If
    
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & RS2("num") & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    
    RS2.MoveNext
status.Text = ProgressBar1.Value & "/" & ProgressBar1.Max & "[3단계]/[4단계] 스케쥴 데이터 생성중..."
Loop
RS2.Close
Con.CommitTrans

status.Text = ""

pro3 = True

Exit Function
ErrTrap:
Con.RollbackTrans


End Function

'==============================================================================
'3. 시험지만들기 빨리 만들기 로직
'==============================================================================
Private Function pro3_F1() As Boolean
On Error GoTo ErrTrap
Dim RS2 As New ADODB.Recordset
Dim SSQL1 As String

Dim makecnt As Long

Dim ymd As String
Dim OBJTABLE As String
Dim i As Long
Dim affected  As Long
Dim maxCode As Long
Dim pMaxCode As Long
Dim URS As ut_bRecordSet

ymd = GETYMD()

Dim chasu As Long

OBJTABLE = "TT03" 'RS2(1)

SSQL1 = "SELECT " & OBJTABLE & ".* FROM " & OBJTABLE & " WHERE USERID='" & gUserid & "' order by chasu,num"

Con.BeginTrans

URS = Fn_SQLExec(SSQL1)
Set RS2 = URS.rs ' Fn_SQLExec(SSQL1, , , True).rs

i = 0
makecnt = 0
Dim makeOrder As Long
Dim preChasu As Long
Dim pn As String

preChasu = -1

ProgressBar1.ToolTipText = "데이터 이동중..."
ProgressBar1.Value = 0
ProgressBar1.Max = URS.nrow

If RS2.EOF = False Then
    pn = lclsSchedule.getPocketNm
    maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
    chasu = getMaxTableVal("CHASU", "tp02", "WHERE USERID='" & gUserid & "' AND (POCKETNM='" & Replace(pn, "!", "") & "' or POCKETNM='!" & Replace(pn, "!", "") & "')")
    pMaxCode = maxCode
    
    SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "',0,0,0,1,2,'" & date2Str(mv3_1.Value) & "','" & date2Str(mv3_2.Value) & "')"
    Fn_SQLExec (SSQL1)
    
    SSQL1 = "select count(*) from tt03 where userid='" & gUserid & "'"
    
    makecnt = Fn_SQLExec(SSQL1).rs(0)
    
    makeOrder = 10
'    If makecnt < 10 Then
'        makeOrder = makecnt
'    End If
    
    If makecnt < 10 Then
        makeOrder = makecnt
    ElseIf makecnt >= 100 Then
        makeOrder = Fix(makecnt ^ 0.5) + 1 '너무 지루한 학습 방지.
    End If

    SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "'," & 1 & "," & 1 & "," & makeOrder & ")"
    affected = Fn_SQLExec(SSQL1).nrow
    Debug.Assert affected > 0
End If

i = 1
'DDD

SSQL1 = "INSERT INTO TP03 (select '" & pn & "'," & chasu & ",'" & gUserid & "'," & i & ",SUBJ,SEQ,0,0,0,'" & ymd & "' FROM " & OBJTABLE & " WHERE USERID='" & gUserid & "' order by chasu,num)" ' VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & i & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
Debug.Print SSQL1
    
i = i + Fn_SQLExec(SSQL1).nrow
    
preChasu = -1
    
ProgressBar1.ToolTipText = "스케쥴 데이터 생성중..."
ProgressBar1.Value = 0
ProgressBar1.Max = URS.nrow

Do Until RS2.EOF

ProgressBar1.Value = ProgressBar1.Value + 1
    If RS2("chasu") <> preChasu Then
        'sub 시험지의 첫단에서 실행됨
        pn = lclsSchedule.getSubPocketNm(RS2("chasu"), str2Date(RS2("fromilja")), str2Date(RS2("toilja")))
        maxCode = getMaxTableVal("code", "tp02", "where userid='" & gUserid & "'")
        chasu = getMaxTableVal("CHASU", "tp02", "WHERE USERID='" & gUserid & "' AND (POCKETNM='" & Replace(pn, "!", "") & "' or POCKETNM='!" & Replace(pn, "!", "") & "')")
        
        SSQL1 = "INSERT INTO TP02 VALUES(" & maxCode & ",'" & pn & "'," & chasu & ",'" & gUserid & "',0,'" & ymd & "'," & pMaxCode & ",0,0,1,2,'" & RS2("fromilja") & "','" & RS2("toilja") & "')"
        Fn_SQLExec (SSQL1)
    
        SSQL1 = "select count(*) from tt03 where userid='" & gUserid & "' and  chasu=" & RS2("chasu")
        
        makecnt = Fn_SQLExec(SSQL1).rs(0)
        
        makeOrder = 10
        If makecnt < 10 Then
            makeOrder = makecnt
            If makeOrder = 1 Then
               makeOrder = 2
            End If
        ElseIf makecnt >= 100 Then
            makeOrder = Fix(makecnt ^ 0.5) + 1 '너무 지루한 학습 방지.
        End If
    
        SSQL1 = "insert into tu03 values('" & pn & "'," & chasu & ",'" & gUserid & "'," & 1 & "," & 1 & "," & makeOrder & ")"
        affected = Fn_SQLExec(SSQL1).nrow
        Debug.Assert affected > 0
        preChasu = RS2("chasu")
    End If
    
    SSQL1 = "INSERT INTO TP03 VALUES('" & pn & "'," & chasu & ",'" & gUserid & "'," & RS2("num") & ",'" & RS2("SUBJ") & "'," & RS2("SEQ") & ",0,0,0,'" & ymd & "')"
    Debug.Print SSQL1
    
    affected = Fn_SQLExec(SSQL1).nrow
    makecnt = makecnt + affected
    Debug.Assert affected = 1
    
    RS2.MoveNext
Loop
RS2.Close
Con.CommitTrans

pro3_F1 = True

Exit Function
ErrTrap:
Con.RollbackTrans


End Function
'==============================================================================
'4. 시험지 만들었던 임시 테이블 데이터 삭제
'==============================================================================
Private Function pro4() As Boolean

ProgressBar1.ToolTipText = "[4단계]/[4단계] 임시 데이터 삭제중..."

ProgressBar1.Value = 0
ProgressBar1.Max = 4

If Not IDEMODE Then
ProgressBar1.Value = 1
    sSql = "delete from tt01 where userid = '" & gUserid & "'"
    If Fn_SQLExec(sSql).Error Then Exit Function
    sSql = "delete from tt02 where userid = '" & gUserid & "'"
    
ProgressBar1.Value = 2
    If Fn_SQLExec(sSql).Error Then Exit Function
    sSql = "delete from tt03 where userid = '" & gUserid & "'"
    
ProgressBar1.Value = 3
    If Fn_SQLExec(sSql).Error Then Exit Function
ProgressBar1.Value = 4
End If
    
pro4 = True
End Function


Function selectSeriesTP01(str As String) As String
Dim localSql As String
Dim localStr As String
Dim alertMsg As String
Dim str1 As String
Dim i As Integer
Dim arr1 As Variant
Dim localrs As ADODB.Recordset
Dim str2 As String

On Error GoTo errorTrap

Dim re As New RegExp

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = False
    re.Pattern = "[^\(\s*\)$]"
    
    
'    str2 = re.Replace(str, "")

localSql = "select cond,pocketnm from tp01 where pocketnm in (" & str & ") and cond like '%select%'"

Dim nResult As ut_bRecordSet
localStr = ""
Set localrs = Fn_SQLExec(localSql).rs

'If nResult.Error = False Then
'    Set localRs = nResult.rs
    Do Until localrs.EOF
        localSql = localrs(0)
        arr1 = Split(localSql, "|")
        If UBound(arr1) = 1 Then
            re.Pattern = "^\("
            arr1(1) = re.Replace(arr1(1), "")
            re.Pattern = "\)$"
            arr1(1) = re.Replace(arr1(1), "")
            
            localStr = localStr & " " & arr1(1) & " Union"
'            alertMsg = alertMsg & "'" & localrs(1) & "',"
            
            i = i + 1
            If i = 1 Then
                str1 = arr1(1)
            End If
        End If
        localrs.MoveNext
    Loop
'End If

localStr = Left(localStr, Len(localStr) - Len(" Union"))

'If InStr(localStr, "Union") > 0 Then
'    alertMsg = Left(alertMsg, Len(alertMsg) - 1)
'    alertMsg = alertMsg & "  의 과목중에서는 한개만 선택됩니다."
'    MsgBox alertMsg, vbCritical
'    'localStr = "(" & localStr & ")"
'    localStr = str1
'End If

selectSeriesTP01 = localStr
Exit Function
errorTrap:
End Function




Private Sub dtp1_CloseUp()
nTxt1_2.Text = date2Str(dtp1.Value)
End Sub

Private Sub dtp1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dtp1.Value = str2Date(nTxt1_2.Text)
End Sub

Private Sub dtp2_CloseUp()
nTxt1_3.Text = date2Str(dtp2.Value)
End Sub

Private Sub dtp2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dtp2.Value = str2Date(nTxt1_3.Text)
End Sub


