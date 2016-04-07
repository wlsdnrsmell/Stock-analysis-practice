VERSION 5.00
Begin VB.Form FutureCur 
   BorderStyle     =   1  '단일 고정
   Caption         =   "선물현재가"
   ClientHeight    =   7155
   ClientLeft      =   4020
   ClientTop       =   1650
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5910
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5655
      Begin VB.Label lb_days 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   5100
         TabIndex        =   120
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label lb_sellprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   119
         Top             =   1320
         Width           =   520
      End
      Begin VB.Label lb_buyprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   118
         Top             =   1680
         Width           =   520
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   117
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   116
         Top             =   3120
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   115
         Top             =   2760
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   114
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   113
         Top             =   600
         Width           =   520
      End
      Begin VB.Label lb_kospidaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   2160
         TabIndex        =   112
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label lb_midaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   2160
         TabIndex        =   111
         Top             =   960
         Width           =   570
      End
      Begin VB.Label lb_bottomdaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4920
         TabIndex        =   110
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lb_topdaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4920
         TabIndex        =   109
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lb_lowdaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4920
         TabIndex        =   108
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lb_highdaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4920
         TabIndex        =   107
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lb_startdaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4920
         TabIndex        =   106
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lb_gijundaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4920
         TabIndex        =   105
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lb_curdaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   2160
         TabIndex        =   104
         Top             =   240
         Width           =   520
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "베이시스"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   120
         TabIndex        =   103
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lb_basis 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   102
         Top             =   2400
         Width           =   840
      End
      Begin VB.Label lb_endday 
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   95
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label lb_top 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   94
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lb_low 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   93
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lb_high 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   92
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lb_start 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   91
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lb_murisk 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   90
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lb_bottomday 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   89
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lb_bottom 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   88
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lb_topday 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   4200
         TabIndex        =   87
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lb_gijun 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   86
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lb_geonmi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   85
         Top             =   3480
         Width           =   840
      End
      Begin VB.Label lb_gratio 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   84
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lb_iron 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   83
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lb_kospi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   82
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label lb_mi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   81
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lb_vol 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   80
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lb_curprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1320
         TabIndex        =   79
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lb_buyprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0C0FF&
         Caption         =   "Label3"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   49
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label lb_sellprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFC0C0&
         Caption         =   "Label3"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   43
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "최종거래일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   3000
         TabIndex        =   23
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "무위험이율"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   22
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "    일 자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   21
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "최 저 가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   20
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "    일 자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "최 고 가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "저     가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "고     가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "시     가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "기 준 가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "전일미결제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "괴 리 율"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "이론가격"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "KOSPI200"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "매수호가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "매도호가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "미결제약정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "거 래 량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "현 재 가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Cmd_end 
      Caption         =   "종  료"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cb_Fjongmok 
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lb_compare 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Caption         =   "Label3"
      Height          =   255
      Left            =   2400
      TabIndex        =   101
      Top             =   6800
      Width           =   1935
   End
   Begin VB.Label lb_ha 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Caption         =   "Label3"
      Height          =   255
      Left            =   4800
      TabIndex        =   100
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lb_haprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Caption         =   "Label3"
      Height          =   255
      Left            =   4800
      TabIndex        =   99
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lb_sang 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Caption         =   "Label3"
      Height          =   255
      Left            =   3120
      TabIndex        =   98
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lb_sangprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Caption         =   "Label3"
      Height          =   255
      Left            =   3120
      TabIndex        =   97
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lb_gijun 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   96
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   78
      Top             =   6800
      Width           =   690
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00C4C4C4&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   77
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00CACACA&
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   76
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00D6D6D6&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   75
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00EAEAEA&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   74
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   73
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label3"
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   72
      Top             =   6800
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C4C4C4&
      Caption         =   "Label3"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   71
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00CACACA&
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   70
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00D6D6D6&
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   69
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00EAEAEA&
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   68
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   67
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   66
      Top             =   6800
      Width           =   690
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label3"
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   65
      Top             =   6800
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C4C4C4&
      Caption         =   "Label3"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   64
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00CACACA&
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   63
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00D6D6D6&
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   62
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00EAEAEA&
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   61
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   60
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C4C4C4&
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   59
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00CACACA&
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   58
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00D6D6D6&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   57
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00EAEAEA&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   56
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   55
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   54
      Top             =   6510
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   53
      Top             =   6270
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   52
      Top             =   6030
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   51
      Top             =   5790
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   50
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   48
      Top             =   6510
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   47
      Top             =   6270
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   46
      Top             =   6030
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   45
      Top             =   5790
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   44
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label label 
      BorderStyle     =   1  '단일 고정
      Caption         =   "매수호가"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   3360
      TabIndex        =   42
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "매수잔량"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   4320
      TabIndex        =   41
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "(건수)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   5160
      TabIndex        =   40
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '단일 고정
      Caption         =   "매도호가"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   2400
      TabIndex        =   39
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "매도잔량"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   1560
      TabIndex        =   38
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "(건수)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   960
      TabIndex        =   37
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  5  차"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   120
      TabIndex        =   36
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  4  차"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  3  차"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   120
      TabIndex        =   34
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  2  차"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   120
      TabIndex        =   33
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  1  차"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   32
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " 구  분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   120
      TabIndex        =   31
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "잔량총계"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   30
      Top             =   6800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "하   한"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   4080
      TabIndex        =   29
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "하한가"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   4080
      TabIndex        =   28
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "상   한"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   2400
      TabIndex        =   27
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "상한가"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   2400
      TabIndex        =   26
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "써킷브레이커"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   25
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "기    준    가"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lb_jongmok 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "종목"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   495
   End
End
Attribute VB_Name = "FutureCur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public futurelist As CpFutureCode
Public fmobj As FutureMst
Attribute fmobj.VB_VarHelpID = -1
Public fwide As FutureWide
Public WithEvents fcobj As FutureCurr
Attribute fcobj.VB_VarHelpID = -1
Public WithEvents k200obj As FutureK200
Attribute k200obj.VB_VarHelpID = -1
Public cur_sb_callflag As Boolean
Sub SetColor(Value, obj As Object)
    If (Value > 0) Then
        obj.ForeColor = RGB(255, 0, 0)
    ElseIf (Value < 0) Then
        obj.ForeColor = RGB(0, 0, 255)
    Else
        obj.ForeColor = &H0
    End If
    obj = FormatNumber(Value, 0)
End Sub
Sub Set_start_high_low_price(startindex, obj As Object)
    lb_start.Caption = Format(obj.GetHeaderValue(startindex), "##0.00") + "  "   '시가
    lb_high.Caption = Format(obj.GetHeaderValue(startindex + 1), "##0.00") + "  " '고가
    lb_low.Caption = Format(obj.GetHeaderValue(startindex + 2), "##0.00") + "  " '저가
End Sub

Sub SetjongmokCombo()
    For i = 0 To futurelist.GetCount - 1
        cb_Fjongmok.AddItem futurelist.GetData(0, i)
    Next i
    cb_Fjongmok.Text = futurelist.GetData(0, 0)
End Sub
Sub Call_sb_Object()
    fcobj.Unsubscribe
    fcobj.SetInputValue 0, cb_Fjongmok.Text
    fcobj.SubscribeLatest
    
    k200obj.Unsubscribe
    k200obj.SetInputValue 0, cb_Fjongmok.Text
    k200obj.SubscribeLatest
End Sub
Sub Set_daebi(index1, index2, obj As Object, lb_obj As Object, flag As Boolean)
    tmp = obj.GetHeaderValue(index1) - obj.GetHeaderValue(index2)
    If flag Then
        SetColor tmp, lb_obj
    End If
    lb_obj = "(" + Format(tmp, "##0.00") + ")"
End Sub
Sub Set_gRatio(curr_index, iron_index, obj As Object, k200_flag As Boolean)
    '☞ 괴리율 = (선물현재가 - 선물이론지수) / 선물이론지수 * 100 (%)
    If cb_Fjongmok.ListIndex < 7 Then
        If k200_flag Then
            If cur_sb_callflag Then
                tmp = (fcobj.GetHeaderValue(1) - obj.GetHeaderValue(iron_index)) / obj.GetHeaderValue(iron_index) * 100
            Else
                tmp = (fmobj.GetHeaderValue(71) - obj.GetHeaderValue(iron_index)) / obj.GetHeaderValue(iron_index) * 100
            End If
        Else
            tmp = (obj.GetHeaderValue(curr_index) - obj.GetHeaderValue(iron_index)) / obj.GetHeaderValue(iron_index) * 100
        End If
        SetColor tmp, lb_gratio
        lb_gratio.Caption = Format(tmp, "##0.00") + " %"
    Else
        tmp = 0
        SetColor tmp, lb_gratio
        lb_gratio.Caption = Format(tmp, "##0.00") + " %"
    End If
End Sub
Sub Set_bid(startindex, obj As Object)
    lb_sellprice(0) = Format(obj.GetHeaderValue(startindex), "##0.00") '16
    For i = 1 To 5
        lb_sellprice(i) = Format(obj.GetHeaderValue(startindex + i - 1), "##0.00") '16
    Next
    For i = 1 To 6
        lb_sellsu(i) = Format(obj.GetHeaderValue(startindex + 5 + i - 1), "#,###") + "  " '21
    Next
    For i = 1 To 6
        lb_sellgeonsu(i) = "( " + Format(obj.GetHeaderValue(startindex + 11 + i - 1), "#,###") + " )" '27
    Next
    
    lb_buyprice(0) = Format(obj.GetHeaderValue(startindex + 17), "##0.00") '33
    For i = 1 To 5
        lb_buyprice(i) = Format(obj.GetHeaderValue(startindex + 17 + i - 1), "##0.00") '33
    Next
    For i = 1 To 6
        lb_buysu(i) = Format(obj.GetHeaderValue(startindex + 22 + i - 1), "#,###") + "  " '38
    Next
    For i = 1 To 6
        lb_buygeonsu(i) = "( " + Format(obj.GetHeaderValue(startindex + 28 + i - 1), "#,###") + " )" '44
    Next

    tmp = obj.GetHeaderValue(startindex + 27) - obj.GetHeaderValue(startindex + 10) '43-26
    SetColor tmp, lb_compare
    lb_compare = Format(tmp, "#,###")
End Sub
Private Sub Form_Load()
    Set futurelist = New CpFutureCode
    Set fmobj = New FutureMst
    Set fcobj = New FutureCurr
    Set fwide = New FutureWide
    Set k200obj = New FutureK200

    Call SetjongmokCombo
    cur_sb_callflag = False
    Call data_susin
End Sub
Private Sub k200obj_Received()
    lb_kospi.Caption = Format(k200obj.GetHeaderValue(1), "##0.00")  '코스피
    
    tmp = (k200obj.GetHeaderValue(1) - fmobj.GetHeaderValue(89)) + fmobj.GetHeaderValue(91)
    SetColor fmobj.GetHeaderValue(91), lb_kospidaebi
    lb_kospidaebi = "(" + Format(tmp, "##0.00") + ")"

    SetColor k200obj.GetHeaderValue(2), lb_basis
    lb_basis.Caption = Format(k200obj.GetHeaderValue(2), "0.00")  '베이시스
    
    lb_iron.Caption = Format(k200obj.GetHeaderValue(3), "##0.00") '이론가
    Call Set_gRatio(0, 3, k200obj, True)
End Sub
Private Sub fcobj_Received()
    lb_curprice.Caption = Format(fcobj.GetHeaderValue(1), "##0.00") '현재가
    
    SetColor fcobj.GetHeaderValue(2), lb_curdaebi
    lb_curdaebi.Caption = "(" + Format(fcobj.GetHeaderValue(2), "#0.00") + ")" '현재가(전일대비)
    
    lb_vol.Caption = Format(fcobj.GetHeaderValue(13), "#,###")   '거래량
    
    lb_mi.Caption = Format(fcobj.GetHeaderValue(14), "#,###")  '미결제약정
    
    SetColor fcobj.GetHeaderValue(5), lb_basis
    lb_basis.Caption = Format(fcobj.GetHeaderValue(5), "0.00")   '베이시스
    
    lb_iron.Caption = Format(fcobj.GetHeaderValue(3), "##0.00")  '이론가
    
    lb_top.Caption = Format(fcobj.GetHeaderValue(10), "##0.00") + "  "  '최고가
    lb_bottom.Caption = Format(fcobj.GetHeaderValue(11), "##0.00") + "  "   '최저가
    ''''''''''''''''''''''
    Call Set_start_high_low_price(7, fcobj)
    Call Set_daebi(1, 11, fcobj, lb_bottomdaebi, False)
    Call Set_daebi(1, 10, fcobj, lb_topdaebi, False)
    Call Set_daebi(9, 6, fcobj, lb_lowdaebi, True)
    Call Set_daebi(8, 6, fcobj, lb_highdaebi, True)
    Call Set_daebi(7, 6, fcobj, lb_startdaebi, True)
    ''''''''''''''''''''''
    tmp = fcobj.GetHeaderValue(14) - fmobj.GetHeaderValue(25)
    SetColor tmp, lb_midaebi
    lb_midaebi = "(" + Format(tmp, "#,###") + ")"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call Set_gRatio(1, 3, fcobj, False)
    Call Set_bid(16, fcobj)
    
    'tmp=최고가=고가
    If fcobj.GetHeaderValue(10) = fcobj.GetHeaderValue(8) Then
        lb_topday.Caption = Date '최고일자=당일
    End If
    'tmp=최저가=저가
    If fcobj.GetHeaderValue(9) = fcobj.GetHeaderValue(11) Then
        lb_bottomday.Caption = Date ''최저일자=당일
    End If
    
    cur_sb_callflag = True
End Sub
Sub data_susin()
    fmobj.SetInputValue 0, cb_Fjongmok.Text
    fmobj.BlockRequest
    fwide.SetInputValue 0, cb_Fjongmok.Text
    fwide.BlockRequest
    
    lb_curprice.Caption = Format(fmobj.GetHeaderValue(71), "##0.00")
    SetColor fmobj.GetHeaderValue(77), lb_curdaebi
    lb_curdaebi.Caption = "(" + Format(fmobj.GetHeaderValue(77), "#0.00") + ")" '현재가(전일대비)
    
    lb_vol.Caption = Format(fmobj.GetHeaderValue(75), "#,###") '거래량
    lb_mi.Caption = Format(fmobj.GetHeaderValue(80), "#,###") '미결제약정

    lb_kospi.Caption = Format(fmobj.GetHeaderValue(89), "##0.00") '코스피
    
    SetColor fmobj.GetHeaderValue(91), lb_kospidaebi
    lb_kospidaebi = "(" + Format(fmobj.GetHeaderValue(91), "##0.00") + ")"
    
    SetColor fmobj.GetHeaderValue(90), lb_basis
    lb_basis.Caption = Format(fmobj.GetHeaderValue(90), "##0.00") '베이시스
    
    lb_iron.Caption = Format(fmobj.GetHeaderValue(88), "##0.00") '이론가
    lb_geonmi.Caption = Format(fmobj.GetHeaderValue(25), "#,###") '전일미결재
    
    lb_gijun(0).Caption = Format(fmobj.GetHeaderValue(13), "##0.00") + "  "  '기준가
    lb_gijun(1).Caption = Format(fmobj.GetHeaderValue(13), "##0.00") '기준가
   
    lb_topday.Caption = Format(fmobj.GetHeaderValue(26), "####/##/##") '최고일자
    lb_top.Caption = Format(fmobj.GetHeaderValue(27), "##0.00") + "  "  '최고가
    lb_bottomday.Caption = Format(fmobj.GetHeaderValue(28), "####/##/##") '최저일자
    lb_bottom.Caption = Format(fmobj.GetHeaderValue(29), "##0.00") + "  "  '최저가
    lb_endday.Caption = Format(fmobj.GetHeaderValue(9), "####/##/##") '최종거래일
    lb_days.Caption = "(" + Str(fmobj.GetHeaderValue(8)) + ")"        '잔존일수"
    lb_murisk = Format(fmobj.GetHeaderValue(16), "##0.00") + " %     " '무위험
    
    Call Set_start_high_low_price(72, fmobj)
    Call Set_daebi(71, 29, fmobj, lb_bottomdaebi, False)
    Call Set_daebi(71, 27, fmobj, lb_topdaebi, False)
    Call Set_daebi(74, 13, fmobj, lb_lowdaebi, True)
    Call Set_daebi(73, 13, fmobj, lb_highdaebi, True)
    Call Set_daebi(72, 13, fmobj, lb_startdaebi, True)
''''''''''''''''''''''''''''''''''
    tmp = fmobj.GetHeaderValue(80) - fmobj.GetHeaderValue(25)
    SetColor tmp, lb_midaebi
    lb_midaebi = "(" + Format(tmp, "#,###") + ")"
''''''''''''''''''''''''''''''''''''''''''''''
    Call Set_gRatio(71, 88, fmobj, False)
    Call Set_bid(37, fmobj)
    
    lb_sangprice = Format(fwide.GetHeaderValue(2), "##0.00")
    lb_haprice = Format(fwide.GetHeaderValue(3), "##0.00")
    lb_sang = Format(fwide.GetHeaderValue(4), "##0.00")
    lb_ha = Format(fwide.GetHeaderValue(5), "##0.00")
   '/////////////////////////////////////////////////////////////////////
    Call Call_sb_Object
End Sub
Private Sub cb_Fjongmok_Change()
    lb_jongmok.Caption = futurelist.CodeToName(cb_Fjongmok.Text)
End Sub
Private Sub cb_Fjongmok_Click()
    lb_jongmok.Caption = futurelist.CodeToName(cb_Fjongmok.Text)
    
    cur_sb_callflag = False
    Call data_susin
End Sub
Private Sub Cmd_end_Click()
Unload Me
End Sub


