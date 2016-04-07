VERSION 5.00
Begin VB.Form tr9721 
   BorderStyle     =   1  '단일 고정
   Caption         =   "옵션 현재가"
   ClientHeight    =   7260
   ClientLeft      =   4020
   ClientTop       =   1650
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   5955
   Begin VB.CommandButton Command1 
      Caption         =   "검색"
      Height          =   350
      Left            =   1560
      TabIndex        =   105
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txt_OCode 
      Height          =   350
      Left            =   600
      TabIndex        =   104
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5775
      Begin VB.Label Label5 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2655
         TabIndex        =   128
         Top             =   3480
         Width           =   300
      End
      Begin VB.Label lb_updown 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   4720
         TabIndex        =   127
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lb_updown 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   4720
         TabIndex        =   126
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lb_updown 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   125
         Top             =   2760
         Width           =   520
      End
      Begin VB.Label lb_updown 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   124
         Top             =   2040
         Width           =   520
      End
      Begin VB.Label Label10 
         BackColor       =   &H00EBEBEB&
         Caption         =   "(%)"
         Height          =   255
         Left            =   5300
         TabIndex        =   115
         Top             =   3480
         Width           =   400
      End
      Begin VB.Label Label8 
         BackColor       =   &H00EBEBEB&
         Caption         =   "(%)"
         Height          =   255
         Left            =   5300
         TabIndex        =   114
         Top             =   3120
         Width           =   400
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  변 동 성"
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
         Left            =   3000
         TabIndex        =   113
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lb_vRatio 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4200
         TabIndex        =   112
         Top             =   3120
         Width           =   1100
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "행사가격"
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
         TabIndex        =   111
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lb_Hprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   110
         Top             =   2400
         Width           =   840
      End
      Begin VB.Label Label4 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   109
         Top             =   2400
         Width           =   795
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
         TabIndex        =   108
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lb_iron 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   107
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label lb_irondaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   106
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label lb_days 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2280
         TabIndex        =   103
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label lb_sellprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   102
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lb_buyprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   101
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   100
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lb_kospidaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   99
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label lb_midaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   98
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lb_curdaebi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   2160
         TabIndex        =   97
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lb_endday 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   90
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lb_top 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4200
         TabIndex        =   89
         Top             =   1680
         Width           =   500
      End
      Begin VB.Label lb_low 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4200
         TabIndex        =   88
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lb_high 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4200
         TabIndex        =   87
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lb_start 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4200
         TabIndex        =   86
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lb_murisk 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4200
         TabIndex        =   85
         Top             =   3480
         Width           =   1100
      End
      Begin VB.Label lb_bottomday 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4700
         TabIndex        =   84
         Top             =   2400
         Width           =   1030
      End
      Begin VB.Label lb_bottom 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   4200
         TabIndex        =   83
         Top             =   2400
         Width           =   500
      End
      Begin VB.Label lb_topday 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4700
         TabIndex        =   82
         Top             =   1680
         Width           =   1030
      End
      Begin VB.Label lb_gijun 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   81
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lb_money 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   80
         Top             =   3480
         Width           =   1340
      End
      Begin VB.Label lb_kospi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   79
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lb_mi 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   78
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lb_vol 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   77
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lb_curprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00EBEBEB&
         Height          =   255
         Left            =   1320
         TabIndex        =   76
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lb_buyprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   46
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label lb_sellprice 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   40
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
         Left            =   120
         TabIndex        =   20
         Top             =   3120
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
         TabIndex        =   19
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "등 락(률)"
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "등 락(률)"
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "거래대금"
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
         TabIndex        =   10
         Top             =   3480
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
         TabIndex        =   9
         Top             =   2760
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Cmd_end 
      Caption         =   "종  료"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lb_fcurdaebi 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   5280
      TabIndex        =   123
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label lb_Theta 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   1035
      TabIndex        =   122
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Theta"
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
      Index           =   40
      Left            =   120
      TabIndex        =   121
      Top             =   5040
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Vega"
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
      Index           =   39
      Left            =   2055
      TabIndex        =   120
      Top             =   5040
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Rho"
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
      Index           =   38
      Left            =   3960
      TabIndex        =   119
      Top             =   5040
      Width           =   900
   End
   Begin VB.Label lb_Vega 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   2955
      TabIndex        =   118
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label lb_Rho 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   4840
      TabIndex        =   117
      Top             =   5040
      Width           =   900
   End
   Begin VB.Label lb_IV 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   1030
      TabIndex        =   116
      Top             =   4800
      Width           =   1005
   End
   Begin VB.Label lb_compare 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   2400
      TabIndex        =   96
      Top             =   6915
      Width           =   1935
   End
   Begin VB.Label lb_Gamma 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   4840
      TabIndex        =   95
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label lb_fcurrent 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   4655
      TabIndex        =   94
      Top             =   4440
      Width           =   650
   End
   Begin VB.Label lb_Delta 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   2950
      TabIndex        =   93
      Top             =   4800
      Width           =   1005
   End
   Begin VB.Label lb_bottomHprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   2640
      TabIndex        =   92
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label lb_topHprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00EBEBEB&
      Height          =   255
      Left            =   1080
      TabIndex        =   91
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   75
      Top             =   6915
      Width           =   690
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00C4C4C4&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   74
      Top             =   6630
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00CACACA&
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   73
      Top             =   6390
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00D6D6D6&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   72
      Top             =   6150
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00EAEAEA&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   71
      Top             =   5910
      Width           =   615
   End
   Begin VB.Label lb_buygeonsu 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   70
      Top             =   5670
      Width           =   615
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   69
      Top             =   6915
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C4C4C4&
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   68
      Top             =   6630
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00CACACA&
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   67
      Top             =   6390
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00D6D6D6&
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   66
      Top             =   6150
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00EAEAEA&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   65
      Top             =   5910
      Width           =   855
   End
   Begin VB.Label lb_buysu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   64
      Top             =   5670
      Width           =   855
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   63
      Top             =   6915
      Width           =   690
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   62
      Top             =   6915
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C4C4C4&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   61
      Top             =   6630
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00CACACA&
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   60
      Top             =   6390
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00D6D6D6&
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   59
      Top             =   6150
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00EAEAEA&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   58
      Top             =   5910
      Width           =   855
   End
   Begin VB.Label lb_sellsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   57
      Top             =   5670
      Width           =   855
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C4C4C4&
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   56
      Top             =   6630
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00CACACA&
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   55
      Top             =   6390
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00D6D6D6&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   54
      Top             =   6150
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00EAEAEA&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   53
      Top             =   5910
      Width           =   615
   End
   Begin VB.Label lb_sellgeonsu 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   52
      Top             =   5670
      Width           =   615
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   51
      Top             =   6630
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   50
      Top             =   6390
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   49
      Top             =   6150
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   48
      Top             =   5910
      Width           =   975
   End
   Begin VB.Label lb_buyprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   47
      Top             =   5670
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   45
      Top             =   6630
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   44
      Top             =   6390
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   43
      Top             =   6150
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   42
      Top             =   5910
      Width           =   975
   End
   Begin VB.Label lb_sellprice 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   41
      Top             =   5670
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
      TabIndex        =   39
      Top             =   5400
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
      TabIndex        =   38
      Top             =   5400
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
      TabIndex        =   37
      Top             =   5400
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
      TabIndex        =   36
      Top             =   5400
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
      TabIndex        =   35
      Top             =   5400
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
      TabIndex        =   34
      Top             =   5400
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
      TabIndex        =   33
      Top             =   6630
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
      TabIndex        =   32
      Top             =   6390
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
      TabIndex        =   31
      Top             =   6150
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
      TabIndex        =   30
      Top             =   5910
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
      TabIndex        =   29
      Top             =   5670
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
      TabIndex        =   28
      Top             =   5400
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
      TabIndex        =   27
      Top             =   6915
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Gamma"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "선물최근월물"
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
      Left            =   3360
      TabIndex        =   25
      Top             =   4440
      Width           =   1280
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Delta"
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
      Left            =   2050
      TabIndex        =   24
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "최저호가 "
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
      Left            =   1800
      TabIndex        =   23
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "I.V"
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
      TabIndex        =   22
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0C0&
      Caption         =   "최고호가"
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
      TabIndex        =   21
      Top             =   4440
      Width           =   975
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
      Height          =   345
      Left            =   2250
      TabIndex        =   1
      Top             =   120
      Width           =   1455
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
Attribute VB_Name = "tr9721"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public futurelist As CpFutureCode
Public WithEvents fk200obj As FutureK200
Attribute fk200obj.VB_VarHelpID = -1
Public WithEvents fiobj As FutureIndexi
Attribute fiobj.VB_VarHelpID = -1

Public WithEvents fcobj As FutureCurr
Attribute fcobj.VB_VarHelpID = -1
Public fmobj As FutureMst
Attribute fmobj.VB_VarHelpID = -1

Public WithEvents ocobj As OptionCur
Attribute ocobj.VB_VarHelpID = -1
Public WithEvents ogobj As OptionGreek
Attribute ogobj.VB_VarHelpID = -1

Public omobj As OptionMst

Public cur_sb_callflag As Boolean
Private Function truncate(b, c As Integer)
        Dim a As Double
        Dim sTmp As String
        Dim itmp As Integer
    
 '   sTmp = CStr(b) '90->9로 됨.
     sTmp = b
    itmp = InStr(sTmp, ".")
    If itmp = 0 Then sTmp = sTmp + ".0000"
    
'    If itmp > 0 Then
        sTmp = Mid(sTmp, 1, itmp + c)
        a = CDbl(sTmp)
        truncate = a
 '   Else
    '    MsgBox "소숫점 없음", vbInformation, "알림"
  '  End If
End Function
Sub Object_init()
    Set futurelist = New CpFutureCode
    Set fk200obj = New FutureK200
    Set fiobj = New FutureIndexi
    
    Set fcobj = New FutureCurr
    Set fmobj = New FutureMst
    
    Set ocobj = New OptionCur
    Set ogobj = New OptionGreek
    
    Set omobj = New OptionMst
End Sub
Sub Call_sb_Object()
    fk200obj.Unsubscribe
    fk200obj.SetInputValue 0, "00800"
    fk200obj.SubscribeLatest
    
    fiobj.Unsubscribe
    fiobj.SetInputValue 0, "00800"
    fiobj.SubscribeLatest
    
    fcobj.Unsubscribe
    fcobj.SetInputValue 0, futurelist.GetData(0, 0)
    fcobj.SubscribeLatest
    
    ocobj.Unsubscribe
    ocobj.SetInputValue 0, txt_OCode.Text
    ocobj.SubscribeLatest

    ogobj.Unsubscribe
    ogobj.SetInputValue 0, txt_OCode.Text
    ogobj.SubscribeLatest
End Sub
Sub Set_bid(startindex, obj As Object)
    Set_Label obj.GetHeaderValue(startindex), "##0.00", lb_sellprice(0), False, False, obj.GetHeaderValue(startindex + 5)
    For i = 1 To 5
        Set_Label obj.GetHeaderValue(startindex + i - 1), "##0.00", lb_sellprice(i), False, False, obj.GetHeaderValue(startindex + 5 + i - 1)
    Next
    For i = 1 To 6
        Set_Label obj.GetHeaderValue(startindex + 5 + i - 1), "#,###  ", lb_sellsu(i), False, False, obj.GetHeaderValue(startindex + 5 + i - 1)
    Next
    For i = 1 To 6
        Set_Label obj.GetHeaderValue(startindex + 37 + i - 1), "#,###", lb_sellgeonsu(i), False, True, obj.GetHeaderValue(startindex + 37 + i - 1)
    Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set_Label obj.GetHeaderValue(startindex + 11), "##0.00", lb_buyprice(0), False, False, Format(obj.GetHeaderValue(startindex + 11 + 1 - 1))
    For i = 1 To 5
        Set_Label obj.GetHeaderValue(startindex + 11 + i - 1), "##0.00", lb_buyprice(i), False, False, Format(obj.GetHeaderValue(startindex + 11 + i - 1))
    Next
    For i = 1 To 6
        Set_Label obj.GetHeaderValue(startindex + 16 + i - 1), "#,###  ", lb_buysu(i), False, False, Format(obj.GetHeaderValue(startindex + 16 + i - 1))
    Next
    For i = 1 To 6
        Set_Label obj.GetHeaderValue(startindex + 43 + i - 1), "#,###", lb_buygeonsu(i), False, True, obj.GetHeaderValue(startindex + 43 + i - 1)
    Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    tmp = obj.GetHeaderValue(startindex + 21) - obj.GetHeaderValue(startindex + 10)
    Set_minus_convert_Label tmp, "#,###", lb_compare, True, False, tmp
End Sub

Sub Set_omst_bid(startindex, obj As Object)
    Set_Label obj.GetHeaderValue(startindex), "##0.00", lb_sellprice(0), False, False, obj.GetHeaderValue(startindex + i * 2)
    For i = 1 To 5
        Set_Label obj.GetHeaderValue(startindex + (i - 1) * 4), "##0.00", lb_sellprice(i), False, False, obj.GetHeaderValue(startindex + (i - 1) * 4)
    Next
    For i = 1 To 5
        Set_Label obj.GetHeaderValue(startindex + (i - 1) * 4 + 2), "#,###  ", lb_sellsu(i), False, False, obj.GetHeaderValue(startindex + (i - 1) * 4 + 2)
    Next
    For i = 1 To 6
        Set_Label obj.GetHeaderValue(startindex + 21 + i), "#,###", lb_sellgeonsu(i), False, True, obj.GetHeaderValue(startindex + 21 + i)
    Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set_Label obj.GetHeaderValue(startindex + 1), "##0.00", lb_buyprice(0), False, False, Format(obj.GetHeaderValue(startindex + 1 + 2))
    For i = 1 To 5
        Set_Label obj.GetHeaderValue(startindex + (i - 1) * 4 + 1), "##0.00", lb_buyprice(i), False, False, Format(obj.GetHeaderValue(startindex + (i - 1) * 4 + 1))
    Next
    For i = 1 To 5
        Set_Label obj.GetHeaderValue(startindex + (i - 1) * 4 + 3), "#,###  ", lb_buysu(i), False, False, Format(obj.GetHeaderValue(startindex + (i - 1) * 4 + 3))
    Next
    For i = 1 To 6
        Set_Label obj.GetHeaderValue(startindex + 28 + i - 1), "#,###", lb_buygeonsu(i), False, True, obj.GetHeaderValue(startindex + 28 + i - 1)
    Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set_Label obj.GetHeaderValue(startindex + 20), "#,###  ", lb_sellsu(6), False, False, obj.GetHeaderValue(startindex + 20)
    Set_Label obj.GetHeaderValue(startindex + 21), "#,###  ", lb_buysu(6), False, False, obj.GetHeaderValue(startindex + 21)
    tmp = obj.GetHeaderValue(startindex + 21) - obj.GetHeaderValue(startindex + 20)
    Set_minus_convert_Label tmp, "#,###", lb_compare, True, False, tmp
End Sub
Private Sub Command1_Click()
    optionsearch.Show
End Sub




Private Sub fcobj_Received()
    Set_Label fcobj.GetHeaderValue(1), "##0.00", lb_fcurrent, False, False, fcobj.GetHeaderValue(1)
    Set_Label fcobj.GetHeaderValue(2), "##0.00", lb_fcurdaebi, True, True, fcobj.GetHeaderValue(1)
End Sub

Private Sub Form_Load()
    Call Object_init
    optionsearch.set_OCode
    cur_sb_callflag = False
    optionsearch.Show
End Sub
Private Sub fiobj_Received()
    Set_Label fiobj.GetHeaderValue(2), "##0.00", lb_kospi, False, False, fiobj.GetHeaderValue(2)
    Set_Label fiobj.GetHeaderValue(4), "##0.00", lb_kospidaebi, True, True, fiobj.GetHeaderValue(2)
End Sub

Private Sub fk200obj_Received()
    Set_Label fk200obj.GetHeaderValue(1), "##0.00", lb_kospi, False, False, fk200obj.GetHeaderValue(1)
    tmp = (fk200obj.GetHeaderValue(1) - fmobj.GetHeaderValue(89)) + fmobj.GetHeaderValue(91)
    Set_Label tmp, "##0.00", lb_kospidaebi, True, True, fk200obj.GetHeaderValue(1)
End Sub
Sub data_susin()
    omobj.SetInputValue 0, txt_OCode.Text
    omobj.BlockRequest
    fmobj.SetInputValue 0, futurelist.GetData(0, 0)
    fmobj.BlockRequest
    
    Set_Label omobj.GetHeaderValue(93), "##0.00", lb_curprice, False, False, omobj.GetHeaderValue(93)
    tmp = omobj.GetHeaderValue(93) - omobj.GetHeaderValue(51) '27
    Set_Label tmp, "##0.00", lb_curdaebi, True, True, omobj.GetHeaderValue(93)
    Set_Label fmobj.GetHeaderValue(71), "##0.00", lb_fcurrent, False, False, fmobj.GetHeaderValue(71)
    Set_Label fmobj.GetHeaderValue(77), "##0.00", lb_fcurdaebi, True, True, fmobj.GetHeaderValue(71)
    
    Set_Label omobj.GetHeaderValue(97), "#,###", lb_vol, False, False, omobj.GetHeaderValue(97)
    Set_Label omobj.GetHeaderValue(99), "#,##0", lb_mi, False, False, omobj.GetHeaderValue(99)
    tmp = omobj.GetHeaderValue(99) - omobj.GetHeaderValue(37)
    Set_Label tmp, "#,##0", lb_midaebi, True, True, omobj.GetHeaderValue(99)
   
    Set_Label omobj.GetHeaderValue(114), "##0.00", lb_iron, False, False, omobj.GetHeaderValue(114)
    tmp = omobj.GetHeaderValue(93) - omobj.GetHeaderValue(114)
    Set_minus_convert_Label tmp, "##0.00", lb_irondaebi, True, True, omobj.GetHeaderValue(114)
    Set_Label omobj.GetHeaderValue(6), "##0.00", lb_Hprice, False, False, omobj.GetHeaderValue(6)
    Set_Label fmobj.GetHeaderValue(89), "##0.00", lb_kospi, False, False, fmobj.GetHeaderValue(89)
    Set_Label fmobj.GetHeaderValue(91), "##0.00", lb_kospidaebi, True, True, fmobj.GetHeaderValue(89)
    
    Set_Label omobj.GetHeaderValue(18), "####/##/##", lb_endday, False, False, omobj.GetHeaderValue(18)
    Set_Label omobj.GetHeaderValue(13), "", lb_days, False, True, omobj.GetHeaderValue(13)
    
    Set_Label omobj.GetHeaderValue(98), "#,###(백만)", lb_money, False, False, omobj.GetHeaderValue(98)
   ' '================================================
    Set_Label omobj.GetHeaderValue(51), "##0.00", lb_gijun(0), False, False, omobj.GetHeaderValue(51)
    Set_Label omobj.GetHeaderValue(94), "##0.00", lb_start, False, False, omobj.GetHeaderValue(94)
    Set_Label omobj.GetHeaderValue(95), "##0.00", lb_high, False, False, omobj.GetHeaderValue(95)
    Set_Label omobj.GetHeaderValue(96), "##0.00", lb_low, False, False, omobj.GetHeaderValue(96)
    
    Set_Label omobj.GetHeaderValue(40), "####/##/##", lb_topday, False, True, omobj.GetHeaderValue(40)
    Set_Label omobj.GetHeaderValue(41), "##0.00", lb_top, False, False, omobj.GetHeaderValue(41)
    Set_Label omobj.GetHeaderValue(42), "####/##/##", lb_bottomday, False, True, omobj.GetHeaderValue(42)
    Set_Label omobj.GetHeaderValue(43), "##0.00", lb_bottom, False, False, omobj.GetHeaderValue(43)
    
    If omobj.GetHeaderValue(41) Then
        tmp = omobj.GetHeaderValue(93) - omobj.GetHeaderValue(41)
        Set_Label tmp, "##0.00", lb_updown(0), False, False, omobj.GetHeaderValue(41)
'''''''''''''''등락률
        tmp = truncate(tmp / omobj.GetHeaderValue(41) * 100, 2)
        Set_minus_convert_Label tmp, "0.00", lb_updown(2), False, False, omobj.GetHeaderValue(41)
        lb_updown(2) = "(" + lb_updown(2) + "%)"
    Else
        lb_updown(0) = ""
        lb_updown(2) = ""
    End If
    
    If omobj.GetHeaderValue(43) Then
        tmp = omobj.GetHeaderValue(93) - omobj.GetHeaderValue(43)
        Set_Label tmp, "##0.00", lb_updown(1), False, False, omobj.GetHeaderValue(43)
        tmp = truncate(tmp / omobj.GetHeaderValue(43) * 100, 2)
        Set_minus_convert_Label tmp, "0.00", lb_updown(3), False, False, omobj.GetHeaderValue(43)
        lb_updown(3) = "(" + lb_updown(3) + "%)"
    Else
        lb_updown(1) = ""
        lb_updown(3) = ""
    End If
    
    Set_Label omobj.GetHeaderValue(115), "##0.00", lb_vRatio, False, False, omobj.GetHeaderValue(115)
    Set_Label omobj.GetHeaderValue(36) / 10, "#0.00", lb_murisk, False, False, omobj.GetHeaderValue(36)
    '''''''''''''''''''''''''''''''''''''''''''''
    Set_Label omobj.GetHeaderValue(31), "##0.00", lb_topHprice, False, False, omobj.GetHeaderValue(31)
    Set_Label omobj.GetHeaderValue(32), "##0.00", lb_bottomHprice, False, False, omobj.GetHeaderValue(32)
    Set_Label omobj.GetHeaderValue(108), "##0.00", lb_IV, False, False, omobj.GetHeaderValue(108)
    Set_Label omobj.GetHeaderValue(109), "##0.00", lb_Delta, False, False, omobj.GetHeaderValue(109)
    Set_Label omobj.GetHeaderValue(110), "##0.00", lb_Gamma, False, False, omobj.GetHeaderValue(110)
    Set_Label omobj.GetHeaderValue(111), "##0.0000", lb_Theta, False, False, omobj.GetHeaderValue(111)
    Set_Label omobj.GetHeaderValue(112), "##0.0000", lb_Vega, False, False, omobj.GetHeaderValue(112)
    Set_Label omobj.GetHeaderValue(113), "##0.0000", lb_Rho, False, False, omobj.GetHeaderValue(113)
    Call Set_omst_bid(58, omobj)
    Call Call_sb_Object
End Sub
Private Sub Cmd_end_Click()
Unload optionsearch
Unload Me
End Sub
'값, 자리표시str,obj,()flag,refvalue
Sub Set_Label(value, str, obj As Object, color_flg As Boolean, flg As Boolean, ref_value)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If ref_value = 0 Then
    obj = ""
    Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If color_flg Then
    SetColor value, obj
Else
    obj.ForeColor = &H0
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'format에서 소숫점은 반올림이 된다.
If flg Then '()표시이면
        obj = "(" + Format(value, str) + ")"
Else  '()표시가 아니면
    obj.Caption = Format(value, str)
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Sub Set_minus_convert_Label(value, str, obj As Object, color_flg As Boolean, flg As Boolean, ref_value)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If ref_value = 0 Then
    obj = ""
    Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If color_flg Then
    SetColor value, obj
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If value < 0 Then value = -value
If flg Then '()표시이면
        obj = "(" + Format(value, str) + ")"
Else  '()표시가 아니면
    obj.Caption = Format(value, str)
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Sub SetColor(value, obj As Object) '음수값이 나오는 경우 음수값을 표시하지 않고 색깔만 표시
    If (value > 0) Then
        obj.ForeColor = RGB(255, 0, 0)
        obj = FormatNumber(value, 0)
    ElseIf (value < 0) Then
        obj.ForeColor = RGB(0, 0, 255)
        obj = FormatNumber(-value, 0)
    Else
        obj.ForeColor = &H0
        obj = FormatNumber(value, 0)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload optionsearch
End Sub

Private Sub ocobj_Received()
    Set_Label ocobj.GetHeaderValue(24), "##0.00", lb_curprice, False, False, ocobj.GetHeaderValue(24)
    tmp = ocobj.GetHeaderValue(24) - omobj.GetHeaderValue(27)
    Set_Label tmp, "##0.00", lb_curdaebi, True, True, ocobj.GetHeaderValue(24)
    
    Set_Label ocobj.GetHeaderValue(29), "#,###", lb_vol, False, False, ocobj.GetHeaderValue(29)
    Set_Label ocobj.GetHeaderValue(38), "#,##0", lb_mi, False, False, ocobj.GetHeaderValue(38)
    tmp = ocobj.GetHeaderValue(38) - omobj.GetHeaderValue(37)
    Set_Label tmp, "#,##0", lb_midaebi, True, True, ocobj.GetHeaderValue(38)
   
    Set_Label ocobj.GetHeaderValue(31), "##0.00", lb_iron, False, False, ocobj.GetHeaderValue(31)
    tmp = ocobj.GetHeaderValue(24) - ocobj.GetHeaderValue(31)
    Set_minus_convert_Label tmp, "##0.00", lb_irondaebi, True, True, ocobj.GetHeaderValue(24)
    Set_Label ocobj.GetHeaderValue(52), "##0.00", lb_Hprice, False, False, ocobj.GetHeaderValue(52)
       
    Set_Label ocobj.GetHeaderValue(30), "#,###(백만)", lb_money, False, False, ocobj.GetHeaderValue(30)

    Set_Label ocobj.GetHeaderValue(26), "##0.00", lb_start, False, False, ocobj.GetHeaderValue(26)
    Set_Label ocobj.GetHeaderValue(27), "##0.00", lb_high, False, False, ocobj.GetHeaderValue(27)
    Set_Label ocobj.GetHeaderValue(28), "##0.00", lb_low, False, False, ocobj.GetHeaderValue(28)
   ' '비교해야함....================================================
    tmp1 = omobj.GetHeaderValue(41)
    If ocobj.GetHeaderValue(27) > omobj.GetHeaderValue(41) Then
        Set_Label Date, "", lb_topday, False, True, omobj.GetHeaderValue(42)
        Set_Label ocobj.GetHeaderValue(27), "##0.00", lb_top, False, False, omobj.GetHeaderValue(41)
        tmp1 = ocobj.GetHeaderValue(27)
    End If
    tmp2 = omobj.GetHeaderValue(43)
    If ocobj.GetHeaderValue(28) < omobj.GetHeaderValue(43) Then
        Set_Label Date, "", lb_bottomday, False, True, omobj.GetHeaderValue(44)
        Set_Label ocobj.GetHeaderValue(28), "##0.00", lb_bottom, False, False, omobj.GetHeaderValue(43)
        tmp2 = ocobj.GetHeaderValue(28)
    End If
        
    If omobj.GetHeaderValue(41) Then
        tmp = ocobj.GetHeaderValue(24) - tmp1
        Set_Label tmp, "##0.00", lb_updown(0), False, False, ocobj.GetHeaderValue(41)
        tmp = truncate(tmp / omobj.GetHeaderValue(41) * 100, 2)
        Set_minus_convert_Label tmp, "0.00", lb_updown(2), False, False, omobj.GetHeaderValue(41)
        lb_updown(2) = "(" + lb_updown(2) + "%)"
    Else
        lb_updown(0) = ""
        lb_updown(2) = ""
    End If
    
    If omobj.GetHeaderValue(43) Then
        tmp = ocobj.GetHeaderValue(24) - tmp2
        Set_Label tmp, "##0.00", lb_updown(1), False, False, ocobj.GetHeaderValue(43)
        tmp = truncate(tmp / omobj.GetHeaderValue(43) * 100, 2)
        Set_minus_convert_Label tmp, "0.00", lb_updown(3), False, False, omobj.GetHeaderValue(43)
        lb_updown(3) = "(" + lb_updown(3) + "%)"
    Else
        lb_updown(1) = ""
        lb_updown(3) = ""
    End If
 ' '================================================
    Set_Label omobj.GetHeaderValue(115), "##0.00", lb_vRatio, False, False, omobj.GetHeaderValue(115)
    Set_Label omobj.GetHeaderValue(36) / 10, "#0.00", lb_murisk, False, False, omobj.GetHeaderValue(36)
    '''''''''''''''''''''''''''''''''''''''''''''
    Set_Label omobj.GetHeaderValue(31), "##0.00", lb_topHprice, False, False, omobj.GetHeaderValue(31)
    Set_Label omobj.GetHeaderValue(32), "##0.00", lb_bottomHprice, False, False, omobj.GetHeaderValue(32)
    Set_Label ocobj.GetHeaderValue(32), "##0.00", lb_IV, False, False, ocobj.GetHeaderValue(32)
    Set_Label ocobj.GetHeaderValue(33), "##0.00", lb_Delta, False, False, ocobj.GetHeaderValue(33)
    Set_Label ocobj.GetHeaderValue(34), "##0.00", lb_Gamma, False, False, ocobj.GetHeaderValue(34)
    Set_Label ocobj.GetHeaderValue(35), "##0.0000", lb_Theta, False, False, ocobj.GetHeaderValue(35)
    Set_Label ocobj.GetHeaderValue(36), "##0.0000", lb_Vega, False, False, ocobj.GetHeaderValue(36)
    Set_Label ocobj.GetHeaderValue(37), "##0.0000", lb_Rho, False, False, ocobj.GetHeaderValue(37)
    Call Set_bid(2, ocobj)
End Sub

Private Sub ogobj_Received()
    Set_Label ogobj.GetHeaderValue(1), "##0.00", lb_IV, False, False, ogobj.GetHeaderValue(1)
    Set_Label ogobj.GetHeaderValue(2), "##0.00", lb_Delta, False, False, ogobj.GetHeaderValue(2)
    Set_Label ogobj.GetHeaderValue(3), "##0.00", lb_Gamma, False, False, ogobj.GetHeaderValue(3)
    Set_Label ogobj.GetHeaderValue(4), "##0.0000", lb_Theta, False, False, ogobj.GetHeaderValue(4)
    Set_Label ogobj.GetHeaderValue(5), "##0.0000", lb_Vega, False, False, ogobj.GetHeaderValue(5)
    Set_Label ogobj.GetHeaderValue(6), "##0.0000", lb_Rho, False, False, ogobj.GetHeaderValue(6)
    
    Set_Label ogobj.GetHeaderValue(7), "##0.00", lb_iron, False, False, ogobj.GetHeaderValue(7)
    tmp = ogobj.GetHeaderValue(10) - ogobj.GetHeaderValue(7)
    Set_minus_convert_Label tmp, "##0.00", lb_irondaebi, True, True, ogobj.GetHeaderValue(7)
End Sub
