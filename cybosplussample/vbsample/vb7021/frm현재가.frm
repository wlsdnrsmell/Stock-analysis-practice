VERSION 5.00
Begin VB.Form frm���簡 
   BackColor       =   &H00FFFFFF&
   Caption         =   "VB���簡"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdsetup 
      BackColor       =   &H00FFC0FF&
      Caption         =   "����"
      Height          =   375
      Left            =   9000
      Style           =   1  '�׷���
      TabIndex        =   139
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmd��ȸ 
      Caption         =   "��ȸ"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtJongMok 
      Height          =   375
      Left            =   0
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "003540"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblbprice 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   6000
      TabIndex        =   138
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label lbltprice 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H000000FF&
      Height          =   200
      Left            =   6000
      TabIndex        =   137
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "���Ѱ�"
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   7080
      TabIndex        =   136
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  '������ ����
      BackColor       =   &H00E0E0E0&
      Caption         =   "���Ѱ�"
      ForeColor       =   &H000000FF&
      Height          =   200
      Left            =   4920
      TabIndex        =   135
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblfinishmonth 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   3840
      TabIndex        =   134
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblrightvalue 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   9360
      TabIndex        =   133
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblnominal 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   6600
      TabIndex        =   132
      Top             =   120
      Width           =   1190
   End
   Begin VB.Label lblgubun 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   3000
      TabIndex        =   131
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblrightgubun 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   8640
      TabIndex        =   130
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblgiupgubun 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   7800
      TabIndex        =   129
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblhiregubun 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   4800
      TabIndex        =   128
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lbldown 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   1560
      TabIndex        =   127
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lbltop 
      BackStyle       =   0  '����
      Height          =   255
      Left            =   1560
      TabIndex        =   126
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lbl23 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   125
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label lbl23 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   124
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label lbl23 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   123
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lbl23 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   122
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label lbl23 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   121
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lbl22 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   120
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label lbl22 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   119
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lbl22 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   118
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lbl22 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   117
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lbl22 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   116
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lbl21 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   115
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label lbl21 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   114
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label lbl21 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   113
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lbl21 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   112
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label lbl21 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   111
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lbl20 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   110
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lbl20 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   109
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label lbl20 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   108
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lbl20 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   107
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lbl20 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   106
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lbl18 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  '���� ����
      Caption         =   "�ż�����"
      Height          =   255
      Left            =   2880
      TabIndex        =   105
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label lbll17 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '���� ����
      Caption         =   "�ŵ�����"
      Height          =   255
      Left            =   840
      TabIndex        =   104
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label lbl16 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "52�� ����"
      Height          =   255
      Left            =   240
      TabIndex        =   103
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lbl13 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "52�� �ְ�"
      Height          =   255
      Left            =   240
      TabIndex        =   102
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   6000
      TabIndex        =   101
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   100
      Top             =   2960
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   99
      Top             =   2680
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   3
      Left            =   6000
      TabIndex        =   98
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   4
      Left            =   6000
      TabIndex        =   97
      Top             =   2120
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   5
      Left            =   6000
      TabIndex        =   96
      Top             =   1840
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   6
      Left            =   6000
      TabIndex        =   95
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   6000
      TabIndex        =   94
      Top             =   1280
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   8
      Left            =   6000
      TabIndex        =   93
      Top             =   1000
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      Height          =   195
      Index           =   9
      Left            =   6000
      TabIndex        =   92
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   49
      Left            =   7080
      TabIndex        =   91
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   59
      Left            =   8160
      TabIndex        =   90
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   58
      Left            =   8160
      TabIndex        =   89
      Top             =   6048
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   57
      Left            =   8160
      TabIndex        =   88
      Top             =   5742
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   56
      Left            =   8160
      TabIndex        =   87
      Top             =   5436
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   55
      Left            =   8160
      TabIndex        =   86
      Top             =   5130
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   48
      Left            =   7080
      TabIndex        =   85
      Top             =   6048
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   47
      Left            =   7080
      TabIndex        =   84
      Top             =   5742
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   46
      Left            =   7080
      TabIndex        =   83
      Top             =   5436
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   45
      Left            =   7080
      TabIndex        =   82
      Top             =   5130
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   39
      Left            =   6000
      TabIndex        =   81
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   38
      Left            =   6000
      TabIndex        =   80
      Top             =   6048
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   37
      Left            =   6000
      TabIndex        =   79
      Top             =   5742
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   36
      Left            =   6000
      TabIndex        =   78
      Top             =   5436
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   35
      Left            =   6000
      TabIndex        =   77
      Top             =   5130
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   29
      Left            =   3840
      TabIndex        =   76
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   28
      Left            =   3840
      TabIndex        =   75
      Top             =   1000
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   27
      Left            =   3840
      TabIndex        =   74
      Top             =   1280
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   26
      Left            =   3840
      TabIndex        =   73
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   25
      Left            =   3840
      TabIndex        =   72
      Top             =   1840
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   19
      Left            =   4920
      TabIndex        =   71
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   18
      Left            =   4920
      TabIndex        =   70
      Top             =   1000
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   17
      Left            =   4920
      TabIndex        =   69
      Top             =   1280
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   16
      Left            =   4920
      TabIndex        =   68
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   15
      Left            =   4920
      TabIndex        =   67
      Top             =   1840
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   165
      Left            =   960
      Picture         =   "frm���簡.frx":0000
      Top             =   1440
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   960
      Picture         =   "frm���簡.frx":00DA
      Top             =   1440
      Width           =   165
   End
   Begin VB.Label lbprevdaebi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   2880
      TabIndex        =   66
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����:"
      Height          =   255
      Left            =   2160
      TabIndex        =   65
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "%"
      Height          =   255
      Left            =   1920
      TabIndex        =   64
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbmedan 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   375
      Left            =   1920
      TabIndex        =   63
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Ÿż�������"
      Height          =   375
      Left            =   360
      TabIndex        =   62
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lb�ܱ��κ��� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   375
      Left            =   720
      TabIndex        =   61
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label Label11 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   375
      Left            =   120
      TabIndex        =   60
      Top             =   4200
      Width           =   555
   End
   Begin VB.Label Label19 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ð���"
      Height          =   195
      Left            =   6000
      TabIndex        =   59
      Top             =   7440
      Width           =   1005
   End
   Begin VB.Label lb�ð��ܸż��ܷ� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7080
      TabIndex        =   58
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lb�ð��ܸż��ܷ���� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   195
      Left            =   8160
      TabIndex        =   57
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lb�Ѹż��ܷ� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7080
      TabIndex        =   56
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lb�Ѹż��ܷ���� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   195
      Left            =   8160
      TabIndex        =   55
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lb�ð��ܸŵ��ܷ���� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   195
      Left            =   3840
      TabIndex        =   54
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lb�ð��ܸŵ��ܷ� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4920
      TabIndex        =   53
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lb�Ѹŵ��ܷ���� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   195
      Left            =   3600
      TabIndex        =   52
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label lb�Ѹŵ��ܷ� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   4920
      TabIndex        =   51
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lb�ð� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   195
      Left            =   6000
      TabIndex        =   50
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lb�ܱ����ѵ����� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   49
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lb�ܱ��ΰ��� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   48
      Top             =   3540
      Width           =   1035
   End
   Begin VB.Label lb�ܱ��ΰ��ɺ��� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   47
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label lb�ܱ��κ��� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   46
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label lb�ܱ����ѵ� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   45
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label Label10 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3840
      Width           =   555
   End
   Begin VB.Label Label9 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ѵ�"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label Label8 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   3540
      Width           =   555
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ܱ���"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   2940
      Width           =   735
   End
   Begin VB.Label lb����� 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   40
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   54
      Left            =   8160
      TabIndex        =   39
      Top             =   4824
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   53
      Left            =   8160
      TabIndex        =   38
      Top             =   4518
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   52
      Left            =   8160
      TabIndex        =   37
      Top             =   4212
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   51
      Left            =   8160
      TabIndex        =   36
      Top             =   3906
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   50
      Left            =   8160
      TabIndex        =   35
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   44
      Left            =   7080
      TabIndex        =   34
      Top             =   4824
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   43
      Left            =   7080
      TabIndex        =   33
      Top             =   4518
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   42
      Left            =   7080
      TabIndex        =   32
      Top             =   4212
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   41
      Left            =   7080
      TabIndex        =   31
      Top             =   3906
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   40
      Left            =   7080
      TabIndex        =   30
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   34
      Left            =   6000
      TabIndex        =   29
      Top             =   4824
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   33
      Left            =   6000
      TabIndex        =   28
      Top             =   4518
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   32
      Left            =   6000
      TabIndex        =   27
      Top             =   4212
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   31
      Left            =   6000
      TabIndex        =   26
      Top             =   3906
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   30
      Left            =   6000
      TabIndex        =   25
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   24
      Left            =   3840
      TabIndex        =   24
      Top             =   2120
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   23
      Left            =   3840
      TabIndex        =   23
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   22
      Left            =   3840
      TabIndex        =   22
      Top             =   2680
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   21
      Left            =   3840
      TabIndex        =   21
      Top             =   2960
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFE0&
      Caption         =   "0"
      Height          =   195
      Index           =   20
      Left            =   3840
      TabIndex        =   20
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   14
      Left            =   4920
      TabIndex        =   19
      Top             =   2120
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   13
      Left            =   4920
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   12
      Left            =   4920
      TabIndex        =   17
      Top             =   2680
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   11
      Left            =   4920
      TabIndex        =   16
      Top             =   2960
      Width           =   975
   End
   Begin VB.Label lbȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   195
      Index           =   10
      Left            =   4920
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lb�ŷ���� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ŷ����"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lb�ŷ��� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ŷ���"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lb������ 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lb���簡 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lb��� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lb�ż�ȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lb�ŵ�ȣ�� 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "���"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ż�ȣ��"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "���簡"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10000
      Y1              =   410
      Y2              =   410
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ŵ�ȣ��"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frm���簡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2002.1.2~1.8 modified by ldh (5��ȣ������ 10�� ȣ�� �������� ����)
'2002.1.11 by ldh
'5��/10�� ȣ�� ����,���� �ְ�/����,52�� �ְ�/���� ���� �����ϴ� �� ����
Private WithEvents smobj As StockMst
Attribute smobj.VB_VarHelpID = -1
Private WithEvents sjb2obj As StockJpbid2
Attribute sjb2obj.VB_VarHelpID = -1
Private WithEvents sjbobj As StockJpbid
Attribute sjbobj.VB_VarHelpID = -1
Private WithEvents scobj As StockCur
Attribute scobj.VB_VarHelpID = -1
Private WithEvents member1obj As StockMember1
Attribute member1obj.VB_VarHelpID = -1
Private WithEvents memberobj As StockMember
Attribute memberobj.VB_VarHelpID = -1

Private m_CodeMgr As CpCodeMgr
Private aPrevRest(44) As Long ' ����ȣ�� �ܷ� �� ����
Private bPrevRestInit As Boolean ' ����ȣ�� �ܷ� ���� ����
Private highprice As Long   '52�� �ְ�
Private lowprice As Long    '52�� ������
Private sanghanga As Long   '���Ѱ�
Private hahanga As Long     '���Ѱ�
Private gubun As Integer    '5�� ȣ��/10�� ȣ�� ����
Private gubun2 As Integer   '���� �ְ�/����,52�� �ְ�/����
Private weekobj As StockWeek
Private Sub cmdsetup_Click()
        frmsetup.Show vbModal
        If txtJongMok.Text <> "" Then
           cmd��ȸ_Click
        End If
End Sub
Private Sub Form_Load()
    Set smobj = New StockMst
    Set sjb2obj = New StockJpbid2
    Set sjbobj = New StockJpbid
    Set scobj = New StockCur
    Set weekobj = New StockWeek
    Set member1obj = New StockMember1
    Set memberobj = New StockMember
    Set m_CodeMgr = New CpCodeMgr
    
    bPrevRestInit = False
    Image1.Visible = False
    Image2.Visible = False
    Load frmsetup
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set smobj = Nothing
    Set sjb2obj = Nothing
    Set sjbobj = Nothing
    Set scobj = Nothing
    Set weekobj = Nothing
    Set member1obj = Nothing
    Set memberobj = Nothing
    Set m_CodeMgr = Nothing
    Unload frmsetup
End Sub
Private Sub cmd��ȸ_Click()
    '���а��� setting �ؼ� �־��ش�
    gubun = frmsetup.envirhoga
    gubun2 = frmsetup.enviryear
    If Left(txtJongMok.Text, 1) > "5" Then 'elw ������ 5�� �����Ѵ�.
        jcode = "J" + txtJongMok
    Else
        jcode = "A" + txtJongMok
    End If
    bPrevRestInit = False
    
    smobj.SetInputValue 0, jcode
    smobj.Request
    
    sjb2obj.SetInputValue 0, jcode
    sjb2obj.Request
    
    member1obj.SetInputValue 0, jcode
    member1obj.Request
 
    weekobj.SetInputValue 0, jcode
    weekobj.BlockRequest
    
    'by ldh ���� �߰�(2001.12.10)
    lb�ܱ��κ��� = FormatNumber(weekobj.GetDataValue(9, 0), 2) & "%"
     
    memberobj.Unsubscribe
    memberobj.SetInputValue 0, jcode
    memberobj.SubscribeLatest
    
    sjbobj.Unsubscribe
    sjbobj.SetInputValue 0, jcode
    sjbobj.SubscribeLatest

    scobj.Unsubscribe
    scobj.SetInputValue 0, jcode
    scobj.SubscribeLatest
    
    txtJongMok.SetFocus
    txtJongMok.SelStart = Len(txtJongMok) - 5
    txtJongMok.SelLength = Len(txtJongMok)
End Sub
' ����ȣ���ܷ� ����� ���ϴ� �Լ�(�ŵ�)
Function GetRestDiffer(Hoga, value)
    For i = 0 To 9
        If aPrevRest(20 + i) = Hoga Then
            GetRestDiffer = value - aPrevRest(i)
            Exit Function
        End If
    Next
    GetRestDiffer = 0
End Function
' ����ȣ���ܷ� ����� ���ϴ� �Լ�(�ż�)
Function GetRstDiffer(Hoga, value)
    For i = 0 To 9
        If aPrevRest(30 + i) = Hoga Then
            GetRstDiffer = value - aPrevRest(i + 10)
            Exit Function
        End If
    Next
    GetRstDiffer = 0
End Function
Sub SetColorOfValue(value, obj)
    If (value > 0) Then
        obj.ForeColor = RGB(255, 0, 0)
    ElseIf (value < 0) Then
        obj.ForeColor = RGB(0, 0, 255)
    Else
        obj.ForeColor = &H0
    End If
End Sub
Sub SetValueWithColor(value, obj As Object, gubun As Boolean)
If value < 0 Then
    st = Mid(CStr(value), 2)
    If gubun = True Then
    strtemp = FormatNumber(st, 0)
    Else
    strtemp = FormatNumber(st, 2)
    End If
    obj.ForeColor = RGB(0, 0, 255)
 ElseIf value > 0 Then
   If gubun = True Then
   strtemp = FormatNumber(value, 0)
   Else
   strtemp = FormatNumber(value, 2)
   End If
   obj.ForeColor = RGB(255, 0, 0)
 Else
   strtemp = 0
   obj.ForeColor = RGB(0, 0, 0)
 End If
   obj = strtemp
End Sub
'�ŵ��� ��� �ܷ� ǥ��(�Ķ���),�ż��� ��� �ܷ� ǥ��(������),0�� ��� ǥ�� �� �ϱ�(by ldh,
Sub leftright(value, obj As Object, pos As Integer)
If value = 0 Then
   strtemp = ""
Else
    If pos = 1 Then
       obj.ForeColor = RGB(0, 0, 255)
    ElseIf pos = 2 Then
       obj.ForeColor = RGB(255, 0, 0)
    End If
       strtemp = FormatNumber(value, 0)
End If
    obj = strtemp
End Sub
Private Sub member1obj_Received()
        For i = 0 To member1obj.GetHeaderValue(1) - 1
            lbl20(i).ForeColor = RGB(0, 0, 0)
            lbl22(i).ForeColor = RGB(0, 0, 0)
        Next
        '�ŵ��ŷ����� �ܱ��� ȸ���� ���Ե� ��� �ŵ�ȸ����� û��(2002.1.9 BY LDH)
        '�ż��ŷ����� �ܱ��� ȸ���� ���Ե� ��� �ż�ȸ����� ����
        For i = 0 To member1obj.GetHeaderValue(1) - 1
        medotemp = member1obj.GetDataValue(0, i)
        If IsFrn(medotemp) Then
           lbl20(i).ForeColor = RGB(0, 0, 255)
        End If
           lbl20(i) = Convert(medotemp)
        lbl21(i) = FormatNumber(member1obj.GetDataValue(2, i), 0)
        mesutemp = member1obj.GetDataValue(1, i)
        If IsFrn(mesutemp) Then
           lbl22(i).ForeColor = RGB(255, 0, 0)
        End If
        lbl22(i) = Convert(mesutemp)
        lbl23(i) = FormatNumber(member1obj.GetDataValue(3, i), 0)
        Next
End Sub
Private Sub memberobj_Received()

        For i = 0 To memberobj.GetHeaderValue(1) - 1
            lbl20(i).ForeColor = RGB(0, 0, 0)
            lbl22(i).ForeColor = RGB(0, 0, 0)
        Next
        
        '�ŵ��ŷ����� �ܱ��� ȸ���� ���Ե� ��� �ŵ�ȸ����� û��(2002.1.9 BY LDH)
        '�ż��ŷ����� �ܱ��� ȸ���� ���Ե� ��� �ż�ȸ����� ����
        For i = 0 To memberobj.GetHeaderValue(1) - 1
        medotemp = memberobj.GetDataValue(0, i)
        If IsFrn(medotemp) Then
           lbl20(i).ForeColor = RGB(0, 0, 255)
        End If
        lbl20(i) = Convert(medotemp)
        lbl21(i) = FormatNumber(memberobj.GetDataValue(2, i), 0)
        mesutemp = memberobj.GetDataValue(1, i)
        If IsFrn(mesutemp) Then
           lbl22(i).ForeColor = RGB(255, 0, 0)
        End If
        lbl22(i) = Convert(mesutemp)
        lbl23(i) = FormatNumber(memberobj.GetDataValue(3, i), 0)
        Next
End Sub
Private Sub displ(gubun As Integer)
        If gubun = 0 Then
           For k = 0 To 5
                For i = 5 To 9
                lbȣ��(i + 10 * k).Visible = False
                Next i
           Next k
        Else
           For k = 0 To 5
                For i = 5 To 9
                lbȣ��(i + 10 * k).Visible = True
                Next i
           Next k
        End If
End Sub
' StockMst ����
Private Sub smobj_Received()
    lb���簡.BorderStyle = 0
    lb���簡.BackColor = RGB(255, 255, 255)
    lb����� = smobj.GetHeaderValue(1)
    sanghanga = smobj.GetHeaderValue(8)
    hahanga = smobj.GetHeaderValue(9)
    
    If smobj.GetHeaderValue(11) = sanghanga Then
    display smobj.GetHeaderValue(11), smobj.GetHeaderValue(12), 1, lb���簡
    lb���簡.BackColor = RGB(255, 0, 0)
    lb���簡.BorderStyle = 1
    ElseIf smobj.GetHeaderValue(11) = hahanga Then
    display smobj.GetHeaderValue(11), smobj.GetHeaderValue(12), 1, lb���簡
    lb���簡.BackColor = RGB(0, 0, 255)
    lb���簡.BorderStyle = 1
    Else
    display smobj.GetHeaderValue(11), smobj.GetHeaderValue(12), 0, lb���簡
    End If

    lb�ŵ�ȣ�� = FormatNumber(smobj.GetHeaderValue(16), 0)
    lb�ż�ȣ�� = FormatNumber(smobj.GetHeaderValue(17), 0)
    dis = smobj.GetHeaderValue(12)
    If dis > 0 Then
    Image2.Visible = False
    Image1.Visible = True
    ElseIf dis < 0 Then
    Image1.Visible = False
    Image2.Visible = True
    Else
    Image1.Visible = False
    Image2.Visible = False
    End If
    SetValueWithColor smobj.GetHeaderValue(12), lb���, True
    
    If smobj.GetHeaderValue(12) > 0 And smobj.GetHeaderValue(11) > 0 Then
        temp = CStr(smobj.GetHeaderValue(12) / (smobj.GetHeaderValue(11) - smobj.GetHeaderValue(12)) * 100)
    Else
        temp = 0
    End If
    pos = InStr(1, temp, ".")
    temptemp = Mid(temp, 1, pos + 2)
    
    SetValueWithColor temptemp, lb������, False
    lb�ŷ��� = FormatNumber(smobj.GetHeaderValue(18), 0)
    lbprevdaebi = FormatNumber(smobj.GetHeaderValue(46), 0)
    lb�ŷ���� = FormatNumber(smobj.GetHeaderValue(19), 0) + "����"
    lb�ܱ����ѵ� = FormatNumber(smobj.GetHeaderValue(37), 0)
    lb�ܱ����ѵ����� = "(" + FormatNumber(smobj.GetHeaderValue(38), 2) + "%)"
    lb�ܱ��ΰ��� = FormatNumber(smobj.GetHeaderValue(39), 0)
    lb�ܱ��ΰ��ɺ��� = "(" + FormatNumber(smobj.GetHeaderValue(40), 2) + "%)"
    n = smobj.GetHeaderValue(39) - smobj.GetHeaderValue(36)
    SetValueWithColor n, lb�ܱ��κ���, True
    lbmedan = smobj.GetHeaderValue(43) & "��"
    
    lbltprice = FormatNumber(sanghanga, 0)
    lblbprice = FormatNumber(hahanga, 0)
    
    If gubun2 = 0 Then
       lbl13.Caption = "52�� �ְ�"
       lbl16.Caption = "52�� ����"
       highprice = smobj.GetHeaderValue(47)
       highdate = smobj.GetHeaderValue(48)
       lowprice = smobj.GetHeaderValue(49)
       lowdate = smobj.GetHeaderValue(50)
    Else
        lbl13.Caption = "���� �ְ�"
        lbl16.Caption = "���� ����"
        highprice = smobj.GetHeaderValue(21)
        highdate = smobj.GetHeaderValue(22)
        lowprice = smobj.GetHeaderValue(23)
        lowdate = smobj.GetHeaderValue(24)
    End If
    
    lbltop = FormatNumber(highprice, 0) & "(" & datedisplay(highdate) & ")"
    lbldown = FormatNumber(lowprice, 0) & "(" & datedisplay(lowdate) & ")"
    
    'kospi200,kosdaq50 ä�� ����
    If Chr(smobj.GetHeaderValue(45)) = "1" Then
       lblhiregubun = "KOSPI200" & "(" & smobj.GetHeaderValue(53) & ")"
    ElseIf Chr(smobj.GetHeaderValue(45)) = "5" Then
       lblhiregubun = "KOSDAQ50" & "(" & smobj.GetHeaderValue(53) & ")"
    Else
       lblhiregubun = ""
    End If
    
    lblgiupgubun = smobj.GetHeaderValue(52)
    '20020107 by ldh (�Ǹ� ������ ���� ��쿡�� ���ذ��� �����ش�)
    righttemp = smobj.GetHeaderValue(51)
    If righttemp = "" Then
       lblrightgubun = ""
       lblrightvalue = ""
    Else
       lblrightgubun = righttemp
       lblrightgubun.ForeColor = RGB(255, 0, 0)
       lblrightvalue = "(" & FormatNumber(CStr(smobj.GetHeaderValue(27)), 0) & ")"
    End If
    
    lblgubun = smobj.GetHeaderValue(5)
    lblfinishmonth = smobj.GetHeaderValue(26) & "�� ���"
    
    lblnominal = "�׸鰡: " & smobj.GetHeaderValue(54)    '�׸鰡
End Sub
' StockJpbid2 ����
Private Sub sjb2obj_Received()
    For i = 0 To 9
        lbȣ��(i + 0) = FormatNumber(sjb2obj.GetDataValue(0, i), 0) '�ŵ�ȣ��
        lbȣ��(i + 30) = FormatNumber(sjb2obj.GetDataValue(1, i), 0) '�ż�ȣ��
        lbȣ��(i + 10) = FormatNumber(sjb2obj.GetDataValue(2, i), 0) '�ŵ�ȣ���ܷ�
        lbȣ��(i + 40) = FormatNumber(sjb2obj.GetDataValue(3, i), 0) '�ż�ȣ���ܷ�
        
        SetValueWithColor sjb2obj.GetDataValue(4, i), lbȣ��(i + 20), True '�ŵ�ȣ���ܷ����
        SetValueWithColor sjb2obj.GetDataValue(5, i), lbȣ��(i + 50), True '�ż�ȣ���ܷ����
        If bPrevRestInit = False Then
            aPrevRest(i) = sjb2obj.GetDataValue(2, i) ' �ŵ�ȣ���ܷ�
            aPrevRest(i + 20) = sjb2obj.GetDataValue(0, i) ' �ŵ�ȣ��
            aPrevRest(i + 10) = sjb2obj.GetDataValue(3, i) ' �ż�ȣ���ܷ�
            aPrevRest(i + 30) = sjb2obj.GetDataValue(1, i) ' �ż�ȣ��
        End If
    Next
    '�ð� ó��
    segan = CStr(sjb2obj.GetHeaderValue(3))
    If Len(segan) = 3 Then
    lb�ð� = Left(segan, 1) & ":" & Right(segan, 2)
    Else
    lb�ð� = Left(segan, 2) & ":" & Right(segan, 2)
    End If
    'ending
    leftright sjb2obj.GetHeaderValue(4), lb�Ѹŵ��ܷ�, 1
    leftright sjb2obj.GetHeaderValue(6), lb�Ѹż��ܷ�, 2
    leftright sjb2obj.GetHeaderValue(8), lb�ð��ܸŵ��ܷ�, 1
    leftright sjb2obj.GetHeaderValue(10), lb�ð��ܸż��ܷ�, 2

    SetValueWithColor sjb2obj.GetHeaderValue(5), lb�Ѹŵ��ܷ����, True
    SetValueWithColor sjb2obj.GetHeaderValue(7), lb�Ѹż��ܷ����, True
    SetValueWithColor sjb2obj.GetHeaderValue(9), lb�ð��ܸŵ��ܷ����, True
    SetValueWithColor sjb2obj.GetHeaderValue(11), lb�ð��ܸż��ܷ����, True
    
    aPrevRest(40) = sjb2obj.GetHeaderValue(4)
    aPrevRest(41) = sjb2obj.GetHeaderValue(6)
    aPrevRest(42) = sjb2obj.GetHeaderValue(8)
    aPrevRest(43) = sjb2obj.GetHeaderValue(10)
    
    bPrevRestInit = True
    
    displ (gubun)
    
End Sub
' StockJpbid ����
Private Sub sjbobj_Received()
    For i = 0 To 9
        '5������ 10���� �Ǹ鼭 ����� ���� �����־�� �Ѵ�.. 6�� �̻��� ���� 4���� ������� �Ѵ�
        If i < 5 Then
        lbȣ��(i + 0) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 3), 0) '�ŵ�ȣ��
        lbȣ��(i + 30) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 4), 0) '�ż�ȣ��
        lbȣ��(i + 10) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 5), 0) '�ŵ�ȣ���ܷ�
        lbȣ��(i + 40) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 6), 0) '�ż�ȣ���ܷ�
        Else
        lbȣ��(i + 0) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 7), 0) '�ŵ�ȣ��
        lbȣ��(i + 30) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 8), 0) '�ż�ȣ��
        lbȣ��(i + 10) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 9), 0) '�ŵ�ȣ���ܷ�
        lbȣ��(i + 40) = FormatNumber(sjbobj.GetHeaderValue(i * 4 + 10), 0) '�ż�ȣ���ܷ�
        End If
        
        If bPrevRestInit = True Then
            If i < 5 Then
            n = GetRestDiffer(sjbobj.GetHeaderValue(i * 4 + 3), sjbobj.GetHeaderValue(i * 4 + 5))
            n2 = GetRstDiffer(sjbobj.GetHeaderValue(i * 4 + 4), sjbobj.GetHeaderValue(i * 4 + 6))
            Else
            n = GetRestDiffer(sjbobj.GetHeaderValue(i * 4 + 7), sjbobj.GetHeaderValue(i * 4 + 9))
            n2 = GetRstDiffer(sjbobj.GetHeaderValue(i * 4 + 8), sjbobj.GetHeaderValue(i * 4 + 10))
            End If
            SetValueWithColor n, lbȣ��(i + 20), True
            SetValueWithColor n2, lbȣ��(i + 50), True
        End If
        
        If i < 5 Then
        aPrevRest(i) = sjbobj.GetHeaderValue(i * 4 + 5) '�ŵ�ȣ���ܷ�
        aPrevRest(i + 20) = sjbobj.GetHeaderValue(i * 4 + 3) '�ŵ�ȣ��
        aPrevRest(i + 10) = sjbobj.GetHeaderValue(i * 4 + 6) '�ż�ȣ���ܷ�
        aPrevRest(i + 30) = sjbobj.GetHeaderValue(i * 4 + 4) ' �ż�ȣ��
        Else
        aPrevRest(i) = sjbobj.GetHeaderValue(i * 4 + 9) '�ŵ�ȣ���ܷ�
        aPrevRest(i + 20) = sjbobj.GetHeaderValue(i * 4 + 7) '�ŵ�ȣ��
        aPrevRest(i + 10) = sjbobj.GetHeaderValue(i * 4 + 10) '�ż�ȣ���ܷ�
        aPrevRest(i + 30) = sjbobj.GetHeaderValue(i * 4 + 8) ' �ż�ȣ��
        End If
    Next
    '�ð� ǥ��
    seegan = CStr(sjbobj.GetHeaderValue(1))
    If Len(seegan) = 3 Then
    lb�ð� = Left(seegan, 1) & ":" & Right(seegan, 2)
    Else
    lb�ð� = Left(seegan, 2) & ":" & Right(seegan, 2)
    End If
    'ending
    
    leftright sjbobj.GetHeaderValue(23), lb�Ѹŵ��ܷ�, 1
    leftright sjbobj.GetHeaderValue(24), lb�Ѹż��ܷ�, 2
    leftright sjbobj.GetHeaderValue(25), lb�ð��ܸŵ��ܷ�, 1
    leftright sjbobj.GetHeaderValue(26), lb�ð��ܸż��ܷ�, 2
    
    SetValueWithColor sjbobj.GetHeaderValue(23) - aPrevRest(40), lb�Ѹŵ��ܷ����, True
    SetValueWithColor sjbobj.GetHeaderValue(24) - aPrevRest(41), lb�Ѹż��ܷ����, True
    SetValueWithColor sjbobj.GetHeaderValue(25) - aPrevRest(42), lb�ð��ܸŵ��ܷ����, True
    SetValueWithColor sjbobj.GetHeaderValue(26) - aPrevRest(43), lb�ð��ܸż��ܷ����, True
    
    
    aPrevRest(40) = sjbobj.GetHeaderValue(23)
    aPrevRest(41) = sjbobj.GetHeaderValue(24)
    aPrevRest(42) = sjbobj.GetHeaderValue(25)
    aPrevRest(43) = sjbobj.GetHeaderValue(26)
    bPrevRestInit = True
End Sub
' StockCur ����
Private Sub scobj_Received()
    lb���簡.BorderStyle = 0
    lb���簡.BackColor = RGB(255, 255, 255)
    
    lb����� = scobj.GetHeaderValue(1)
    n = scobj.GetHeaderValue(3)
    lb�ð� = FormatNumber(n / 100, 0) & ":" & n Mod 100
    
    If scobj.GetHeaderValue(13) = sanghanga Then
    display scobj.GetHeaderValue(13), scobj.GetHeaderValue(2), 1, lb���簡
    lb���簡.BackColor = RGB(255, 0, 0)
    lb���簡.BorderStyle = 1
    ElseIf scobj.GetHeaderValue(13) = hahanga Then
    display scobj.GetHeaderValue(13), scobj.GetHeaderValue(2), 1, lb���簡
    lb���簡.BackColor = RGB(0, 0, 255)
    lb���簡.BorderStyle = 1
    Else
    display scobj.GetHeaderValue(13), scobj.GetHeaderValue(2), 0, lb���簡
    End If
    lb�ŵ�ȣ�� = FormatNumber(scobj.GetHeaderValue(7), 0)
    lb�ż�ȣ�� = FormatNumber(scobj.GetHeaderValue(8), 0)
    disp = scobj.GetHeaderValue(2)
    If disp > 0 Then
    Image2.Visible = False
    Image1.Visible = True
    ElseIf disp < 0 Then
    Image1.Visible = False
    Image2.Visible = True
    Else
    Image1.Visible = False
    Image2.Visible = False
    End If
    SetValueWithColor scobj.GetHeaderValue(2), lb���, True
    SetValueWithColor scobj.GetHeaderValue(2) / (scobj.GetHeaderValue(13) - scobj.GetHeaderValue(2)) * 100, lb������, False
    lb�ŷ��� = FormatNumber(scobj.GetHeaderValue(9), 0)
    lb�ŷ���� = FormatNumber(scobj.GetHeaderValue(10), 0) + "����"
    'start
    If scobj.GetHeaderValue(13) > highprice Then
       highprice = scobj.GetHeaderValue(13)
       lbltop = FormatNumber(highprice, 0) & "(" & timedisplay() & ")"
    End If
    
    If scobj.GetHeaderValue(13) < lowprice Then
       lowprice = scobj.GetHeaderValue(13)
       lbldown = FormatNumber(lowprice, 0) & "(" & timedisplay() & ")"
    End If
    'ending
End Sub
Private Sub display(value1, value2, gubun, obj As Object)
    If gubun = 0 Then
        If value2 > 0 Then
        obj.ForeColor = RGB(255, 0, 0)
        ElseIf value2 = 0 Then
        obj.ForeColor = RGB(0, 0, 0)
        Else
        obj.ForeColor = RGB(0, 0, 255)
        End If
    ElseIf gubun = 1 Then
        obj.ForeColor = RGB(255, 255, 255)
    End If
        obj = FormatNumber(value1, 0)
End Sub
'�ڵ带 �޾ ����ȸ������� �ٲٱ� ����
Function Convert(code As Variant)
    Convert = m_CodeMgr.GetMemberName(code)
End Function
'���簡�� 52�� �ְ��� �������� �Ѿ��� ��
Function timedisplay()
        Dim str, strslice
        str = CStr(Now())
        strslice = Left(str, 4) & "/" & Mid(str, 6, 2) & "/" & Mid(str, 9, 2)
        timedisplay = strslice
End Function
'�ְ����̳� �������� //�� ǥ���ϱ� ����
Function datedisplay(num As Variant)
        Dim str, strslice
        str = CStr(num)
        strslice = Left(str, 4) & "/" & Mid(str, 5, 2) & "/" & Mid(str, 7, 2)
        datedisplay = strslice
End Function
Private Sub txtJongMok_Change()
       If Len(txtJongMok) = 6 Then
       cmd��ȸ_Click
       End If
End Sub
'�ܱ��� ȸ������ �����ϴ� �ڵ�
Function IsFrn(code As Variant)
    If (code = "033" Or code = "035" Or code = "036" Or code = "037" Or code = "038" Or code = "040" Or code = "041" Or code = "042" Or code = "043" Or code = "045" Or code = "054" Or code = "819" Or code = "820" Or code = "824" Or code = "058" Or code = "059" Or code = "044" Or code = "803" Or code = "807") Then
    IsFrn = True
    Else
    IsFrn = False
    End If
End Function
 
