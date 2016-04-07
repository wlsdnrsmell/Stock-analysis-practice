VERSION 5.00
Begin VB.Form optionsearch 
   BorderStyle     =   1  '단일 고정
   Caption         =   "종목검색"
   ClientHeight    =   3405
   ClientLeft      =   5745
   ClientTop       =   2190
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5895
   Begin VB.ListBox ListBox2 
      Height          =   2400
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   2295
   End
   Begin VB.ListBox ListBox1 
      Height          =   2400
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   3
      Left            =   4920
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   2
      Left            =   4920
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   700
   End
   Begin VB.TextBox TextBox1 
      Height          =   370
      Left            =   1080
      TabIndex        =   1
      Top             =   130
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "풋"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "콜"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "종목코드"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   178
      Width           =   975
   End
End
Attribute VB_Name = "optionsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public optionlist As CpOptionCode

Private Sub List2_Click()
TextBox1.Text = Left(ListBox1.Text, 8)
ListBox2.ListIndex = -1
End Sub

Private Sub Button_Click(Index As Integer)
For i = 0 To ListBox1.ListCount - 1
    If Right(ListBox1.List(i), 9) > Button(Index).Caption Then
        ListBox1.ListIndex = -1
       ListBox2.ListIndex = -1
        ListBox1.TopIndex = i
        ListBox2.TopIndex = i
        Exit Sub
    End If
Next i

End Sub

Private Sub Command1_Click()
tmp = TextBox1.Text
If Len(tmp) = 8 Then
    'mainform.O_textbox.Text = tmp
    tr9721.txt_OCode = tmp
    tr9721.lb_jongmok = optionlist.CodeToName(tmp)
    tr9721.data_susin
    Hide
Else
    MsgBox ("종목을 정확히 선택해주십시오")
End If
TextBox1.Text = ""
ListBox1.ListIndex = -1
ListBox2.ListIndex = -1
End Sub

Private Sub Command2_Click()
Hide
TextBox1.Text = ""
ListBox1.ListIndex = -1
ListBox2.ListIndex = -1
End Sub

Private Sub Form_Load()
    SetTopMostWindow Me.hWnd, True
End Sub

Private Sub ListBox1_Click()
TextBox1.Text = Left(ListBox1.Text, 8)
ListBox2.ListIndex = -1
End Sub

Private Sub ListBox1_DblClick()
TextBox1.Text = Left(ListBox1.Text, 8)
ListBox2.ListIndex = -1
tmp = TextBox1.Text
If Len(tmp) = 8 Then
    'mainform.O_textbox.Text = tmp
    tr9721.txt_OCode = tmp
    tr9721.lb_jongmok = optionlist.CodeToName(tmp)
    tr9721.data_susin
    Hide
Else
    MsgBox ("종목을 정확히 선택해주십시오")
End If
TextBox1.Text = ""
ListBox1.ListIndex = -1
ListBox2.ListIndex = -1
End Sub

Private Sub ListBox2_Click()
ListBox1.ListIndex = -1
TextBox1.Text = Left(ListBox2.Text, 8)
End Sub
Sub set_OCode()
    Set optionlist = New CpOptionCode
    
    Button(0).Caption = optionlist.GetData(3, 0)
    
    For i = 1 To optionlist.GetCount - 1
    If Button(0).Caption <> optionlist.GetData(3, i) Then
       Button(1).Caption = optionlist.GetData(3, i)
       Exit For
    End If
    Next i
    
    For j = i To optionlist.GetCount - 1
    If Button(1).Caption <> optionlist.GetData(3, j) Then
       Button(2).Caption = optionlist.GetData(3, j)
       Exit For
    End If
    Next j
    
    
    For k = j To optionlist.GetCount - 1
    If Button(2).Caption <> optionlist.GetData(3, k) Then
       Button(3).Caption = optionlist.GetData(3, k)
       Exit For
    End If
    Next k
    
    '리스트 채우기

For i = 0 To optionlist.GetCount - 1
    If optionlist.GetData(2, i) = "풋" Then
        Exit For
    End If
    ListBox1.AddItem optionlist.GetData(0, i) + "   " + optionlist.GetData(1, i)
Next i

For j = i To optionlist.GetCount - 1
    ListBox2.AddItem optionlist.GetData(0, j) + "   " + optionlist.GetData(1, j)
Next j
End Sub

Private Sub ListBox2_DblClick()
ListBox1.ListIndex = -1
TextBox1.Text = Left(ListBox2.Text, 8)

tmp = TextBox1.Text
If Len(tmp) = 8 Then
    'mainform.O_textbox.Text = tmp
    tr9721.txt_OCode = tmp
    tr9721.lb_jongmok = optionlist.CodeToName(tmp)
    tr9721.data_susin
    Hide
Else
    MsgBox ("종목을 정확히 선택해주십시오")
End If
TextBox1.Text = ""
ListBox1.ListIndex = -1
ListBox2.ListIndex = -1
End Sub
