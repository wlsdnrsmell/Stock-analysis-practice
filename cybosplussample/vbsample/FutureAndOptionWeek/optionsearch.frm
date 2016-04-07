VERSION 5.00
Begin VB.Form optionsearch 
   Caption         =   "종목검색"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox ListBox2 
      Height          =   2400
      Left            =   2520
      TabIndex        =   9
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ListBox ListBox1 
      Height          =   2400
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   3
      Left            =   4920
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   2
      Left            =   4920
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Button 
      Caption         =   "Command3"
      Height          =   495
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TextBox1 
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "풋"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "콜"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "종목코드"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "optionsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    mainform.O_textbox.Text = tmp
    Hide
Else
    MsgBox ("종목을 정확히 선택해주십시오")
End If
End Sub

Private Sub Command2_Click()
Hide
End Sub


Private Sub ListBox1_Click()
TextBox1.Text = Left(ListBox1.Text, 8)
ListBox2.ListIndex = -1
End Sub

Private Sub ListBox2_Click()
ListBox1.ListIndex = -1
TextBox1.Text = Left(ListBox2.Text, 8)
End Sub
