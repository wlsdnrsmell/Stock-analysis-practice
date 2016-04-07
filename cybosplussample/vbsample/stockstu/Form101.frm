VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdclear 
      Caption         =   "지우기"
      Height          =   615
      Left            =   4920
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox lsttotal 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   3840
      Left            =   3480
      TabIndex        =   9
      Top             =   1200
      Width           =   4575
   End
   Begin VB.CommandButton cmdsearch 
      BackColor       =   &H00FFC0FF&
      Caption         =   "조회"
      Height          =   615
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtcounter 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   800
   End
   Begin VB.TextBox txtju 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   800
   End
   Begin VB.TextBox txtmode 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   800
   End
   Begin VB.TextBox txtjong 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   800
   End
   Begin VB.Label lbdisplay 
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3480
      TabIndex        =   10
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Label lb4 
      BorderStyle     =   1  '단일 고정
      Caption         =   "카운터"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1000
   End
   Begin VB.Label lb3 
      BorderStyle     =   1  '단일 고정
      Caption         =   "주기"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1000
   End
   Begin VB.Label lb2 
      BorderStyle     =   1  '단일 고정
      Caption         =   "모드"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1000
   End
   Begin VB.Label lb1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "종목 코드"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents stockstuobj As StockStu
Attribute stockstuobj.VB_VarHelpID = -1

Private Sub cmdclear_Click()
        lsttotal.Clear
        txtjong = ""
        txtmode = ""
        txtcounter = ""
        txtju = ""
        lbdisplay = ""
        txtjong.SetFocus
        
        
End Sub

Private Sub cmdsearch_Click()
        Dim s, n
        stockstuobj.SetInputValue 0, txtjong
        
        stockstuobj.SetInputValue 1, Asc(txtmode)
        stockstuobj.SetInputValue 2, CInt(txtju)
        stockstuobj.SetInputValue 3, CInt(txtcounter)
        stockstuobj.BlockRequest
        s = "일자      시간  시가  고가   저가  종가  거래량"
        lsttotal.AddItem (s)
        n = stockstuobj.GetHeaderValue(3)
    
    For i = 0 To n - 1                  ' 수신 데이터 수만큼 루프를 돔
        
        s = stockstuobj.GetDataValue(0, i) & " "           ' 일자
        s = s & stockstuobj.GetDataValue(1, i) & " "       ' 시간
        s = s & stockstuobj.GetDataValue(2, i) & " "       ' 시가
        s = s & stockstuobj.GetDataValue(3, i) & " "       ' 고가
        s = s & stockstuobj.GetDataValue(4, i) & " "       ' 저다
        s = s & stockstuobj.GetDataValue(5, i) & " "       ' 종가
        s = s & stockstuobj.GetDataValue(6, i) & " "   ' 거래량
        ' 해당 항목을 HTML에 추가한다.
        lsttotal.AddItem (s)
        
    Next

         
End Sub

Private Sub Form_Load()
        MsgBox ("종목 코드을 입력하실때는 A를 포함한 6자리 코드를 입력하시어야 하며 모드 코드는 영어 대문자로 틱 차트 데이터를 원하시면 T를,분 차트 데이터를 원하시면 M를  입력하시면 되고 주기에는 데이터의 간격을 입력하시면 되고 카운터에는 보고싶은 토털 데이터 갯수를 입력하시면 됩니다")
        
        Set stockstuobj = New StockStu
        txtjong = "A00660"
        txtmode = "M"
        txtcounter = 20
        txtju = 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set stockstuobj = Nothing
End Sub
Private Sub lsttotal_Click()
         lbdisplay = lsttotal.Text
         
End Sub
