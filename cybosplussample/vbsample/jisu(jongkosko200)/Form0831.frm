VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdpi 
      Caption         =   "kospi200"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdko 
      Caption         =   "코스닥"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdjong 
      Caption         =   "종합 "
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtkospi 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtkosdaq 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtjong 
      Height          =   390
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "kospi 200 지수"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "코스닥 지수"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "종합 주가 지수"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents srjobj As StockIndexir
Attribute srjobj.VB_VarHelpID = -1
Private WithEvents srkobj As StockIndexir
Attribute srkobj.VB_VarHelpID = -1
Private WithEvents ssjobj As StockIndexis
Attribute ssjobj.VB_VarHelpID = -1
Private WithEvents sskobj As StockIndexis
Attribute sskobj.VB_VarHelpID = -1
Private WithEvents frobj As FutureIndexh
Attribute frobj.VB_VarHelpID = -1
Private WithEvents fsobj As FutureIndexi
Attribute fsobj.VB_VarHelpID = -1
 

Private Sub cmdjong_Click()
            srjobj.SetInputValue 0, "001"
            srjobj.Request
            
            ssjobj.Unsubscribe
            ssjobj.SetInputValue 0, "001"
            ssjobj.SubscribeLatest
                        
End Sub

Private Sub cmdko_Click()
            srkobj.SetInputValue 0, "201"
            srkobj.Request
            
            sskobj.Unsubscribe
            sskobj.SetInputValue 0, "201"
            sskobj.SubscribeLatest
End Sub

Private Sub cmdpi_Click()
            frobj.SetInputValue 0, "00800"
            frobj.Request
            
            fsobj.Unsubscribe
            fsobj.SetInputValue 0, "00800"
            fsobj.SubscribeLatest
            
End Sub

Private Sub Form_Load()
        Set srjobj = New StockIndexir
        Set srkobj = New StockIndexir
        Set ssjobj = New StockIndexis
        Set sskobj = New StockIndexis
        Set frobj = New FutureIndexh
        Set fsobj = New FutureIndexi
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set srjobj = Nothing
        Set srkobj = Nothing
        Set ssjobj = Nothing
        Set sskobj = Nothing
        Set frobj = Nothing
        Set fsobj = Nothing
End Sub


Private Sub frobj_Received()
         txtkospi.Text = FormatNumber(frobj.GetDataValue(1, 0), 2)
         
End Sub

Private Sub fsobj_Received()
        txtkospi.Text = fsobj.GetHeaderValue(2)
        
End Sub

Private Sub srjobj_Received()
          txtjong.Text = FormatNumber(srjobj.GetDataValue(1, 0), 2)
          
End Sub

Private Sub srkobj_Received()
          txtkosdaq.Text = FormatNumber(srkobj.GetDataValue(1, 0), 2)
          
End Sub

Private Sub ssjobj_Received()
         txtjong.Text = ssjobj.GetHeaderValue(1) & "\" & FormatNumber(ssjobj.GetHeaderValue(2), 2)
         
End Sub

Private Sub sskobj_Received()
         txtkosdaq.Text = sskobj.GetHeaderValue(1) & "\" & FormatNumber(sskobj.GetHeaderValue(2), 2)
End Sub
