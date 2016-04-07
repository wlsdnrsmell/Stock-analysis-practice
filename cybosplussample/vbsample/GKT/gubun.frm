VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdjong 
      Caption         =   "거/코/3"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtjong 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lbljong 
      BorderStyle     =   1  '단일 고정
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'종목코드를 입력 받아 그 종목의 시장(거래소/코스닥/프리보드)을 알아내는 vb 코드이다

Private Sub cmdjong_Click()
            Set util = New CpCodeMgr
            
            code = txtjong.Text
            codename = util.CodeToName(code)
            Select Case util.GetStockMarketKind(txtjong.Text)
                Case CPC_MARKET_KOSPI
                    lbljong = "종목명은 " & codename & "이며 " & "거래소 종목입니다"
                Case CPC_MARKET_KOSDAQ
                    lbljong = "종목명은 " & codename & "이며 " & "코스닥 종목입니다"
                Case CPC_MARKET_FREEBOARD
                   lbljong = "종목명은 " & codename & "이며 " & "프리보드 종목입니다"
                Case CPC_MARKET_KRX
                   lbljong = "종목명은 " & codename & "이며 " & "krx 종목입니다"
                Case Else
                   lbljong = "종목코드를 틀리게 입력하셨거나 없는 종목 코드입니다"
            End Select
                
            Set util = Nothing
End Sub
