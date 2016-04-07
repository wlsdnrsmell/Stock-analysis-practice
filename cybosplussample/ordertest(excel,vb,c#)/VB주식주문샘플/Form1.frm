VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdOrder 
      Caption         =   "주    문"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox txt주문호가구분 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1440
      TabIndex        =   16
      Text            =   "01"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txt주문조건 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1440
      TabIndex        =   14
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1440
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtVol 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1440
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtCode 
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtAccGubun 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1440
      TabIndex        =   6
      Text            =   "10"
      Top             =   1560
      Width           =   375
   End
   Begin VB.ComboBox cmbAccount 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.OptionButton optOrder 
      Caption         =   "매수"
      Height          =   300
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   795
      Width           =   1215
   End
   Begin VB.OptionButton optOrder 
      Caption         =   "매도"
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   795
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Label lb공지 
      BackColor       =   &H00C0E0FF&
      Height          =   550
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label10 
      Caption         =   "체결수신내용(CYBOS의 티커바)"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Label Label9 
      Caption         =   "주문결과내용(CYBOS의 0311 주문결과)"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label lb체결수신 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5925
      Width           =   3735
   End
   Begin VB.Label lbResult 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   4650
      Width           =   3735
   End
   Begin VB.Label lbCodeName 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "주문호가구분"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "주문조건코드"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "종목가격"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "종목수량"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "종목코드"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "상품관리구분"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "계좌번호"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "주문종류"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   795
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_CodeMgr As CpCodeMgr
Dim m_CpTdUtil As CpTdUtil
Dim m_0311 As CpTd0311
Private WithEvents cpc As CpConclusion
Attribute cpc.VB_VarHelpID = -1

'체결수신 이벤트입니다.
Private Sub cpc_Received()
    lbResult.Caption = "[" + CStr(cpc.GetHeaderValue(2)) + "] 주문번호[" + CStr(cpc.GetHeaderValue(5)) + "]" + 체결구분내용(cpc.GetHeaderValue(14))
End Sub

Private Sub Form_Load()
    Set m_CpTdUtil = New CpTdUtil
    
    If m_CpTdUtil.TradeInit() <> 0 Then
        MsgBox ("입력값 오류입니다. 재실행 하십시오")
        Unload Me
    End If
    
    lb공지.Caption = "본 샘플은 주문용 간략샘플입니다." + Chr(13) + "주문조건코드,주문호가코드등의 " + Chr(13) + "상세한 입력값은 도움말을 참고하셔요"
    
    Me.BackColor = RGB(206, 219, 239)
    
    Set m_CodeMgr = New CpCodeMgr
    Set m_0311 = New CpTd0311
    Set cpc = New CpConclusion
    cpc.Subscribe
    
    '대표계좌표시
    cmbAccount.Text = m_CpTdUtil.AccountNumber(0)
    
    '복수계좌들 다 표시할려고 다음처럼했습니다.
    '복수계좌를 사용하지 않으면 다음처리는 생략해도 됩니다.
    Dim ar, i
    ar = m_CpTdUtil.AccountNumber
    For i = LBound(ar) To UBound(ar)
        cmbAccount.AddItem (m_CpTdUtil.AccountNumber(i))
    Next
End Sub
Private Sub cmdOrder_Click()

    If optOrder(0).Value = True Then '매도
        m_0311.SetInputValue 0, "1"
    Else '매수
        m_0311.SetInputValue 0, "2"
    End If
    
    m_0311.SetInputValue 1, cmbAccount.Text
    m_0311.SetInputValue 2, txtAccGubun.Text
    m_0311.SetInputValue 3, "A" + txtCode.Text '주식인경우 "A" elw이면 "J"
    m_0311.SetInputValue 4, CLng(txtVol.Text)
    m_0311.SetInputValue 5, CLng(txtPrice.Text)
    m_0311.SetInputValue 7, txt주문조건.Text
    m_0311.SetInputValue 8, txt주문호가구분.Text
    
    Dim ret, msg
    ret = m_0311.BlockRequest
    
    '****************************************************
    '모든 CybosPlus 오브젝트는 BlockRequest 리턴값과 GetDibStatus 속성값으로
    
    'BlockRequest 리턴값으로 통신 결과를 알수 있습니다.
        '0:  Success
        '1:  Error -TimeOut
        '3 : Error - 그 밖의 오류
    
    'GetDibStatus 속성값으로 요청 결과를 알수 있습니다.
        '-1 - 오류
        '0 - 정상
        '1 - 수신대기
    '****************************************************
    
    '****************************************************
    'BlockRequest 성공 (통신성공)
    '****************************************************
    If ret = 0 Then
        '****************************************************
        '주문요청(통신)은 성공이고. 요청결과도 성공인 경우
        '****************************************************
        If m_0311.GetDibStatus = 0 Then
            lbResult.Caption = "[주문성공][주문번호]" + CStr(m_0311.GetHeaderValue(8)) + "[" + m_0311.GetDibMsg1 + "]"
        '****************************************************
        '주문요청은 성공이나, 증거금 부족 등등의 이유로 주문이 접수 되지 못하였음.
        '****************************************************
        Else
            lbResult.Caption = "[주문실패][" + m_0311.GetDibMsg1 + "]"
        End If
    '****************************************************
    'BlockRequest 실패 (통신실패)
    '****************************************************
    Else
        lbResult.Caption = "BlockRequest TimeOut 및 그밖의 오류"
    End If
End Sub
Private Function 체결구분내용(sGubun)
    Dim sRet
    Select Case sGubun
        Case "1"
            sRet = "[체결]"
        Case "2"
            sRet = "[확인]"
        Case "3"
            sRet = "[거부]"
        Case "4"
            sRet = "[접수]"
        Case "5"
            sRet = "[접수대기]"
    End Select
    체결구분내용 = "[" + sGubun + "]" + sRet
End Function

Private Sub optOrder_Click(Index As Integer)
    If Index = 0 Then
        Me.BackColor = RGB(206, 219, 239)
    Else
        Me.BackColor = RGB(255, 215, 222)
    End If
End Sub

Private Sub txtCode_Change()
    If Len(txtCode.Text) = 6 Then
        Dim name
        name = m_CodeMgr.CodeToName(txtCode.Text)
        If Len(name) = 0 Then
            lbCodeName.Caption = "입력 종목코드 오류"
        Else
            lbCodeName.Caption = name
        End If
    End If
End Sub
