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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdjong 
      Caption         =   "��/��/3"
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
      BorderStyle     =   1  '���� ����
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
'�����ڵ带 �Է� �޾� �� ������ ����(�ŷ���/�ڽ���/��������)�� �˾Ƴ��� vb �ڵ��̴�

Private Sub cmdjong_Click()
            Set util = New CpCodeMgr
            
            code = txtjong.Text
            codename = util.CodeToName(code)
            Select Case util.GetStockMarketKind(txtjong.Text)
                Case CPC_MARKET_KOSPI
                    lbljong = "������� " & codename & "�̸� " & "�ŷ��� �����Դϴ�"
                Case CPC_MARKET_KOSDAQ
                    lbljong = "������� " & codename & "�̸� " & "�ڽ��� �����Դϴ�"
                Case CPC_MARKET_FREEBOARD
                   lbljong = "������� " & codename & "�̸� " & "�������� �����Դϴ�"
                Case CPC_MARKET_KRX
                   lbljong = "������� " & codename & "�̸� " & "krx �����Դϴ�"
                Case Else
                   lbljong = "�����ڵ带 Ʋ���� �Է��ϼ̰ų� ���� ���� �ڵ��Դϴ�"
            End Select
                
            Set util = Nothing
End Sub
