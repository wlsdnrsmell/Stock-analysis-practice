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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdOrder 
      Caption         =   "��    ��"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox txt�ֹ�ȣ������ 
      Alignment       =   1  '������ ����
      Height          =   270
      Left            =   1440
      TabIndex        =   16
      Text            =   "01"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txt�ֹ����� 
      Alignment       =   1  '������ ����
      Height          =   270
      Left            =   1440
      TabIndex        =   14
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  '������ ����
      Height          =   270
      Left            =   1440
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtVol 
      Alignment       =   1  '������ ����
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
      Alignment       =   1  '������ ����
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
      Caption         =   "�ż�"
      Height          =   300
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   795
      Width           =   1215
   End
   Begin VB.OptionButton optOrder 
      Caption         =   "�ŵ�"
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   795
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Label lb���� 
      BackColor       =   &H00C0E0FF&
      Height          =   550
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label10 
      Caption         =   "ü����ų���(CYBOS�� ƼĿ��)"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Label Label9 
      Caption         =   "�ֹ��������(CYBOS�� 0311 �ֹ����)"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label lbü����� 
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
      Caption         =   "�ֹ�ȣ������"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "�ֹ������ڵ�"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "���񰡰�"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "�������"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "�����ڵ�"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "��ǰ��������"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "���¹�ȣ"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�ֹ�����"
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

'ü����� �̺�Ʈ�Դϴ�.
Private Sub cpc_Received()
    lbResult.Caption = "[" + CStr(cpc.GetHeaderValue(2)) + "] �ֹ���ȣ[" + CStr(cpc.GetHeaderValue(5)) + "]" + ü�ᱸ�г���(cpc.GetHeaderValue(14))
End Sub

Private Sub Form_Load()
    Set m_CpTdUtil = New CpTdUtil
    
    If m_CpTdUtil.TradeInit() <> 0 Then
        MsgBox ("�Է°� �����Դϴ�. ����� �Ͻʽÿ�")
        Unload Me
    End If
    
    lb����.Caption = "�� ������ �ֹ��� ���������Դϴ�." + Chr(13) + "�ֹ������ڵ�,�ֹ�ȣ���ڵ���� " + Chr(13) + "���� �Է°��� ������ �����ϼſ�"
    
    Me.BackColor = RGB(206, 219, 239)
    
    Set m_CodeMgr = New CpCodeMgr
    Set m_0311 = New CpTd0311
    Set cpc = New CpConclusion
    cpc.Subscribe
    
    '��ǥ����ǥ��
    cmbAccount.Text = m_CpTdUtil.AccountNumber(0)
    
    '�������µ� �� ǥ���ҷ��� ����ó���߽��ϴ�.
    '�������¸� ������� ������ ����ó���� �����ص� �˴ϴ�.
    Dim ar, i
    ar = m_CpTdUtil.AccountNumber
    For i = LBound(ar) To UBound(ar)
        cmbAccount.AddItem (m_CpTdUtil.AccountNumber(i))
    Next
End Sub
Private Sub cmdOrder_Click()

    If optOrder(0).Value = True Then '�ŵ�
        m_0311.SetInputValue 0, "1"
    Else '�ż�
        m_0311.SetInputValue 0, "2"
    End If
    
    m_0311.SetInputValue 1, cmbAccount.Text
    m_0311.SetInputValue 2, txtAccGubun.Text
    m_0311.SetInputValue 3, "A" + txtCode.Text '�ֽ��ΰ�� "A" elw�̸� "J"
    m_0311.SetInputValue 4, CLng(txtVol.Text)
    m_0311.SetInputValue 5, CLng(txtPrice.Text)
    m_0311.SetInputValue 7, txt�ֹ�����.Text
    m_0311.SetInputValue 8, txt�ֹ�ȣ������.Text
    
    Dim ret, msg
    ret = m_0311.BlockRequest
    
    '****************************************************
    '��� CybosPlus ������Ʈ�� BlockRequest ���ϰ��� GetDibStatus �Ӽ�������
    
    'BlockRequest ���ϰ����� ��� ����� �˼� �ֽ��ϴ�.
        '0:  Success
        '1:  Error -TimeOut
        '3 : Error - �� ���� ����
    
    'GetDibStatus �Ӽ������� ��û ����� �˼� �ֽ��ϴ�.
        '-1 - ����
        '0 - ����
        '1 - ���Ŵ��
    '****************************************************
    
    '****************************************************
    'BlockRequest ���� (��ż���)
    '****************************************************
    If ret = 0 Then
        '****************************************************
        '�ֹ���û(���)�� �����̰�. ��û����� ������ ���
        '****************************************************
        If m_0311.GetDibStatus = 0 Then
            lbResult.Caption = "[�ֹ�����][�ֹ���ȣ]" + CStr(m_0311.GetHeaderValue(8)) + "[" + m_0311.GetDibMsg1 + "]"
        '****************************************************
        '�ֹ���û�� �����̳�, ���ű� ���� ����� ������ �ֹ��� ���� ���� ���Ͽ���.
        '****************************************************
        Else
            lbResult.Caption = "[�ֹ�����][" + m_0311.GetDibMsg1 + "]"
        End If
    '****************************************************
    'BlockRequest ���� (��Ž���)
    '****************************************************
    Else
        lbResult.Caption = "BlockRequest TimeOut �� �׹��� ����"
    End If
End Sub
Private Function ü�ᱸ�г���(sGubun)
    Dim sRet
    Select Case sGubun
        Case "1"
            sRet = "[ü��]"
        Case "2"
            sRet = "[Ȯ��]"
        Case "3"
            sRet = "[�ź�]"
        Case "4"
            sRet = "[����]"
        Case "5"
            sRet = "[�������]"
    End Select
    ü�ᱸ�г��� = "[" + sGubun + "]" + sRet
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
            lbCodeName.Caption = "�Է� �����ڵ� ����"
        Else
            lbCodeName.Caption = name
        End If
    End If
End Sub
