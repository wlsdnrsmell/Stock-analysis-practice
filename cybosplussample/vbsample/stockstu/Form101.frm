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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdclear 
      Caption         =   "�����"
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
      Caption         =   "��ȸ"
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
      BorderStyle     =   1  '���� ����
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3480
      TabIndex        =   10
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Label lb4 
      BorderStyle     =   1  '���� ����
      Caption         =   "ī����"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1000
   End
   Begin VB.Label lb3 
      BorderStyle     =   1  '���� ����
      Caption         =   "�ֱ�"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1000
   End
   Begin VB.Label lb2 
      BorderStyle     =   1  '���� ����
      Caption         =   "���"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1000
   End
   Begin VB.Label lb1 
      BorderStyle     =   1  '���� ����
      Caption         =   "���� �ڵ�"
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
        s = "����      �ð�  �ð�  ��   ����  ����  �ŷ���"
        lsttotal.AddItem (s)
        n = stockstuobj.GetHeaderValue(3)
    
    For i = 0 To n - 1                  ' ���� ������ ����ŭ ������ ��
        
        s = stockstuobj.GetDataValue(0, i) & " "           ' ����
        s = s & stockstuobj.GetDataValue(1, i) & " "       ' �ð�
        s = s & stockstuobj.GetDataValue(2, i) & " "       ' �ð�
        s = s & stockstuobj.GetDataValue(3, i) & " "       ' ��
        s = s & stockstuobj.GetDataValue(4, i) & " "       ' ����
        s = s & stockstuobj.GetDataValue(5, i) & " "       ' ����
        s = s & stockstuobj.GetDataValue(6, i) & " "   ' �ŷ���
        ' �ش� �׸��� HTML�� �߰��Ѵ�.
        lsttotal.AddItem (s)
        
    Next

         
End Sub

Private Sub Form_Load()
        MsgBox ("���� �ڵ��� �Է��ϽǶ��� A�� ������ 6�ڸ� �ڵ带 �Է��Ͻþ�� �ϸ� ��� �ڵ�� ���� �빮�ڷ� ƽ ��Ʈ �����͸� ���Ͻø� T��,�� ��Ʈ �����͸� ���Ͻø� M��  �Է��Ͻø� �ǰ� �ֱ⿡�� �������� ������ �Է��Ͻø� �ǰ� ī���Ϳ��� ������� ���� ������ ������ �Է��Ͻø� �˴ϴ�")
        
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
