VERSION 5.00
Begin VB.Form frmsetup 
   Caption         =   "���� ȭ��"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdexe 
      Caption         =   "Ȯ��"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "ȣ�� ǥ��"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
      Begin VB.OptionButton Option4 
         Caption         =   "10�� ȣ��"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "5�� ȣ��"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ְ������� ǥ��"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "���� �ְ�/�������� ǥ��"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "52�� �ְ�/�������� ǥ��"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2002.1.11 by ldh ����/52�� �ְ�,5��/10�� ȣ�� setting form
Public envirhoga As Integer
Public enviryear As Integer
Private Sub cmdexe_Click()
        If Option3.value = True Then
           envirhoga = 0
        Else
           envirhoga = 1
        End If
        If Option1.value = True Then
           enviryear = 0
        Else
           enviryear = 1
        End If
        Me.Hide
End Sub
Private Sub Form_Load()
            Option4.value = True
            Option1.value = True
            envirhoga = 1
            enviryear = 0
            frmsetup.Left = frm���簡.Width / 2
            frmsetup.Top = frm���簡.Height / 2
End Sub
