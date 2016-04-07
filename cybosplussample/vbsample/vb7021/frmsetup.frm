VERSION 5.00
Begin VB.Form frmsetup 
   Caption         =   "설정 화면"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdexe 
      Caption         =   "확인"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "호가 표시"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
      Begin VB.OptionButton Option4 
         Caption         =   "10차 호가"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "5차 호가"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "최고최저가 표시"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "연중 최고/최저가로 표시"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "52주 최고/최저가로 표시"
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
'2002.1.11 by ldh 연중/52주 최고,5차/10차 호가 setting form
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
            frmsetup.Left = frm현재가.Width / 2
            frmsetup.Top = frm현재가.Height / 2
End Sub
