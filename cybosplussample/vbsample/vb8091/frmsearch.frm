VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form jongmoksearch 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.TabStrip Tabjong 
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "색인"
            Object.ToolTipText     =   "찾을 종목명을 입력해주세요"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstjong 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtjongword 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblshow 
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "jongmoksearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrjong() As String
Private cnt As Integer
Private Sub Form_Load()
          Set stockobj = New CpStockCode
          cnt = stockobj.GetCount
          ReDim arrjong(cnt)
          For m = 0 To cnt - 1
              arrjong(m) = stockobj.GetData(0, m) & " " & stockobj.GetData(1, m)
              lstjong.AddItem arrjong(m)
          Next m
          lblshow.Caption = "찾을 종목명 입력"
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set stockobj = Nothing
End Sub
Private Sub lstjong_DblClick()
            If Tabjong.SelectedItem.Index = 1 Then
               Form1.txtjongmok = Left(CStr(lstjong.List(lstjong.ListIndex)), 7)
               jongmoksearch.Hide
            Else
                Exit Sub
            End If
End Sub
Private Sub txtjongword_Change()
          Dim buffer As String
          buffer = txtjongword.Text
         If Tabjong.SelectedItem.Index = 1 Then
               For i = 0 To lstjong.ListCount
                If Mid(lstjong.List(i), 8, Len(buffer)) = UCase(buffer) Then
                    lstjong.ListIndex = i
                Exit For
                End If
               Next i
         End If
End Sub
Private Sub txtjongword_KeyDown(KeyCode As Integer, Shift As Integer)
             If Tabjong.SelectedItem.Index = 1 Then
                 If KeyCode = 13 Then
                    Form1.txtjongmok = Left(CStr(lstjong.List(lstjong.ListIndex)), 6)
                    jongmoksearch.Hide
                 ElseIf KeyCode = 38 Then
                     If lstjong.ListIndex > 0 Then
                        lstjong.ListIndex = lstjong.ListIndex - 1
                     End If
                     txtjongword.Text = lstjong.Text
                 ElseIf KeyCode = 40 Then
                    If lstjong.ListIndex < lstjong.ListCount - 1 Then
                    lstjong.ListIndex = lstjong.ListIndex + 1
                    End If
                    txtjongword.Text = lstjong.Text
                 Else
                 End If
             End If
End Sub

