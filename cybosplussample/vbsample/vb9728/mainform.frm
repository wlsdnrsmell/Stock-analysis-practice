VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MainForm 
   Caption         =   "옵션 정보 조회"
   ClientHeight    =   6105
   ClientLeft      =   4140
   ClientTop       =   2070
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   7935
   Begin VB.CommandButton cmd_exit 
      Caption         =   "종료"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483639
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10186
      TabWidthStyle   =   2
      TabFixedWidth   =   6879
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "기본 정보 조회"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "이론가 산출변수 조회"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oinfo_obj As New OptionInfo
Private otv_obj As New OptionTV
Sub init_optiontv()
    otv_obj.BlockRequest
    MSFlexGrid1.Cols = otv_obj.Data.Count
    MSFlexGrid1.Rows = otv_obj.GetHeaderValue(1) + 1

    MSFlexGrid1.FormatString = "<                     만기월|<               배당액지수현재가|<              잔존일수|<             이자율(%)"
    MSFlexGrid1.Redraw = False
    For i = 0 To otv_obj.GetHeaderValue(1) - 1
        For j = 0 To otv_obj.Data.Count - 1
            Select Case j
            Case 0:
                MSFlexGrid1.ColAlignment(j) = 4
                MSFlexGrid1.TextMatrix(i + 1, j) = Left(otv_obj.GetDataValue(j, i), 4) + "-" + Mid(otv_obj.GetDataValue(j, i), 5, 2)
            Case 1:
                MSFlexGrid1.ColAlignment(j) = 4
                MSFlexGrid1.TextMatrix(i + 1, j) = Format(otv_obj.GetDataValue(j, i), "#0.000000") + "  "
            Case 2:
                MSFlexGrid1.ColAlignment(j) = flexAlignRightCenter
                MSFlexGrid1.TextMatrix(i + 1, j) = otv_obj.GetDataValue(j, i)
            Case Else:
                MSFlexGrid1.ColAlignment(j) = 4
                MSFlexGrid1.TextMatrix(i + 1, j) = Format(otv_obj.GetDataValue(j, i), "#0.000") + "  "
            End Select
        Next
    Next

    MSFlexGrid1.Redraw = True
End Sub
Sub init_optioninfo()
    oinfo_obj.BlockRequest
    
    MSFlexGrid1.Cols = oinfo_obj.Data.Count
    MSFlexGrid1.Rows = oinfo_obj.GetHeaderValue(1) + 1

    MSFlexGrid1.FormatString = "<      옵션코드|<            종목명|<   행사가격|<       만기일|<      상한 |<       하한 |>   호가단위"
    MSFlexGrid1.Redraw = False
    For i = 0 To oinfo_obj.GetHeaderValue(1) - 1
        For j = 0 To oinfo_obj.Data.Count - 1
            Select Case j
            Case 3:
                MSFlexGrid1.ColAlignment(j) = 4
                MSFlexGrid1.TextMatrix(i + 1, j) = Left(oinfo_obj.GetDataValue(j, i), 4) + "-" + Mid(oinfo_obj.GetDataValue(j, i), 5, 2)
            Case 2, 4, 5, 6:
                MSFlexGrid1.ColAlignment(j) = flexAlignRightCenter
                MSFlexGrid1.TextMatrix(i + 1, j) = Format(oinfo_obj.GetDataValue(j, i), "#0.00") + "  "
            Case Else:
                MSFlexGrid1.ColAlignment(j) = 4
                MSFlexGrid1.TextMatrix(i + 1, j) = oinfo_obj.GetDataValue(j, i)
            End Select
        Next
    Next
    MSFlexGrid1.Redraw = True

End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set oinfo_obj = New OptionInfo
    Set otv_obj = New OptionTV
    Call init_optioninfo
    MSFlexGrid1.AllowBigSelection = True
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oinfo_obj = Nothing
    Set otv_obj = Nothing
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem = "기본 정보 조회" Then
        MSFlexGrid1.Clear
        Call init_optioninfo
    Else
        MSFlexGrid1.Clear
        Call init_optiontv
    End If
End Sub
