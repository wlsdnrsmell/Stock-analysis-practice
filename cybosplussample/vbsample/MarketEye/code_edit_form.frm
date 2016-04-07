VERSION 5.00
Begin VB.Form code_edit_form 
   BorderStyle     =   1  '단일 고정
   Caption         =   "종목검색"
   ClientHeight    =   4740
   ClientLeft      =   4260
   ClientTop       =   2310
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7245
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   27
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<<-"
      Height          =   495
      Left            =   3360
      TabIndex        =   26
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txt_cnt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<-"
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "->"
      Height          =   495
      Left            =   3360
      TabIndex        =   21
      Top             =   1200
      Width           =   495
   End
   Begin VB.ListBox list_code 
      Height          =   2400
      Left            =   3960
      MultiSelect     =   2  '확장형
      TabIndex        =   20
      Top             =   960
      Width           =   3135
   End
   Begin VB.ListBox lst_jongmok 
      Height          =   2400
      Left            =   120
      MultiSelect     =   2  '확장형
      TabIndex        =   19
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txt_jongmok 
      Height          =   375
      IMEMode         =   10  '한글 
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "취  소"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "확  인"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   4200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   375
         Index           =   15
         Left            =   2640
         TabIndex        =   28
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "A"
         Height          =   375
         Index           =   14
         Left            =   2280
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "아"
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "하"
         Height          =   375
         Index           =   13
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "파"
         Height          =   375
         Index           =   12
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "타"
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "카"
         Height          =   375
         Index           =   10
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "차"
         Height          =   375
         Index           =   9
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "자"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "사"
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "바"
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "마"
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "라"
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "다"
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "나"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "가"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      Caption         =   "개"
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "선택된 종목 수 "
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "code_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cpcodelist As CpStockCode
Private name_sort As Boolean
Dim name_sortcode() As String
Dim num_sortcode() As String
'임시로 코드그룹을 저장하는 배열(편집시 그룹을 오간 후 취소버튼을 클릭하면 저장 하지 말아야 하기때문에....)
Private tmp_code_group(10) As String
'그룹이 변경되면 이전 tmp_group을 저장하기위한 매개체.
Public before_codes As String
Public before_index As Integer
Public chang_flag As Boolean
Sub set_before_codes()
    before_codes = ""
    For i = 0 To list_code.ListCount - 1
        list_code.ListIndex = i
        before_codes = before_codes + Left(list_code.Text, 6)
    Next
    Call set_tmp_code_group(before_codes, before_index)
    chang_flag = False
End Sub
Public Function get_tmp_code_group(i)
    tmpstr = tmp_code_group(i)
    get_tmp_code_group = tmpstr
End Function
Sub set_tmp_code_group(value, index)
    tmp_code_group(index) = value
End Sub
Sub tmp_code_savecancel()
    For i = 0 To 9
        tmp_code_group(i) = main_form.get_code_group(i) ' 임시그룹 저장
        Call load_edit_codelist(tmp_code_group(i))
    Next
End Sub

Sub load_edit_codelist(value)
i = 0
cnt = 0
list_code.Clear
Do
    c = i + 1
    i = InStr(c, value, "A")
    If i > 0 Then
      tmp = Mid(value, i, 6)
      list_code.AddItem tmp + "     " + cpcodelist.CodeToName(tmp)
      cnt = cnt + 1
    Else
        Exit Do
    End If
Loop
txt_cnt = cnt
End Sub

Sub sort_codelist()
    Set cpcodelist = New CpStockCode
    ReDim name_sortcode(cpcodelist.GetCount) As String
    ReDim num_sortcode(cpcodelist.GetCount) As String
    
    For i = 0 To cpcodelist.GetCount - 1
       num_sortcode(i) = cpcodelist.GetData(0, i) + "     " + cpcodelist.GetData(1, i)
       name_sortcode(i) = cpcodelist.GetData(0, i) + "     " + cpcodelist.GetData(1, i)
    Next i
     '이름 순이면 sort를 하자...
    For i = 0 To cpcodelist.GetCount - 1
        For j = i + 1 To cpcodelist.GetCount - 1
            If Mid(name_sortcode(i), 11) > Mid(name_sortcode(j), 11) Then
                TEMP = name_sortcode(i)
                name_sortcode(i) = name_sortcode(j)
                name_sortcode(j) = TEMP
            End If
        Next j
    Next i
    
    Call add_sortcode_tolist
End Sub
Private Sub add_sortcode_tolist()
    lst_jongmok.Clear
    If name_sort Then
        For i = 0 To cpcodelist.GetCount - 1
            lst_jongmok.AddItem name_sortcode(i)
        Next i
    Else
        For i = 0 To cpcodelist.GetCount - 1
            lst_jongmok.AddItem num_sortcode(i)
        Next i
    End If

End Sub

Private Sub cmd_save_Click()
Call set_before_codes
Call main_form.save_code
Call main_form.set_code_byindex(Combo1.ListIndex)
main_form.Combo1.ListIndex = Combo1.ListIndex
Hide
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex >= 0 Then
        If chang_flag = True Then
            Call set_before_codes
        End If
        list_code.Clear
        load_edit_codelist (get_tmp_code_group(Combo1.ListIndex))
    End If
        before_index = Combo1.ListIndex
End Sub

Private Sub Command1_Click(index As Integer)
If index < 15 And name_sort = False Then
    name_sort = True
    'MsgBox ("이름순 정렬")
    Call add_sortcode_tolist
ElseIf index = 15 And name_sort = True Then
    name_sort = False
    'MsgBox ("코드순 정렬")
    Call add_sortcode_tolist
End If

For i = 0 To lst_jongmok.ListCount - 1
    tmp = Mid(lst_jongmok.List(i), 12)
    
    If tmp > Command1(index).Caption Then
        lst_jongmok.TopIndex = i
        lst_jongmok.Selected(i) = True
        Exit Sub
    End If
Next i
End Sub
Private Sub Command3_Click()
For i = 0 To lst_jongmok.ListCount - 1
    If lst_jongmok.Selected(i) = True Then
        lst_jongmok.Selected(i) = False
    End If
Next

For i = 0 To list_code.ListCount - 1
    If list_code.Selected(i) = True Then
        list_code.Selected(i) = False
    End If
Next
Call tmp_code_savecancel
Hide
End Sub

Private Sub Command4_Click()
Dim exist_flag As Boolean

If lst_jongmok.SelCount <> 0 Then
    For i = 0 To lst_jongmok.ListCount - 1
        If list_code.ListCount >= 50 Then
            MsgBox ("50개까지만 가능합니다")
            Exit For 'Sub
        End If
        
        If lst_jongmok.Selected(i) = True Then
            exist_flag = False
            For k = 0 To list_code.ListCount - 1
                If Left(lst_jongmok.List(i), 6) = Left(list_code.List(k), 6) Then
                    MsgBox Mid(lst_jongmok.List(i), 11) + "종목은 이미 등록되어있습니다."
                    exist_flag = True
                End If
            Next
            If (exist_flag <> True) Then
                list_code.AddItem lst_jongmok.List(i)
                txt_cnt.Text = list_code.ListCount
                chang_flag = True
            End If
            lst_jongmok.Selected(i) = False
        End If
     Next i
End If

For i = 0 To list_code.ListCount - 1
    If list_code.Selected(i) = True Then
        list_code.Selected(i) = False
    End If
Next
End Sub

Private Sub Command5_Click()
If list_code.ListIndex = -1 Then
    MsgBox ("삭제할 종목을 선택하여 주십시오")
    Exit Sub
End If
    For i = list_code.ListCount - 1 To 0 Step -1
        If list_code.Selected(i) = True Then
            list_code.RemoveItem i
            chang_flag = True
        End If
    Next i
txt_cnt.Text = list_code.ListCount
End Sub

Private Sub Command6_Click()
For i = list_code.ListCount - 1 To 0 Step -1
    list_code.RemoveItem i
    chang_flag = True
Next
txt_cnt.Text = 0
End Sub



Private Sub Form_Load()
chang_flag = False
name_sort = False

For i = 700 To 709
    Combo1.AddItem Str(i) + " " + "포트폴리오"
Next
Combo1.ListIndex = 0
before_index = Combo1.ListIndex
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cpcodelist = Nothing
End Sub

Private Sub list_code_DblClick()
If list_code.ListIndex = -1 Then
    Exit Sub
End If
list_code.RemoveItem list_code.ListIndex
chang_flag = True
txt_cnt.Text = list_code.ListCount
End Sub

Private Sub lst_jongmok_DblClick()
If list_code.ListCount >= 50 Then
    MsgBox ("50개까지만 가능합니다")
    Exit Sub
End If
For i = 1 To list_code.ListCount
    list_code.ListIndex = i - 1
    If Left(lst_jongmok.Text, 6) = Left(list_code.Text, 6) Then
        MsgBox Mid(lst_jongmok.Text, 11) + "종목은 이미 등록되어있습니다."
        Exit Sub
    End If
Next

If lst_jongmok.Text <> "" Then
    list_code.AddItem lst_jongmok.Text
    chang_flag = True
    list_code.ListIndex = list_code.ListCount - 1
    txt_cnt.Text = list_code.ListCount
End If
End Sub

Private Sub txt_jongmok_Change()
If txt_jongmok.Text = "" Then
    Exit Sub
End If
'리스트 속도개선을 위해.
If IsNumeric(Mid(txt_jongmok.Text, 1)) And Len(txt_jongmok.Text) <> 5 Then
    Exit Sub
End If

For i = 0 To lst_jongmok.ListCount - 1
    If IsNumeric(Mid(txt_jongmok.Text, 1)) And Len(txt_jongmok.Text) = 5 Then '코드
       If name_sort = True Then
            name_sort = False
            Call add_sortcode_tolist
        End If
        
        tmp = Mid(lst_jongmok.List(i), 2)
    Else '이름
       If name_sort = False Then
            name_sort = True
            Call add_sortcode_tolist
        End If

        tmp = Mid(lst_jongmok.List(i), 12)
    End If
    
    If InStr(tmp, txt_jongmok.Text) = 1 Then
        lst_jongmok.TopIndex = i
        Exit Sub
    End If
Next i
End Sub
Private Sub txt_jongmok_Click()
txt_jongmok.Text = ""
End Sub
