VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form main_form 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2895
         Left            =   80
         TabIndex        =   6
         Top             =   150
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   5106
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   350
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "관심종목편집"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "종   료"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      BackColorSel    =   12640511
      ForeColorSel    =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents sb_obj As StockCur
Attribute sb_obj.VB_VarHelpID = -1

Public m_cnt As Integer 'sb_code 배열 갯수
Private sb_code(50) As String  'sb하기위한 코드배열
Private code_group(10) As String '코드그룹을 저장하는 배열

Sub set_tree()
               
          Dim nodx As Node
          With TreeView1
         .LineStyle = tvwRootLines
          Set nodx = .Nodes.Add(, , "up", "관심종목")
                  
          For i = 0 To 99
             Set nodx = .Nodes.Add("up", tvwChild, "upup" & i, i + 700)
                 
          Next i
          
           End With
End Sub

Public Function get_code_group(i)
    tmpstr = code_group(i)
    get_code_group = tmpstr
End Function
Sub set_code_byvalue(i, value)
    sb_code(i) = value
End Sub
Sub set_code_byindex(index)
'If index > -1 Then Combo1.ListIndex = index

i = 0
j = 0
Do
    c = j + 1
    j = InStr(c, code_group(index), "A")
    If j Then
        If i < 50 Then sb_code(i) = Mid(code_group(index), j, 6)
        i = i + 1
    Else
        Exit Do
    End If
Loop
MSFlexGrid1.Redraw = False
set_cnt (i)
code_rq
code_sb
MSFlexGrid1.Redraw = True
End Sub
Sub load_code()
    Dim tmpcode As String
    i = -1
    Open App.Path & "\code1.txt" For Input As #1

    Do While Not EOF(1)
        Input #1, tmpcode
        If tmpcode <> "" Then '그룹단위로 읽음.
            i = i + 1
            code_group(i) = Mid(tmpcode, 3) '그룹저장
        End If
    Loop
    Close #1

    For i = 0 To 9
         Call code_edit_form.set_tmp_code_group(code_group(i), i)
    Next
    '리스트박스에 저장된 코드나열되게...
    code_edit_form.before_codes = code_edit_form.get_tmp_code_group(i)
    Call code_edit_form.load_edit_codelist(code_edit_form.before_codes)
End Sub
Sub save_code()
    Open App.Path & "\code1.txt" For Output As #1
    For i = 0 To 9
        code_group(i) = code_edit_form.get_tmp_code_group(i)
        Print #1, "<>" + code_group(i) + ","
    Next
    Close #1
End Sub
Sub set_cnt(value)
    m_cnt = value
    MSFlexGrid1.Rows = value + 1
End Sub
 
Sub code_rq()
'"|종목코드|종목명|현재가|대비|대비(%)|매도호가|매수호가|누적거래량|누적거래대금"
    Me.MousePointer = 11
    MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = "|종목코드 |종목명          |현재가    |대비     |대비(%)   |매도호가 |매수호가 |누적거래량  |누적거래대금   "

    Dim rq_obj As New StockMst
    
    For i = 0 To m_cnt - 1
        rq_obj.SetInputValue 0, sb_code(i)
        ret = rq_obj.BlockRequest
     '   MsgBox (Str(ret))
        If ret <> 0 Then
            Set rq_obj = Nothing
            MsgBox (Str(ret) + "통신상태가 비정상입니다.HTS종료후 재접속하여 주십시오")
            Exit Sub
        End If
        tmp = rq_obj.GetHeaderValue(0)
    '    MsgBox (tmp)
        MSFlexGrid1.TextMatrix(i + 1, 1) = rq_obj.GetHeaderValue(0)
        MSFlexGrid1.TextMatrix(i + 1, 2) = rq_obj.GetHeaderValue(1)
        MSFlexGrid1.TextMatrix(i + 1, 3) = Format(rq_obj.GetHeaderValue(11), "###,###") '현재가
        
        Call set_color(i + 1, 4, rq_obj.GetHeaderValue(12))
        tmp = rq_obj.GetHeaderValue(12)
        If tmp > 0 Then
            tmp = "+"
        Else
            tmp = ""
        End If
        MSFlexGrid1.TextMatrix(i + 1, 4) = tmp + Format(rq_obj.GetHeaderValue(12), "###,###") '대비
        
        If rq_obj.GetHeaderValue(11) > 0 Then
        tmp = (rq_obj.GetHeaderValue(11) - rq_obj.GetHeaderValue(10)) / rq_obj.GetHeaderValue(10) * 100
        Call set_color(i + 1, 5, tmp)
        c = InStr(1, tmp, ".")
        tmp = Mid(tmp, 1, c + 2)
            If tmp > 0 Then
            tmp = "+" + Format(CCur(tmp), "0.00")
            MSFlexGrid1.TextMatrix(i + 1, 5) = tmp '대비(%)
            Else
                MSFlexGrid1.TextMatrix(i + 1, 5) = Format(tmp, "0.00")
            End If
              
        MSFlexGrid1.TextMatrix(i + 1, 6) = Format(rq_obj.GetHeaderValue(16), "###,###")
        MSFlexGrid1.TextMatrix(i + 1, 7) = Format(rq_obj.GetHeaderValue(17), "###,###")
        MSFlexGrid1.TextMatrix(i + 1, 8) = Format(rq_obj.GetHeaderValue(18), "###,###")
        MSFlexGrid1.TextMatrix(i + 1, 9) = Format(rq_obj.GetHeaderValue(19), "###,###")
        End If
    Next
    If m_cnt > 0 Then
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 1

        MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    End If
    Set rq_obj = Nothing
    Me.MousePointer = 0
End Sub
Sub code_sb()
sb_obj.SetInputValue 0, ""
sb_obj.Unsubscribe
    For i = 0 To m_cnt - 1
        sb_obj.SetInputValue 0, sb_code(i)
        sb_obj.SubscribeLatest
    Next
End Sub
Private Sub Combo1_Click()
    If Combo1.ListIndex < 0 Then Exit Sub
    set_code_byindex (Combo1.ListIndex)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Frame1.Visible Then
    Frame1.Visible = False
    Command3.Caption = "+"
End If
If Combo1.ListIndex < 0 Then '업종 계열사일경우
    code_edit_form.Combo1.ListIndex = 0
    code_edit_form.load_edit_codelist (get_code_group(0))
Else
    code_edit_form.Combo1.ListIndex = Combo1.ListIndex
    code_edit_form.load_edit_codelist (get_code_group(Combo1.ListIndex))
End If
code_edit_form.txt_cnt.Text = code_edit_form.list_code.ListCount

code_edit_form.lst_jongmok.TopIndex = 0
code_edit_form.lst_jongmok.ListIndex = -1
code_edit_form.list_code.TopIndex = 0
code_edit_form.list_code.ListIndex = -1

code_edit_form.Show
End Sub


Private Sub Command3_Click()
If Command3.Caption = "+" Then
    Command3.Caption = "-"
    Frame1.Visible = True
Else
    Command3.Caption = "+"
    Frame1.Visible = False
End If
Combo1.ListIndex = -1
End Sub

Private Sub Form_Load()
'    wait_form.Show
 Me.MousePointer = 11
 
    Call set_tree
    Call code_edit_form.sort_codelist
        
    MSFlexGrid1.Cols = 10
    MSFlexGrid1.FormatString = "|종목코드 |종목명          |현재가    |대비     |대비(%)   |매도호가 |매수호가 |누적거래량  |누적거래대금   "
        
    Set sb_obj = New StockCur

    Call load_code
    Call set_code_byindex(0)

    For i = 700 To 709
        Combo1.AddItem Str(i) + " " + "포트폴리오"
    Next
    Combo1.ListIndex = 0 '내부적으로 click을 호출함.
    
    'Unload wait_form
Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
    MSFlexGrid1.Height = 6135 / 7230 * main_form.Height
    MSFlexGrid1.Width = 9255 / 9585 * main_form.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set sb_obj = Nothing
    Unload code_edit_form
End Sub
Sub set_selcolor(i, j, value)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = j
    If Left(value, 1) = "+" Then
        MSFlexGrid1.ForeColorSel = vbRed
    ElseIf Left(value, 1) = "-" Then
        MSFlexGrid1.ForeColorSel = vbBlue
    Else
        MSFlexGrid1.ForeColorSel = vbBlack
    End If
End Sub
Sub set_color(i, j, value)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = j
    If CCur(value) > 0 Then
        MSFlexGrid1.CellForeColor = vbRed
'        MSFlexGrid1.ForeColorSel = vbRed
    ElseIf CCur(value) < 0 Then
        MSFlexGrid1.CellForeColor = vbBlue
 '       MSFlexGrid1.ForeColorSel = vbBlue
    Else
        MSFlexGrid1.CellForeColor = vbBlack
  '      MSFlexGrid1.ForeColorSel = vbBlack
    End If
End Sub
Private Sub sb_obj_Received()
'"|종목코드|종목명|현재가|대비|대비(%)|저가|매도호가|매수호가|누적거래량|누적거래대금"
    
    For i = 0 To m_cnt - 1
        If MSFlexGrid1.TextMatrix(i, 1) = sb_obj.GetHeaderValue(0) Then
            Exit For
        End If
    Next
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Redraw = False
    
    MSFlexGrid1.TextMatrix(i, 3) = Format(sb_obj.GetHeaderValue(13), "###,###")  '현재가
    
    Call set_color(i, 4, sb_obj.GetHeaderValue(2))
    tmp = sb_obj.GetHeaderValue(2)
        If tmp > 0 Then
            tmp = "+"
        Else
            tmp = ""
        End If
    MSFlexGrid1.TextMatrix(i, 4) = tmp + Format(sb_obj.GetHeaderValue(2), "###,###") '전일대비
    
    If sb_obj.GetHeaderValue(13) Then
        yesterday = sb_obj.GetHeaderValue(13) - sb_obj.GetHeaderValue(2)
        tmp = (sb_obj.GetHeaderValue(13) - yesterday) / yesterday * 100
        Call set_color(i, 5, tmp) '대비(%)
        c = InStr(1, tmp, ".")
        tmp = Mid(tmp, 1, c + 2)
        If tmp > 0 Then
            tmp = "+" + Format(CCur(tmp), "0.00")
            MSFlexGrid1.TextMatrix(i, 5) = tmp
        Else
            MSFlexGrid1.TextMatrix(i, 5) = Format(tmp, "0.00")
        End If
    End If
    MSFlexGrid1.TextMatrix(i, 6) = Format(sb_obj.GetHeaderValue(7), "###,###")  '매도호가
    MSFlexGrid1.TextMatrix(i, 7) = Format(sb_obj.GetHeaderValue(8), "###,###")  '매수호가
    MSFlexGrid1.TextMatrix(i, 8) = Format(sb_obj.GetHeaderValue(9), "###,###")  '누적거래량
    MSFlexGrid1.TextMatrix(i, 9) = Format(sb_obj.GetHeaderValue(10), "###,###")  '누적거래대금

    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    
    MSFlexGrid1.Redraw = True
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If IsNumeric(Left(Node.Text, 1)) Then
    Dim value, cnt
    cnt = 0
    Dim stockobj As New CpStockCode
    
tmp = Left(Node.Text, 1)
'계열사 클릭 시.
If Left(Node.Text, 1) = "9" Then
    For i = 0 To stockobj.GetCount - 1
        If Mid(Node.Text, 2, 2) = stockobj.GetData(8, i) Then
            Call set_code_byvalue(cnt, stockobj.GetData(0, i))
            cnt = cnt + 1
            If cnt = 50 Then
                MsgBox ("표시종목 갯수 50개까지만 표시됩니다")
                Exit For
            End If
        End If
    Next i
'포트폴리오
ElseIf Left(Node.Text, 1) = "7" Then
        Call set_code_byindex(CInt(Mid(Node.Text, 3, 1)))
        Set stockobj = Nothing
        Combo1.Text = Node.Text
        Frame1.Visible = False
        Command3.Caption = "+"
        Exit Sub
'업종 클릭 시
Else
    For i = 0 To stockobj.GetCount - 1
        If Left(Node.Text, 3) = stockobj.GetData(3, i) Then
            Call set_code_byvalue(cnt, stockobj.GetData(0, i))
            cnt = cnt + 1
            If cnt = 50 Then
                MsgBox ("표시종목 갯수 50개까지만 표시됩니다")
                Exit For
            End If
        End If
    Next i
End If

    MSFlexGrid1.Redraw = False
    set_cnt (cnt)
    code_rq
    code_sb
    Set stockobj = Nothing
    Combo1.Text = Node.Text
    Frame1.Visible = False
    Command3.Caption = "+"
    MSFlexGrid1.Redraw = True
End If
End Sub
