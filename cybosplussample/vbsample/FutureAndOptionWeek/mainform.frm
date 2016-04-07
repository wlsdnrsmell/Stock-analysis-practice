VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mainform 
   Caption         =   "선물&옵션 6주간"
   ClientHeight    =   9120
   ClientLeft      =   1005
   ClientTop       =   720
   ClientWidth     =   11640
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   11640
   Begin VB.TextBox O_textbox 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "종목선택"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "옵션"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "선물"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   14420
      _Version        =   393216
      ForeColor       =   4210752
      AllowUserResizing=   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "종     료"
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox F_combo 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public futurelist As CpFutureCode
Public optionlist As CpOptionCode
Public WithEvents futureWeek As FutureWeek1
Attribute futureWeek.VB_VarHelpID = -1
Public WithEvents optWeek As OptionWeek
Attribute optWeek.VB_VarHelpID = -1

Private Sub Command2_Click()
optionsearch.Show
End Sub

Private Sub F_combo_Change()
'MsgBox ("change")
If Len(F_combo.Text) < 5 Then Exit Sub

Call fill_F_data

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub F_combo_Click()
'MsgBox ("click")
If Len(F_combo.Text) < 5 Then Exit Sub


Call fill_F_data

End Sub

Private Sub Form_Load()

    Option1(0).Value = True
    
    Set futurelist = New CpFutureCode
    Set optionlist = New CpOptionCode
    Set futureWeek = New FutureWeek1
    Set optWeek = New OptionWeek
    
    For i = 0 To futurelist.GetCount - 1
        F_combo.AddItem futurelist.GetData(0, i)
    Next i
    F_combo.Text = futurelist.GetData(0, 0)
    
    Call set_optionjongmok

    Call fill_F_data
End Sub

Private Sub set_optionjongmok()
    optionsearch.Button(0).Caption = optionlist.GetData(3, 0)
    
    For i = 1 To optionlist.GetCount - 1
    If optionsearch.Button(0).Caption <> optionlist.GetData(3, i) Then
       optionsearch.Button(1).Caption = optionlist.GetData(3, i)
       Exit For
    End If
    Next i
    
    For j = i To optionlist.GetCount - 1
    If optionsearch.Button(1).Caption <> optionlist.GetData(3, j) Then
       optionsearch.Button(2).Caption = optionlist.GetData(3, j)
       Exit For
    End If
    Next j
    
    
    For k = j To optionlist.GetCount - 1
    If optionsearch.Button(2).Caption <> optionlist.GetData(3, k) Then
       optionsearch.Button(3).Caption = optionlist.GetData(3, k)
       Exit For
    End If
    Next k
    
    '리스트 채우기

For i = 0 To optionlist.GetCount - 1
    If optionlist.GetData(2, i) = "풋" Then
        Exit For
    End If
    optionsearch.ListBox1.AddItem optionlist.GetData(0, i) + "   " + optionlist.GetData(1, i)
Next i

For j = i To optionlist.GetCount - 1
    optionsearch.ListBox2.AddItem optionlist.GetData(0, j) + "   " + optionlist.GetData(1, j)
Next j
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set futurelist = Nothing
    Set optionlist = Nothing
    Set futureWeek = Nothing
    Set optWeek = Nothing
    Unload optionsearch
        
End Sub

Private Sub fill_F_data()
MSFlexGrid1.Clear
    futureWeek.SetInputValue 0, F_combo.Text
    futureWeek.BlockRequest
    
End Sub
Private Sub fill_O_data()
MSFlexGrid1.Clear
    optWeek.SetInputValue 0, O_textbox
    optWeek.BlockRequest
End Sub
Private Sub fill_O_title()
tmpstr = "<  일자|<    시가|<    고가|<    저가|<    종가|<     대비|<  거래량   |<  미결제 |<  이론가 |<  변동성|<      IV|<   Delta|<   Gamma|<   Theta|<     vega|<      rho|>누적거래대금"

'tmpstr = "<"
 '   Set b = optWeek.Data
  '  n = b.Count
   ' For i = 1 To n - 2
    '    Set c = b.Item(i)
     '   tmpstr = tmpstr + Right(c.Name, 5) + "|<"
    'Next
     '   Set c = b.Item(n - 1)
      '  tmpstr = tmpstr + Right(c.Name, 5) + "|>"
    
       ' Set c = b.Item(n)
       ' tmpstr = tmpstr + Right(c.Name, 5)
    
    MSFlexGrid1.FormatString = tmpstr
End Sub

Private Sub futureWeek_Received()
MSFlexGrid1.Redraw = False

MSFlexGrid1.Cols = futureWeek.Data.Count
MSFlexGrid1.Rows = futureWeek.GetHeaderValue(0) + 1
Call fill_F_title
For i = 1 To futureWeek.GetHeaderValue(0)
         For j = 0 To futureWeek.Data.Count - 1 '11개
            MSFlexGrid1.ColAlignment(j) = flexAlignRightCenter
            MSFlexGrid1.Col = j
            MSFlexGrid1.Row = i
            MSFlexGrid1.CellForeColor = vbBlack
            If j = 0 Then
                MSFlexGrid1.TextMatrix(i, j) = Right(futureWeek.GetDataValue(j, i - 1), 4)
            ElseIf j >= 1 And j <= 4 Or j = 9 Then
                MSFlexGrid1.TextMatrix(i, j) = Format(futureWeek.GetDataValue(j, i - 1), "##0.00")
            ElseIf j = 5 Or j = 10 Then
                If futureWeek.GetDataValue(j, i - 1) < 0 Then
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = i
                    MSFlexGrid1.CellForeColor = vbBlue
                ElseIf futureWeek.GetDataValue(j, i - 1) = 0 Then
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = i
                    MSFlexGrid1.CellForeColor = vbBlack
                Else
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = i
                    MSFlexGrid1.CellForeColor = vbRed
                End If
                MSFlexGrid1.TextMatrix(i, j) = Format(futureWeek.GetDataValue(j, i - 1), "##0.00")
                
            ElseIf j >= 6 And j <= 8 Then
                MSFlexGrid1.TextMatrix(i, j) = Format(futureWeek.GetDataValue(j, i - 1), "###,###,###")
            Else
                MSFlexGrid1.TextMatrix(i, j) = Str(futureWeek.GetDataValue(j, i - 1))
            End If
       Next
    Next
MSFlexGrid1.Redraw = True

End Sub

Private Sub fill_F_title()
'"<  날짜     |<  시가   |<  고가   |<  저가   |<  종가   |<  대비   |<  거래량   |<  거래대금 |<  미결제약정 |<이론선물|>베이시스 "
    tmpstr = "<"
    Set b = futureWeek.Data
    n = b.Count
    For i = 1 To n - 2
        Set c = b.Item(i)
        tmpstr = tmpstr + c.Name + "    |<"
    Next
        Set c = b.Item(n - 1)
        tmpstr = tmpstr + c.Name + "|>"
    
        Set c = b.Item(n)
        tmpstr = tmpstr + c.Name
    MSFlexGrid1.FormatString = tmpstr
End Sub

Private Sub O_textbox_Change()
If Len(O_textbox.Text) < 8 Then Exit Sub


Call fill_O_data

End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0) = True Then
        F_combo.Enabled = True
        O_textbox.Enabled = False
        Command2.Enabled = False
        
    Else
        F_combo.Enabled = False
        O_textbox.Enabled = True
        Command2.Enabled = True
    End If
    
End Sub

Private Sub optWeek_Received()
MSFlexGrid1.Redraw = False
MSFlexGrid1.Cols = optWeek.Data.Count
MSFlexGrid1.Rows = optWeek.GetHeaderValue(0) + 1

Call fill_O_title

For i = 1 To optWeek.GetHeaderValue(0)
         For j = 0 To optWeek.Data.Count - 1
            MSFlexGrid1.ColAlignment(j) = flexAlignRightCenter
            MSFlexGrid1.Col = j
            MSFlexGrid1.Row = i
            MSFlexGrid1.CellForeColor = vbBlack
            If j = 0 Then
                 MSFlexGrid1.TextMatrix(i, j) = Right(optWeek.GetDataValue(j, i - 1), 4)
            ElseIf j >= 1 And j <= 4 Or j >= 8 And j <= 12 Then
                 MSFlexGrid1.TextMatrix(i, j) = Format(optWeek.GetDataValue(j, i - 1), "##0.00")
            ElseIf j = 5 Then
                If optWeek.GetDataValue(j, i - 1) < 0 Then
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = i
                    MSFlexGrid1.CellForeColor = vbBlue
                ElseIf optWeek.GetDataValue(j, i - 1) = 0 Then
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = i
                    MSFlexGrid1.CellForeColor = vbBlack
                Else
                    MSFlexGrid1.Col = j
                    MSFlexGrid1.Row = i
                    MSFlexGrid1.CellForeColor = vbRed
                End If
                MSFlexGrid1.TextMatrix(i, j) = Format(optWeek.GetDataValue(j, i - 1), "##0.00")
                
            ElseIf j >= 6 And j < 8 Or j = 16 Then
                MSFlexGrid1.TextMatrix(i, j) = Format(optWeek.GetDataValue(j, i - 1), "###,###,###")
            ElseIf j >= 13 And j < 16 Then
                MSFlexGrid1.TextMatrix(i, j) = Format(optWeek.GetDataValue(j, i - 1), "##0.0000")

            Else
                'MSFlexGrid1.TextMatrix(i, j) = Str(optWeek.GetDataValue(j, i - 1))
            End If
       Next
    Next
MSFlexGrid1.Redraw = True
End Sub
