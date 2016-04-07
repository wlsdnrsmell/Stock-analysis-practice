VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   Caption         =   "회원사 매매현황"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "검색"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin MSComctlLib.ListView lvwcode 
      Height          =   1695
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   4210752
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "↓"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtmember 
      Height          =   270
      Left            =   480
      TabIndex        =   2
      Text            =   "888"
      Top             =   240
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox txtjongmok 
      Height          =   270
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "코드"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "회원"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private memberobj As CpSvr8091
Private WithEvents membersobj As CpSvr8091S
Attribute membersobj.VB_VarHelpID = -1
Private Sub Command1_Click()
            lvwcode.Visible = True
End Sub
Private Sub Command2_Click()
            jongmoksearch.Show
End Sub

Private Sub Form_Load()
            lvwcode.Visible = False
            Set membersobj = New CpSvr8091S
            
            lvwcode.ColumnHeaders.Add , , "코드", lvwcode.Width() / 3
            lvwcode.ListItems.Add 1, , "888"
             
            Dim CodeMgr As New CpCodeMgr
            Dim codes As Variant
            codes = CodeMgr.GetMemberList()
            For i = LBound(codes) To UBound(codes)
                lvwcode.ListItems.Add i + 2, , codes(i)
            Next
            
End Sub
Private Sub Form_Unload(Cancel As Integer)
            Set membersobj = Nothing
            Unload jongmoksearch
End Sub
Private Sub lvwcode_Click()
            txtmember.Text = lvwcode.SelectedItem.Text
            lvwcode.Visible = False
End Sub
Private Sub membersobj_Received()
            strs = membersobj.GetHeaderValue(0)
            strs = strs & "   " & membersobj.GetHeaderValue(1)
            strs = strs & "   " & membersobj.GetHeaderValue(3)
            bigos = membersobj.GetHeaderValue(4)
            If Chr(bigos) = "1" Then
            strs = strs & "   " & "매도"
            Else
            strs = strs & "   " & "매수"
            End If
            strs = strs & "   " & membersobj.GetHeaderValue(5)
            strs = strs & "   " & membersobj.GetHeaderValue(6)
            List1.AddItem strs, List1.TopIndex
End Sub

Private Sub txtjongmok_Change()
        If Len(txtjongmok.Text) = 7 Then
        doit txtmember.Text, txtjongmok.Text
        End If
End Sub

Private Sub txtjongmok_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           doit txtmember.Text, txtjongmok.Text
        End If
End Sub
Private Sub doit(str As Variant, temp As Variant)
        List1.Clear
        Set memberobj = New CpSvr8091
        If Len(temp) = 6 Then
            If StrComp(str, "888", vbTextCompare) = 0 Then
            memberobj.SetInputValue 0, Asc("5")
            Else
            memberobj.SetInputValue 0, Asc("4")
            End If
        Else
            If StrComp(str, "888", vbTextCompare) = 0 Then
            memberobj.SetInputValue 0, Asc("1")
            Else
            memberobj.SetInputValue 0, Asc("3")
            End If
        End If
        memberobj.SetInputValue 1, str
        memberobj.SetInputValue 2, temp
        If lstseq <> 0 Then
        memberobj.SetInputValue 3, lstseq
        End If
        memberobj.BlockRequest
        lstseq = memberobj.GetHeaderValue(1)
        For i = 0 To memberobj.GetHeaderValue(0) - 1
            strr = memberobj.GetDataValue(0, i)
            strr = strr & "   " & memberobj.GetDataValue(1, i)
            strr = strr & "   " & memberobj.GetDataValue(3, i)
            bigo = memberobj.GetDataValue(4, i)
            If Chr(bigo) = "1" Then
            strr = strr & "   " & "매도"
            Else
            strr = strr & "   " & "매수"
            End If
            strr = strr & "   " & memberobj.GetDataValue(5, i)
            strr = strr & "   " & memberobj.GetDataValue(6, i)
            List1.AddItem strr
        Next
        'Do While memberobj.Continue
        'memberobj.BlockRequest
       ' For I = 0 To memberobj.GetHeaderValue(0) - 1
       '     strr = memberobj.GetDataValue(0, I)
       '     strr = strr & "   " & memberobj.GetDataValue(1, I)
       '     strr = strr & "   " & memberobj.GetDataValue(3, I)
       '     bigo = memberobj.GetDataValue(4, I)
       '     If Chr(bigo) = "1" Then
       '     strr = strr & "   " & "매도"
       '     Else
       '     strr = strr & "   " & "매수"
       '     End If
       '     strr = strr & "   " & memberobj.GetDataValue(5, I)
       '     strr = strr & "   " & memberobj.GetDataValue(6, I)
       '     List1.AddItem strr
        'Next
        'Loop
        membersobj.SetInputValue 0, ""
        membersobj.Unsubscribe
        membersobj.SetInputValue 0, str
        If Len(temp) = 0 Then
        membersobj.SetInputValue 1, "*"
        Else
        membersobj.SetInputValue 1, temp
        End If
        membersobj.Subscribe
End Sub



