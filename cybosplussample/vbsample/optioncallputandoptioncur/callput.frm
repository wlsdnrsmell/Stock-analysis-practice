VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "call/put ��ȸ"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cmopt 
      Height          =   300
      ItemData        =   "callput.frx":0000
      Left            =   1080
      List            =   "callput.frx":0002
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   4210752
      ForeColorFixed  =   4210752
      ForeColorSel    =   65535
      GridColor       =   8421631
      GridColorFixed  =   8421631
      FillStyle       =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'1.optioncallput object�� �̿��Ͽ� (MSFlexGrid Control�� �̿�) �ɼ� ������ ���簡 ��ȸ ȭ��
'�ۼ�(rq/rp)
'by leedonghee 2001-06-20

'2.optioncur object(sb/pb)�� �̿� �ɼ� ������ ���簡 ��ȸ ȭ��
'update -by leedonghee 2001-06-22

'3.20010814 �ɼǿ��� ã�⸦ ���α׷� ������ ó���ϱ� ���ؼ�(�������� ���� �޺� �ڽ��� �־��־���)
'����� �����ϴ� ..
'loop�� ���鼭 ���� �ִ� ���� �����ϱ�
'���� ������ �޺� �ڽ��� �߰��Ѵ�
Public WithEvents callputobj As OptionCallput
Attribute callputobj.VB_VarHelpID = -1
Public WithEvents omsbobj As OptionCur
Attribute omsbobj.VB_VarHelpID = -1
Dim t
Private Sub Command1_Click()
             callputobj.SetInputValue 0, cmopt.List(cmopt.ListIndex)
             callputobj.BlockRequest
             t = callputobj.GetHeaderValue(0)
             grid1.Rows = t + 1
             grid1.Cols = 12
             grid1.FixedRows = 1
             grid1.ColWidth(3) = 1000
             grid1.ColWidth(4) = 1000
             grid1.ColWidth(9) = 1000
             grid1.ColWidth(10) = 1200
             grid1.TextMatrix(0, 0) = "��簡"
             grid1.TextMatrix(0, 1) = "���簡"
             grid1.TextMatrix(0, 2) = "���"
             grid1.TextMatrix(0, 3) = "�ŵ�(�ܷ�)"
             grid1.TextMatrix(0, 4) = "�ż�(�ܷ�)"
             grid1.TextMatrix(0, 5) = "�ŷ���"
             grid1.TextMatrix(0, 6) = "��簡"
             grid1.TextMatrix(0, 7) = "���簡"
             grid1.TextMatrix(0, 8) = "���"
             grid1.TextMatrix(0, 9) = "�ŵ�(�ܷ�)"
             grid1.TextMatrix(0, 10) = "�ż�(�ܷ�)"
             grid1.TextMatrix(0, 11) = "�ŷ���"
           
 
            
             For m = 1 To t
                For k = 0 To 11
                     Select Case k
                     Case 6, 0
                     grid1.TextMatrix(m, k) = callputobj.GetDataValue(2, m - 1)
                     Case 5
                     grid1.TextMatrix(m, k) = Format(callputobj.GetDataValue(k + 3, m - 1), "###,###")
                     Case 11
                     grid1.TextMatrix(m, k) = Format(callputobj.GetDataValue(k + 5, m - 1), "###,###")
                     Case 2
                     CON = callputobj.GetDataValue(k + 7, m - 1)
                     If CON > 0 Then
                         grid1.Row = m
                         grid1.Col = 2
                         grid1.CellForeColor = RGB(255, 0, 0)
                         grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 7, m - 1)
                     ElseIf CON < 0 Then
                         grid1.Row = m
                         grid1.Col = 2
                         grid1.CellForeColor = RGB(0, 0, 255)
                         grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 7, m - 1)
                     Else
                         grid1.Row = m
                         grid1.Col = 2
                         grid1.CellForeColor = RGB(0, 0, 0)
                         grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 7, m - 1)
                     End If
                     
                     Case 8
                     CONPUT = callputobj.GetDataValue(k + 9, m - 1)
                     If CONPUT > 0 Then
                         grid1.Row = m
                         grid1.Col = 8
                         grid1.CellForeColor = RGB(255, 0, 0)
                         grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 9, m - 1)
                     ElseIf CONPUT < 0 Then
                         grid1.Row = m
                         grid1.Col = 8
                         grid1.CellForeColor = RGB(0, 0, 255)
                         grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 9, m - 1)
                     Else
                         grid1.Row = m
                         grid1.Col = 8
                         grid1.CellForeColor = RGB(0, 0, 0)
                         grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 9, m - 1)
                     End If
                     
                     Case 1
                     grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 2, m - 1)
                     Case 7
                     grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 4, m - 1)
                     
                     Case 3
                     grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 1, m - 1) & "(" & callputobj.GetDataValue(k + 2, m - 1) & ")"
                     Case 4
                     grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 2, m - 1) & "(" & callputobj.GetDataValue(k + 3, m - 1) & ")"
                     Case 9
                     grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 3, m - 1) & "(" & callputobj.GetDataValue(k + 4, m - 1) & ")"
                     Case 10
                     grid1.TextMatrix(m, k) = callputobj.GetDataValue(k + 4, m - 1) & "(" & callputobj.GetDataValue(k + 5, m - 1) & ")"
                     
                     End Select
                     
                Next
                     
            Next
   
          omsbobj.SetInputValue 0, ""
          omsbobj.Unsubscribe
          omsbobj.SetInputValue 0, "*"
          omsbobj.SubscribeLatest

                   
End Sub

Private Sub Form_Load()
        Set callputobj = New OptionCallput
        Set omsbobj = New OptionCur
        Set optcodeobj = New CpOptionCode
             
             
      
        
        Dim count, stoMonth, sMonth, i
        count = optcodeobj.GetCount()
        For i = 0 To (count / 2) - 1
           sMonth = optcodeobj.GetData(3, i)
           If i = 0 Then
               cmopt.AddItem sMonth
           Else
               If stoMonth <> sMonth Then
                  cmopt.AddItem sMonth
               End If
                  
           End If
        stoMonth = sMonth
        Next
        cmopt.ListIndex = 0
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
            Set callputobj = Nothing
            Set omsbobj = Nothing
End Sub
Private Sub omsbobj_Received()
            Dim p, sel
            s1 = omsbobj.GetHeaderValue(51)     ' ����
            s2 = callputobj.GetInputValue(0)    ' ����
            If s1 <> s2 Then Exit Sub
    
            s3 = omsbobj.GetHeaderValue(52) '��簡
        
            c = omsbobj.GetHeaderValue(0)
            For p = 1 To t
                If CDbl(grid1.TextMatrix(p, 0)) = s3 Then   '20010814 ����
                     sel = p
                Exit For
                End If
            Next
            
            If IsEmpty(sel) = True Then
              Exit Sub
            End If
            
            If Left(c, 1) = "2" Then
                grid1.TextMatrix(sel, 1) = omsbobj.GetHeaderValue(24) '���簡
                CONSB = omsbobj.GetHeaderValue(25)
                If CONSB > 0 Then
                     grid1.Row = sel
                     grid1.Col = 2
                     grid1.CellForeColor = RGB(255, 0, 0)
                     grid1.TextMatrix(sel, 2) = omsbobj.GetHeaderValue(25)  '���
                ElseIf CONSB < 0 Then
                     grid1.Row = sel
                     grid1.Col = 2
                     grid1.CellForeColor = RGB(0, 0, 255)
                     grid1.TextMatrix(sel, 2) = omsbobj.GetHeaderValue(25)  '���
                Else
                     grid1.Row = sel
                     grid1.Col = 2
                     grid1.CellForeColor = RGB(0, 0, 0)
                     grid1.TextMatrix(sel, 2) = omsbobj.GetHeaderValue(25)  '���
                End If
                grid1.TextMatrix(sel, 3) = omsbobj.GetHeaderValue(2) & "(" & omsbobj.GetHeaderValue(7) & ")"
                grid1.TextMatrix(sel, 4) = omsbobj.GetHeaderValue(13) & "(" & omsbobj.GetHeaderValue(18) & ")" '�ż�
                grid1.TextMatrix(sel, 5) = Format(omsbobj.GetHeaderValue(29), "###,###") '�ŷ���
               
            
            Else
            
                grid1.TextMatrix(sel, 7) = omsbobj.GetHeaderValue(24)
                CONPUTSB = omsbobj.GetHeaderValue(25)
                If CONPUTSB > 0 Then
                     grid1.Row = sel
                     grid1.Col = 8
                     grid1.CellForeColor = RGB(255, 0, 0)
                     grid1.TextMatrix(sel, 8) = omsbobj.GetHeaderValue(25)  '���
                ElseIf CONPUTSB < 0 Then
                     grid1.Row = sel
                     grid1.Col = 8
                     grid1.CellForeColor = RGB(0, 0, 255)
                     grid1.TextMatrix(sel, 8) = omsbobj.GetHeaderValue(25)  '���
                Else
                     grid1.Row = sel
                     grid1.Col = 8
                     grid1.CellForeColor = RGB(0, 0, 0)
                     grid1.TextMatrix(sel, 8) = omsbobj.GetHeaderValue(25)  '���
                End If
                 grid1.TextMatrix(sel, 9) = omsbobj.GetHeaderValue(2) & "(" & omsbobj.GetHeaderValue(7) & ")"
                 grid1.TextMatrix(sel, 10) = omsbobj.GetHeaderValue(13) & "(" & omsbobj.GetHeaderValue(18) & ")"
                 grid1.TextMatrix(sel, 11) = Format(omsbobj.GetHeaderValue(29), "###,###")
                 
            
            End If
End Sub
