VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   120
      TabIndex        =   9
      Text            =   "24"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   120
      TabIndex        =   8
      Text            =   "1"
      Top             =   240
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetKosdaqIndustry2List "
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetKosdaqIndustry1List "
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetMemberList "
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetIndustryList "
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetGroupCodeList"
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetStockCodeListByMarket"
      Height          =   615
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Result"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Dim codemgr As New CpCodeMgr
    Dim codes
    
    Combo1.Clear
        
    Select Case (Index)
        Case 0
        codes = codemgr.GetStockListByMarket(CInt(Text1.Text))
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Debug.Print codemgr.CodeToName(codes(i))
            
            Combo1.AddItem codes(i) + "_" + codemgr.CodeToName(codes(i))
        Next
        
        Case 1
        codes = codemgr.GetGroupCodeList(CInt(Text2.Text))
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Combo1.AddItem codes(i) + "_" + codemgr.CodeToName(codes(i))
        Next
        
        Case 2
        codes = codemgr.GetIndustryList()
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Debug.Print codemgr.GetIndustryName(codes(i))
            
            Combo1.AddItem codes(i) + "_" + codemgr.GetIndustryName(codes(i))
        Next
        
        Case 3
        codes = codemgr.GetMemberList()
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Debug.Print codemgr.GetMemberName(codes(i))
            
            Combo1.AddItem codes(i) + "_" + codemgr.GetMemberName(codes(i))
        Next
        
        Case 4
        codes = codemgr.GetKosdaqIndustry1List()
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Debug.Print codemgr.GetIndustryName(codes(i))
            
            Combo1.AddItem codes(i) + "_" + codemgr.GetIndustryName(codes(i))
        Next
        
        Case 5
        codes = codemgr.GetKosdaqIndustry2List()
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Debug.Print codemgr.GetIndustryName(codes(i))
            
            Combo1.AddItem codes(i) + "_" + codemgr.GetIndustryName(codes(i))
        Next
    End Select
    
    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    End If
        
End Sub

