VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   Caption         =   "TR 7031(전업종지수)"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "allupjongjisu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "allupjongjisu.frx":00EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "코스닥"
      Height          =   375
      Left            =   1920
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "장    내"
      Height          =   375
      Left            =   360
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '단일 고정
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents indexobj As StockIndexis
Attribute indexobj.VB_VarHelpID = -1
Private Sub Command1_Click()
        ListView1.ColumnHeaders.Clear
        ListView1.ListItems.Clear
        Dim hs As New CbGraph1
        
        
        Label1.Caption = "장내 지수"
        ListView1.ColumnHeaders.Add , , "업종코드", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "업종명", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "지수", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "대비", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "등락율", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "거래량", ListView1.Width() / 6
        
        
        Dim CodeMgr As New CpCodeMgr
        Dim codes As Variant
        codes = CodeMgr.GetIndustryList()
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Debug.Print CodeMgr.GetIndustryName(codes(i))

         With hs
             codetemp = codes(i)
             nametemp = CodeMgr.GetIndustryName(codes(i))
             
             .SetInputValue 0, "U00" & codetemp
             .SetInputValue 1, Asc("D")
             .SetInputValue 2, 0
             .SetInputValue 3, 1
             .BlockRequest
              For n = 1 To .GetHeaderValue(3)
                    tmpjisu = FormatNumber(.GetDataValue(4, n - 1), 2)  '지수
                    tmp2 = .GetDataValue(4, n - 1)                     '지수
                    tmp = .GetDataValue(6, n - 1)                       '대비
                    tstr = showpercent(tmp2, tmp)                       '등락율
                    If tmp > 0 Then
                    tmpdaebi = FormatNumber(.GetDataValue(6, n - 1), 2)  '대비
                    Else
                    tmpdaebi = Mid(FormatNumber(.GetDataValue(6, n - 1), 2), 2)
                    End If
                    tmpvolume = FormatNumber(.GetDataValue(5, n - 1), 0) '거래량
                    
                    Set JList = ListView1.ListItems.Add(n + i, , codetemp)
                    
                    JList.SubItems(1) = nametemp
                    JList.SubItems(2) = tmpjisu
                    JList.SubItems(3) = tmpdaebi
                    JList.SubItems(4) = tstr
                    JList.SubItems(5) = tmpvolume
              Next n
              
              If tmp > 0 Then
                  ListView1.ListItems(i + 1).ListSubItems(3).ReportIcon = 1
              Else
                  ListView1.ListItems(i + 1).ListSubItems(3).ReportIcon = 2
              End If
              
              If Left(tstr, 1) = "-" Then
              ListView1.ListItems(i + 1).ListSubItems(4).ForeColor = RGB(0, 0, 255)
              Else
              ListView1.ListItems(i + 1).ListSubItems(4).ForeColor = RGB(255, 0, 0)
              End If
              
        End With
        Next i
        indexobj.SetInputValue 0, ""
        indexobj.Unsubscribe
        indexobj.SetInputValue 0, "*"
        indexobj.SubscribeLatest
        
        Set hs = Nothing
End Sub
Private Sub Command2_Click()
        ListView1.ColumnHeaders.Clear
        ListView1.ListItems.Clear
        Dim hs As New CbGraph1
        
         
        Label1.Caption = "코스닥 지수"
        ListView1.ColumnHeaders.Add , , "업종코드", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "업종명", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "지수", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "대비", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "등락율", ListView1.Width() / 6
        ListView1.ColumnHeaders.Add , , "거래량", ListView1.Width() / 6
                  
        Dim CodeMgr As New CpCodeMgr
        Dim codes As Variant
        codes = CodeMgr.GetKosdaqIndustry1List()
        For i = LBound(codes) To UBound(codes)
            Debug.Print codes(i)
            Debug.Print CodeMgr.GetIndustryName(codes(i))
         With hs
         
             codetemp = codes(i)
             nametemp = CodeMgr.GetIndustryName(codes(i))
             .SetInputValue 0, "U00" & codetemp
             .SetInputValue 1, Asc("D")
             .SetInputValue 2, 0
             .SetInputValue 3, 1
             .BlockRequest
            For n = 1 To .GetHeaderValue(3)
                tmpjisu = FormatNumber(.GetDataValue(4, n - 1), 2)  '지수
                tmp2 = .GetDataValue(4, n - 1)                     '지수
                tmp = .GetDataValue(6, n - 1)                       '대비
                tstr = showpercent(tmp2, tmp)                       '등락율
                If tmp > 0 Then
                    tmpdaebi = FormatNumber(.GetDataValue(6, n - 1), 2)  '대비
                Else
                    tmpdaebi = Mid(FormatNumber(.GetDataValue(6, n - 1), 2), 2)
                End If
                
                tmpvolume = FormatNumber(.GetDataValue(5, n - 1), 0) '거래량
                Set JList = ListView1.ListItems.Add(n + i, , codetemp)
                 
                JList.SubItems(1) = nametemp
                JList.SubItems(2) = tmpjisu
                JList.SubItems(3) = tmpdaebi
                JList.SubItems(4) = tstr
                JList.SubItems(5) = tmpvolume
              Next n
              
              If tmp > 0 Then
                  ListView1.ListItems(i + 1).ListSubItems(3).ReportIcon = 1
              Else
                  ListView1.ListItems(i + 1).ListSubItems(3).ReportIcon = 2
              End If
              
              If Left(tstr, 1) = "-" Then
              ListView1.ListItems(i + 1).ListSubItems(4).ForeColor = RGB(0, 0, 255)
              Else
              ListView1.ListItems(i + 1).ListSubItems(4).ForeColor = RGB(255, 0, 0)
              End If
              
        End With
        Next i
        indexobj.SetInputValue 0, ""
        indexobj.Unsubscribe
        indexobj.SetInputValue 0, "*"
        indexobj.SubscribeLatest
        Set hs = Nothing
        Set cpccode = Nothing
        
End Sub
Private Sub Form_Load()
        Set indexobj = New StockIndexis
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set indexobj = Nothing
End Sub
Private Sub indexobj_Received()
            Dim com1, com2, sbstr, daebi
            For vv = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(vv).Text = indexobj.GetHeaderValue(0) Then
            com1 = indexobj.GetHeaderValue(2)
            com2 = indexobj.GetHeaderValue(3)
            If com2 > 0 Then
            daebi = FormatNumber(indexobj.GetHeaderValue(3), 2)
            Else
            daebi = Mid(FormatNumber(indexobj.GetHeaderValue(3), 2), 2)
            End If
            sbstr = showpercent(com1, com2)
            ListView1.ListItems(vv).SubItems(2) = com1
                If com2 > 0 Then
                  ListView1.ListItems(vv).ListSubItems(3).ReportIcon = 1
                Else
                  ListView1.ListItems(vv).ListSubItems(3).ReportIcon = 2
                End If
               
                If Left(sbstr, 1) = "-" Then
                ListView1.ListItems(vv).ListSubItems(4).ForeColor = RGB(0, 0, 255)
                Else
                ListView1.ListItems(vv).ListSubItems(4).ForeColor = RGB(255, 0, 0)
                End If
               
            ListView1.ListItems(vv).SubItems(3) = daebi
            ListView1.ListItems(vv).SubItems(4) = sbstr
            ListView1.ListItems(vv).SubItems(5) = FormatNumber(indexobj.GetHeaderValue(4), 0)
            Exit For
            End If
            Next
End Sub
Private Function showpercent(a As Variant, b As Variant) As String
' 지수와 대비값을 받어서 string으로 출력(소수점 이하 2자리)
    Dim temp As Variant
    Dim strtemp As String
    temp = (b / (a - b)) * 100
    strtemp = FormatNumber(temp, 3)
        c = InStr(1, strtemp, ".")
        strtemp = Mid(strtemp, 1, c + 2)
    showpercent = strtemp
End Function
