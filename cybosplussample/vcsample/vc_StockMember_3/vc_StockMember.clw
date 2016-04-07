; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CVc_StockMemberDlg
LastTemplate=CCmdTarget
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "vc_StockMember.h"

ClassCount=4
Class1=CVc_StockMemberApp
Class2=CVc_StockMemberDlg
Class3=CAboutDlg

ResourceCount=3
Resource1=IDD_ABOUTBOX
Resource2=IDR_MAINFRAME
Class4=CCpDibEvent
Resource3=IDD_VC_STOCKMEMBER_DIALOG

[CLS:CVc_StockMemberApp]
Type=0
HeaderFile=vc_StockMember.h
ImplementationFile=vc_StockMember.cpp
Filter=N

[CLS:CVc_StockMemberDlg]
Type=0
HeaderFile=vc_StockMemberDlg.h
ImplementationFile=vc_StockMemberDlg.cpp
Filter=D
LastObject=IDC_BTN_REQUEST2
BaseClass=CDialog
VirtualFilter=dWC

[CLS:CAboutDlg]
Type=0
HeaderFile=vc_StockMemberDlg.h
ImplementationFile=vc_StockMemberDlg.cpp
Filter=D
LastObject=CAboutDlg

[DLG:IDD_ABOUTBOX]
Type=1
Class=CAboutDlg
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889

[DLG:IDD_VC_STOCKMEMBER_DIALOG]
Type=1
Class=CVc_StockMemberDlg
ControlCount=4
Control1=IDOK,button,1342242816
Control2=IDC_EDT_JONGMOK,edit,1350631560
Control3=IDC_BTN_REQUEST,button,1342242817
Control4=IDC_EDT_OUT,edit,1352730756

[CLS:CCpDibEvent]
Type=0
HeaderFile=CpDibEvent.h
ImplementationFile=CpDibEvent.cpp
BaseClass=CCmdTarget
Filter=N

