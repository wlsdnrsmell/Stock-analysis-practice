; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CTestDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "test.h"

ClassCount=3
Class1=CTestApp
Class2=CTestDlg

ResourceCount=3
Resource2=IDD_TEST_DIALOG
Resource1=IDR_MAINFRAME
Class3=CApple
Resource3=IDD_DIALOG1

[CLS:CTestApp]
Type=0
HeaderFile=test.h
ImplementationFile=test.cpp
Filter=N

[CLS:CTestDlg]
Type=0
HeaderFile=testDlg.h
ImplementationFile=testDlg.cpp
Filter=D
BaseClass=CDialog
VirtualFilter=dWC
LastObject=IDC_LEAP



[DLG:IDD_TEST_DIALOG]
Type=1
Class=CTestDlg
ControlCount=13
Control1=IDC_LEAP,edit,1350631552
Control2=IDC_GRAPH,button,1342242816
Control3=IDC_JONGMOK,listbox,1352728833
Control4=IDC_STOCK,button,1342242816
Control5=IDC_OPTION,button,1342242816
Control6=IDC_TAB1,SysTabControl32,1342177280
Control7=IDC_STATIC,button,1342177287
Control8=IDC_PROING,msctls_progress32,1350565888
Control9=IDC_DEGREE,static,1350696960
Control10=IDC_TIME,static,1350696960
Control11=IDC_TREE,SysTreeView32,1350631424
Control12=IDC_LEAPYEAR,button,1342242817
Control13=IDC_YOON,static,1342308352

[DLG:IDD_DIALOG1]
Type=1
Class=CApple
ControlCount=2
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816

[CLS:CApple]
Type=0
HeaderFile=Apple.h
ImplementationFile=Apple.cpp
BaseClass=CDialog
Filter=D
LastObject=CApple
VirtualFilter=dWC

