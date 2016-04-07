// testDlg.cpp : implementation file
//
/* 2001.5.31 (last version)
   cybosplus object (cpstockcode,cpfuturecode,cbgraph1를 이용하여 dialog based 으로 짬)
   초를 인자로 받어 년,일,시간,분,초로 display 하게 함
	by lee dong hee */

#include "stdafx.h"
#include "test.h"
#include "testDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTestDlg dialog

CTestDlg::CTestDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTestDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTestDlg)
	m_leap = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
//	m_stock.LoadBitmaps(IDB_BUP, IDB_BDOWN,IDB_BFOCUS,IDB_BDISABLE); 
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}
CTestDlg::~CTestDlg()
{
	delete m_pImageList;
		
}

void CTestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTestDlg)
	DDX_Control(pDX, IDC_LEAP, m_year);
	DDX_Control(pDX, IDC_LEAPYEAR, m_show);
	DDX_Control(pDX, IDC_TREE, m_tree);
	DDX_Control(pDX, IDC_TIME, m_time);
	DDX_Control(pDX, IDC_DEGREE, m_degree);
	DDX_Control(pDX, IDC_PROING, m_percent);
	DDX_Control(pDX, IDC_STOCK, m_stock);
	DDX_Control(pDX, IDC_OPTION, m_option);
	DDX_Control(pDX, IDC_GRAPH, m_graph);
	DDX_Control(pDX, IDC_TAB1, m_tab);
	DDX_Control(pDX, IDC_JONGMOK, m_list);
	DDX_Text(pDX, IDC_LEAP, m_leap);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CTestDlg, CDialog)
	//{{AFX_MSG_MAP(CTestDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_OPTION, OnOption)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, OnSelchangeTab1)
	ON_BN_CLICKED(IDC_GRAPH, OnGraph)
	ON_BN_CLICKED(IDC_STOCK, OnStock)
	ON_BN_CLICKED(IDC_LEAPYEAR, OnLeapyear)
	ON_EN_CHANGE(IDC_LEAP, OnChangeLeap)
	//}}AFX_MSG_MAP


END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTestDlg message handlers



// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CTestDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
	
	
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CTestDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

BOOL CTestDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	m_tree.ShowWindow(SW_SHOW);
//	m_tree.EnableToolTips(TRUE);
	m_list.ShowWindow(SW_HIDE);
	m_stock.EnableWindow(TRUE);
	m_option.EnableWindow(TRUE);
	m_graph.EnableWindow(TRUE);
	m_show.EnableWindow(FALSE);
	
	m_pImageList = new CImageList;
	m_pImageList->Create(11,11,ILC_COLOR4,3,0);	
	CBitmap bm1,bm2,bm3;
	bm1.LoadBitmap(IDB_BUP);
	bm2.LoadBitmap(IDB_BDOWN);
	bm3.LoadBitmap(IDB_BFOCUS);
	
	m_pImageList->Add(&bm1,RGB(0,0,0));
	m_pImageList->Add(&bm2,RGB(0,0,0));
	m_pImageList->Add(&bm3,RGB(0,0,0));
	m_tab.SetImageList(m_pImageList);
	
	char szTabItems[3][20] = {  _T("그래프"),
								_T("주식종목 보기"),
								_T("선물종목 보기"),
	};
	TCITEM tcItem;
	tcItem.mask = TCIF_TEXT | TCIF_IMAGE;
	for (int i = 0; i <3; i++)
	{
		tcItem.pszText = szTabItems[i];
		tcItem.iImage = i;
		m_tab.InsertItem(i, &tcItem);
	}
	
		
	
	m_tab.SetCurSel(1);

	m_tooltip.Create(this);
	m_tooltip.Activate(TRUE);
	m_tooltip.AddTool(GetDlgItem(IDC_TREE), "Tree Tool Tip"); 

	CEdit* pedit = (CEdit*) GetDlgItem(IDC_LEAP);
	pedit->SetFocus();
	
	return FALSE;  
}
void CTestDlg::OnOption() 
{
	m_tree.ShowWindow(SW_HIDE);
	m_list.ShowWindow(SW_SHOW);
	m_list.ResetContent();
	m_tab.SetCurSel(2);
	m_stock.EnableWindow(TRUE);
	m_option.EnableWindow(FALSE);
	m_graph.EnableWindow(TRUE);
	CString b,b1,b3;
	int futuresu;
	ICpFutureCodePtr Fut;
	Fut.CreateInstance(__uuidof(CpFutureCode));
	futuresu = Fut->GetCount();
	
	for(int i = 0; i<futuresu; i++)
	{
		b.Empty();
		b1.Empty();
		b3.Empty();
		b1 = (LPCTSTR)(_bstr_t)Fut->GetData(0,i);
		b3= (LPCTSTR)(_bstr_t)Fut->GetData(1,i);
		b.Format("%s    %s",b1,b3);
		m_list.InsertString(0,b);
	}
	Fut.Release();
}

void CTestDlg::OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	UpdateData(TRUE);
	int m = m_tab.GetCurSel();
	switch(m)
	{
	case 0:
		OnGraph();
		break;
	case 1:
		OnStock();
		break;
	case 2:
		OnOption();
	}

	*pResult = 0;
}



void CTestDlg::OnGraph() 
{
	m_list.ShowWindow(TRUE);
	m_tree.ShowWindow(FALSE);
	DWORD diff;
	CString strJongMok,str2,dis,strname,rest,display;
	int su,i;
	m_list.ResetContent();
	m_tab.SetCurSel(0);
	m_stock.EnableWindow(TRUE);
	m_option.EnableWindow(TRUE);
	m_graph.EnableWindow(FALSE);
	IDibPtr dib;
	ICpStockCodePtr util;
	dib.CreateInstance(__uuidof(CbGraph1));
	util.CreateInstance(__uuidof(CpStockCode));
	su = util->GetCount();
	m_percent.SetPos(0);
	m_percent.SetRange(0,su);
	m_percent.SetStep(1);
	DWORD dwStart = GetTickCount();

    long val[6];
	for( i=0 ; i<su; i++)
	{

		m_percent.StepIt();
		strJongMok = (LPCSTR)(_bstr_t)util->GetData(0,i);
		strname= (LPCTSTR)(_bstr_t)util->GetData(1,i);
		dis.Format("%d/%d이 진행되고 있습니다\n(%s) 종목 데이터를 받고 있는 중입니다",i+1,su,strname);
		m_degree.SetWindowText(dis);
		dib->SetInputValue(0, (LPCTSTR)strJongMok);
		dib->SetInputValue(1, (BYTE)'D');
		CString temp;
		BYTE t=dib->GetInputValue(1);
		temp.Format("%c",(BYTE)dib->GetInputValue(1));		
		temp.Format("%d",dib->GetInputValue(1));
		dib->SetInputValue(2, 0L);
		dib->SetInputValue(3, (short)10);
		dib->BlockRequest();
		
		short count = dib->GetHeaderValue(3); 
		for(int j = 0; j<count; j++)
		{
			val[0] = dib->GetDataValue (0,j);
			val[1] = dib->GetDataValue (1,j);
			val[2] = dib->GetDataValue (2,j);
			val[3] = dib->GetDataValue (3,j);
			val[4] = dib->GetDataValue (4,j);
			val[5] = dib->GetDataValue (5,j);
			str2.Format("(%d )%d : %d %d %d %d %d", j, val[0], val[1],val[2],val[3],val[4],val[5]);
			
			m_list.InsertString(0, str2);
		}
	
		diff = DWORD((GetTickCount() - dwStart)/1000); 
		rest = converttime(diff);
		display.Format("경과된 시간은 %s입니다",rest);
		m_time.SetWindowText(display);
		
		
	}
	MessageBox("끝");
	dib.Release();
}

void CTestDlg::OnStock() 
{

	m_tree.ShowWindow(SW_SHOW);
	m_tree.DeleteAllItems();
	m_list.ShowWindow(SW_HIDE);
	m_tab.SetCurSel(1);
	m_stock.EnableWindow(FALSE);
	m_option.EnableWindow(TRUE);
	m_graph.EnableWindow(TRUE);
	
	CTreeCtrl* pCtrl = (CTreeCtrl*) GetDlgItem(IDC_TREE);
	TVINSERTSTRUCT tvInsert;
	tvInsert.hParent = NULL;
	tvInsert.hInsertAfter = NULL;
	tvInsert.item.mask = TVIF_TEXT;
	tvInsert.item.pszText = _T("주식종목");
	
	HTREEITEM hStock = pCtrl->InsertItem(&tvInsert);
	tooltipMap[hStock] = "주식종목";
	HTREEITEM hga = pCtrl->InsertItem(TVIF_TEXT,
		_T("거래소시장"), 0, 0, 0, 0, 0, hStock, NULL);
	tooltipMap[hga] = "거래소시장";
	HTREEITEM hko = pCtrl->InsertItem(_T("코스닥시장"),
		0, 0, hStock,NULL);
	tooltipMap[hko] = "코스닥시장";
	HTREEITEM hmarket = pCtrl->InsertItem(_T("제3시장"),
		0, 0, hStock,NULL);
	tooltipMap[hmarket] = "제3시장";
	
	
	
	CString a,a1,a3,a4;
	char buf[2];
	memset(buf,0,2);
	int su;
	ICpStockCodePtr util;
	util.CreateInstance(__uuidof(CpStockCode));
	su = util->GetCount();
	
	for(int i = 0; i<su; i++)
	{
		a.Empty();
		a1.Empty();
		a3.Empty();
		a1 = (LPCTSTR)(_bstr_t)util->GetData(0,i);
		a3= (LPCTSTR)(_bstr_t)util->GetData(1,i);
		a4 = (LPCTSTR)(_bstr_t)util->GetData(4,i);
		strncpy(buf,a4,1);
		HTREEITEM hItem;
		switch (buf[0])
		{
		case '1':
			a.Format("%s    %s",a1,a3);
			hItem = pCtrl->InsertItem(_T(a), hga, TVI_SORT);
			break;
		case '5':
			a.Format("%s    %s",a1,a3);
			hItem = pCtrl->InsertItem(_T(a), hko, TVI_SORT);
			break;
		case '6':
			a.Format("%s    %s",a1,a3);
			hItem = pCtrl->InsertItem(_T(a), hmarket, TVI_SORT);
			break;
			
		}
		tooltipMap[hItem] = a;   
		
	} 
	util.Release();
}

CString CTestDlg::converttime(DWORD a)
{
	CString m,mi,ho,da,ye;
	long b,c;
	if ( a< 60)
		m.Format("%d초",a);
	else
	{
		if ( a < 3600)
		{
			b = (long)a/60;
			c = (long)a%60;
            mi = converttime(c);
			m.Format("%d분 %s",b,mi);
		}
		else
		{
			if ( a < 86400)
			{
				b = (long)a/3600;
				c = (long)a%3600;
				ho = converttime(c);
				m.Format("%d시간 %s",b,ho);
			}
			else
			{
				if ( a < 86400* 365)
				{
					b = (long) a/86400;
					c = (long) a%86400;
					da = converttime(c);
					m.Format("%d일 %s",b,da);
				}
				else
				{
					b = (long) a/(86400*365);
					c = (long) a%(86400*365);
					ye = converttime(c);
					m.Format("%d년 %s",b,ye);
				}
			}
		}
		
	}
	return m;			   
}

BOOL CTestDlg::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	if(pMsg->message == WM_MOUSEMOVE && pMsg->hwnd == m_tree.m_hWnd)
	{
		CPoint point(LOWORD(pMsg->lParam),HIWORD(pMsg->lParam));
		HTREEITEM hItem = m_tree.HitTest(point);
		if(hItem != NULL)
		{
			CString text = tooltipMap[hItem];				
			m_tooltip.UpdateTipText(text,&m_tree);
			m_tooltip.RelayEvent(pMsg);
		}
		
	}
	else if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN && pMsg->hwnd == m_year.GetSafeHwnd())
	     OnLeapyear();


	return FALSE;
}

void CTestDlg::OnLeapyear() 
{
	int m;
	UpdateData(TRUE);
	m = atoi(m_leap);
	if (((m%4) == 0) && ((m%100) != 0) || ((m% 400)==0)) 
		SetDlgItemText(IDC_YOON,"윤년");
	else
		SetDlgItemText(IDC_YOON,"평년");
}

void CTestDlg::OnChangeLeap() 
{
      UpdateData(TRUE);
	  if (m_leap.GetLength() >= 1) 
		  m_show.EnableWindow(TRUE);
	  else
		  m_show.EnableWindow(FALSE);
	
}
