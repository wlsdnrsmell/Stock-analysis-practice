// StockCurDlgDlg.cpp : implementation file
//

#include "stdafx.h"
#include "StockCurDlg.h"
#include "StockCurDlgDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//추가->
_ATL_FUNC_INFO ReceivedInfo = { CC_STDCALL, VT_EMPTY, 0, { VT_EMPTY } };
//추가->

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CStockCurDlgDlg dialog

CStockCurDlgDlg::CStockCurDlgDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CStockCurDlgDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CStockCurDlgDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CStockCurDlgDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CStockCurDlgDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CStockCurDlgDlg, CDialog)
	//{{AFX_MSG_MAP(CStockCurDlgDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, OnTest)
	ON_BN_CLICKED(IDC_BUTTON2, OnTest2)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CStockCurDlgDlg message handlers

BOOL CStockCurDlgDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
//추가->

	HRESULT hr;
	IUnknownPtr pUnk;

	CComPtr<CPUTILLib::ICpOptionCode> pOptionCode;
	//hr = pOptionCode.CoCreateInstance(__uuidof(CPUTILLib::CpOptionCode)); 
	hr = pOptionCode.CoCreateInstance(CPUTILLib::CLSID_CpOptionCode ); 
	
	if (FAILED(hr))
			_com_raise_error(hr);
	int n = pOptionCode->GetCount();
	try {
		hr = m_pStockCur.CreateInstance(DSCBO1Lib::CLSID_StockCur);
		if (FAILED(hr))
			_com_raise_error(hr);

		hr = m_pCpSvr8092S.CreateInstance(DSCBO1Lib::CLSID_CpSvr8092S);
		if (FAILED(hr))
			_com_raise_error(hr);		

		hr = IDispEventSimpleImpl<MYSTOCKCUR_EVENT_ID, CStockCurDlgDlg, &__uuidof(DSCBO1Lib::_IDibEvents)>::DispEventAdvise(m_pStockCur);
		if (FAILED(hr))
			_com_raise_error(hr);

		hr = IDispEventSimpleImpl<MYCPSVR8092S_EVENT_ID, CStockCurDlgDlg, &__uuidof(DSCBO1Lib::_IDibEvents)>::DispEventAdvise(m_pCpSvr8092S);
		if (FAILED(hr))
			_com_raise_error(hr);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
//추가-<
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CStockCurDlgDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CStockCurDlgDlg::OnPaint() 
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
HCURSOR CStockCurDlgDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CStockCurDlgDlg::OnTest() 
{
	// TODO: Add your control notification handler code here
	m_pCpSvr8092S->SetInputValue (0,"*");
	m_pCpSvr8092S->Subscribe();

	//대신증권을 등록한다. 
	m_pStockCur->SetInputValue (0,"A003540");
	m_pStockCur->Subscribe();

	//추가로 하이닉스도 등록한다. 
	m_pStockCur->SetInputValue (0,"A000660");
	m_pStockCur->Subscribe();
}

void CStockCurDlgDlg::OnTest2() 
{
	// TODO: Add your control notification handler code here
	//시세수신을 안한다. 
	m_pCpSvr8092S->Unsubscribe (); 

	//대신증권 수신을 이제 더이상 안받겠다... 
	m_pStockCur->SetInputValue (0,"A003540");
	m_pStockCur->Unsubscribe();

	//하이닉스 수신을 이제 더이상 안받겠다...
	m_pStockCur->SetInputValue (0,"A000660");
	m_pStockCur->Unsubscribe();
}
//추가->
void __stdcall CStockCurDlgDlg::OnMyStockCurReceived()
{
	TRACE("\nOnMyStockCurReceived");
	//등록한 대신증권, 하이닉스 등등등 등록한 종목이 모두 이곳으로 수신된다. 
	//무슨 종목이 수신온건지는 m_pStockCur->GetHeaderValue(0) 종목코드로 비교한다. 
	TRACE("\n%s 현재가 %d, ",(LPCTSTR)(_bstr_t) m_pStockCur->GetHeaderValue(0),(long)(_variant_t)m_pStockCur->GetHeaderValue(13)); 

}
void __stdcall CStockCurDlgDlg::OnMyCpSvr8092SReceived()
{
	TRACE("\nOnMyCpSvr8092SReceived");
	TRACE("\n%s, %s",(LPCTSTR)(_bstr_t) m_pCpSvr8092S->GetHeaderValue(1), (LPCTSTR)(_bstr_t) m_pCpSvr8092S->GetHeaderValue(5));  //5 - (string) 내용
}
//추가-<


