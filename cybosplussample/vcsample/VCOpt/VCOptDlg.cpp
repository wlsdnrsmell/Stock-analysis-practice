// VCOptDlg.cpp : implementation file
//

#include "stdafx.h"
#include "VCOpt.h"
#include "VCOptDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

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
// CVCOptDlg dialog

CVCOptDlg::CVCOptDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CVCOptDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CVCOptDlg)
	m_strCode = _T("");
	m_strCountBuy1 = _T("");
	m_strCountBuy2 = _T("");
	m_strCountBuy3 = _T("");
	m_strCountBuy4 = _T("");
	m_strCountBuy5 = _T("");
	m_strCountSell1 = _T("");
	m_strCountSell2 = _T("");
	m_strCountSell3 = _T("");
	m_strCountSell4 = _T("");
	m_strCountSell5 = _T("");
	m_strPrice = _T("");
	m_strPriceBuy1 = _T("");
	m_strPriceBuy2 = _T("");
	m_strPriceBuy3 = _T("");
	m_strPriceBuy4 = _T("");
	m_strPriceBuy5 = _T("");
	m_strPriceSell1 = _T("");
	m_strPriceSell2 = _T("");
	m_strPriceSell3 = _T("");
	m_strPriceSell4 = _T("");
	m_strPriceSell5 = _T("");
	m_strUnsettled = _T("");
	m_strVolume = _T("");
	m_strVolumeBuy1 = _T("");
	m_strVolumeBuy2 = _T("");
	m_strVolumeBuy3 = _T("");
	m_strVolumeBuy4 = _T("");
	m_strVolumeBuy5 = _T("");
	m_strVolumeSell1 = _T("");
	m_strVolumeSell2 = _T("");
	m_strVolumeSell3 = _T("");
	m_strVolumeSell4 = _T("");
	m_strVolumeSell5 = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

	m_pOptionCur = NULL;
}

void CVCOptDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CVCOptDlg)
	DDX_Text(pDX, IDC_CODE, m_strCode);
	DDV_MaxChars(pDX, m_strCode, 8);
	DDX_Text(pDX, IDC_COUNT_BUY1, m_strCountBuy1);
	DDX_Text(pDX, IDC_COUNT_BUY2, m_strCountBuy2);
	DDX_Text(pDX, IDC_COUNT_BUY3, m_strCountBuy3);
	DDX_Text(pDX, IDC_COUNT_BUY4, m_strCountBuy4);
	DDX_Text(pDX, IDC_COUNT_BUY5, m_strCountBuy5);
	DDX_Text(pDX, IDC_COUNT_SELL1, m_strCountSell1);
	DDX_Text(pDX, IDC_COUNT_SELL2, m_strCountSell2);
	DDX_Text(pDX, IDC_COUNT_SELL3, m_strCountSell3);
	DDX_Text(pDX, IDC_COUNT_SELL4, m_strCountSell4);
	DDX_Text(pDX, IDC_COUNT_SELL5, m_strCountSell5);
	DDX_Text(pDX, IDC_PRICE, m_strPrice);
	DDX_Text(pDX, IDC_PRICE_BUY1, m_strPriceBuy1);
	DDX_Text(pDX, IDC_PRICE_BUY2, m_strPriceBuy2);
	DDX_Text(pDX, IDC_PRICE_BUY3, m_strPriceBuy3);
	DDX_Text(pDX, IDC_PRICE_BUY4, m_strPriceBuy4);
	DDX_Text(pDX, IDC_PRICE_BUY5, m_strPriceBuy5);
	DDX_Text(pDX, IDC_PRICE_SELL1, m_strPriceSell1);
	DDX_Text(pDX, IDC_PRICE_SELL2, m_strPriceSell2);
	DDX_Text(pDX, IDC_PRICE_SELL3, m_strPriceSell3);
	DDX_Text(pDX, IDC_PRICE_SELL4, m_strPriceSell4);
	DDX_Text(pDX, IDC_PRICE_SELL5, m_strPriceSell5);
	DDX_Text(pDX, IDC_UNSETTLED, m_strUnsettled);
	DDX_Text(pDX, IDC_VOLUME, m_strVolume);
	DDX_Text(pDX, IDC_VOLUME_BUY1, m_strVolumeBuy1);
	DDX_Text(pDX, IDC_VOLUME_BUY2, m_strVolumeBuy2);
	DDX_Text(pDX, IDC_VOLUME_BUY3, m_strVolumeBuy3);
	DDX_Text(pDX, IDC_VOLUME_BUY4, m_strVolumeBuy4);
	DDX_Text(pDX, IDC_VOLUME_BUY5, m_strVolumeBuy5);
	DDX_Text(pDX, IDC_VOLUME_SELL1, m_strVolumeSell1);
	DDX_Text(pDX, IDC_VOLUME_SELL2, m_strVolumeSell2);
	DDX_Text(pDX, IDC_VOLUME_SELL3, m_strVolumeSell3);
	DDX_Text(pDX, IDC_VOLUME_SELL4, m_strVolumeSell4);
	DDX_Text(pDX, IDC_VOLUME_SELL5, m_strVolumeSell5);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CVCOptDlg, CDialog)
	//{{AFX_MSG_MAP(CVCOptDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_REQUEST, OnRequest)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CVCOptDlg message handlers

BOOL CVCOptDlg::OnInitDialog()
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
// 추가
	HRESULT hr;
	IUnknownPtr pUnk;

	try {
		hr = m_pOptionCur.CreateInstance(CLSID_OptionCur);
		if (FAILED(hr))
			_com_raise_error(hr);
		pUnk = m_pOptionCur;
		hr = m_Handler.DispEventAdvise(pUnk);
		if (FAILED(hr))
			_com_raise_error(hr);

		m_Handler.SetIEventHandler(this);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
//
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CVCOptDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CVCOptDlg::OnPaint() 
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
HCURSOR CVCOptDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// 추가
void CVCOptDlg::OnRequest() 
{
	// TODO: Add your control notification handler code here
	IDibPtr pOptionMst;
	CString strErrorMsg = _T("통신에러 : ");
	CString strCode;
	HRESULT hr;

	try {
		UpdateData();

		// 이전에 등록된 것이 있으면 해제
		strCode = (LPCTSTR)(_bstr_t)m_pOptionCur->GetInputValue(0);
		if (!strCode.IsEmpty()) {
			m_pOptionCur->Unsubscribe();
			if (0 != m_pOptionCur->GetDibStatus()) {
				strErrorMsg += m_pOptionCur->GetDibMsg1();
				strErrorMsg += m_pOptionCur->GetDibMsg2();
				strErrorMsg.TrimRight();
				AfxMessageBox(strErrorMsg);
				return;
			}
		}

		hr = pOptionMst.CreateInstance(CLSID_OptionMst);
		if (FAILED(hr))
			_com_raise_error(hr);

		// OptionMst를 읽음
		pOptionMst->SetInputValue(0, (LPCTSTR)m_strCode);
		pOptionMst->BlockRequest();
		if (0 != pOptionMst->GetDibStatus()) {
			strErrorMsg += pOptionMst->GetDibMsg1();
			strErrorMsg += pOptionMst->GetDibMsg2();
			strErrorMsg.TrimRight();
			AfxMessageBox(strErrorMsg);
			return;
		}

		m_strPrice.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(93));		// 현재가
		m_strVolume.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(97));			// 거래량
		m_strUnsettled.Format(_T("%0.2f"), (long)pOptionMst->GetHeaderValue(99));	// 미결제약정
		m_strPriceBuy1.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(59));	// 1차 매수호가
		m_strPriceBuy2.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(63));	// 2차 매수호가
		m_strPriceBuy3.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(67));	// 3차 매수호가
		m_strPriceBuy4.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(71));	// 4차 매수호가
		m_strPriceBuy5.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(75));	// 5차 매수호가
		m_strPriceSell1.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(58));	// 1차 매도호가
		m_strPriceSell2.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(62));	// 2차 매도호가
		m_strPriceSell3.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(66));	// 3차 매도호가
		m_strPriceSell4.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(70));	// 4차 매도호가
		m_strPriceSell5.Format(_T("%0.2f"), (float)pOptionMst->GetHeaderValue(74));	// 5차 매도호가
		m_strVolumeBuy1.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(61));		// 1차 매수잔량
		m_strVolumeBuy2.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(65));		// 2차 매수잔량
		m_strVolumeBuy3.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(69));		// 3차 매수잔량
		m_strVolumeBuy4.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(73));		// 4차 매수잔량
		m_strVolumeBuy5.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(77));		// 5차 매수잔량
		m_strVolumeSell1.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(60));	// 1차 매도잔량
		m_strVolumeSell2.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(64));	// 2차 매도잔량
		m_strVolumeSell3.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(68));	// 3차 매도잔량
		m_strVolumeSell4.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(72));	// 4차 매도잔량
		m_strVolumeSell5.Format(_T("%d"), (long)pOptionMst->GetHeaderValue(76));	// 5차 매도잔량

		// OptionCur를 등록
		m_pOptionCur->SetInputValue(0, (LPCTSTR)m_strCode);
		m_pOptionCur->SubscribeLatest();
		if (0 != m_pOptionCur->GetDibStatus()) {
			strErrorMsg += m_pOptionCur->GetDibMsg1();
			strErrorMsg += m_pOptionCur->GetDibMsg2();
			strErrorMsg.TrimRight();
			AfxMessageBox(strErrorMsg);
			return;
		}

		UpdateData(FALSE);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
}

void CVCOptDlg::Received()
{
	CString strErrorMsg;

	try {
		if (0 != m_pOptionCur->GetDibStatus()) {
			strErrorMsg += m_pOptionCur->GetDibMsg1();
			strErrorMsg += m_pOptionCur->GetDibMsg2();
			strErrorMsg.TrimRight();
			AfxMessageBox(strErrorMsg);
			return;
		}

		m_strPrice.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(24));		// 현재가
		m_strVolume.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(29));			// 거래량
		m_strUnsettled.Format(_T("%0.2f"), (long)m_pOptionCur->GetHeaderValue(38));		// 미결제약정
		m_strPriceBuy1.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(13));	// 1차 매수호가
		m_strPriceBuy2.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(14));	// 2차 매수호가
		m_strPriceBuy3.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(15));	// 3차 매수호가
		m_strPriceBuy4.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(16));	// 4차 매수호가
		m_strPriceBuy5.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(17));	// 5차 매수호가
		m_strPriceSell1.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(2));	// 1차 매도호가
		m_strPriceSell2.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(3));	// 2차 매도호가
		m_strPriceSell3.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(4));	// 3차 매도호가
		m_strPriceSell4.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(5));	// 4차 매도호가
		m_strPriceSell5.Format(_T("%0.2f"), (float)m_pOptionCur->GetHeaderValue(6));	// 5차 매도호가
		m_strVolumeBuy1.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(18));		// 1차 매수잔량
		m_strVolumeBuy2.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(19));		// 2차 매수잔량
		m_strVolumeBuy3.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(20));		// 3차 매수잔량
		m_strVolumeBuy4.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(21));		// 4차 매수잔량
		m_strVolumeBuy5.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(22));		// 5차 매수잔량
		m_strVolumeSell1.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(7));		// 1차 매도잔량
		m_strVolumeSell2.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(8));		// 2차 매도잔량
		m_strVolumeSell3.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(9));		// 3차 매도잔량
		m_strVolumeSell4.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(10));		// 4차 매도잔량
		m_strVolumeSell5.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(11));		// 5차 매도잔량
		m_strCountBuy1.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(45));		// 1차 매수건수
		m_strCountBuy2.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(46));		// 2차 매수건수
		m_strCountBuy3.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(47));		// 3차 매수건수
		m_strCountBuy4.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(48));		// 4차 매수건수
		m_strCountBuy5.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(49));		// 5차 매수건수
		m_strCountSell1.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(39));		// 1차 매도건수
		m_strCountSell2.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(40));		// 2차 매도건수
		m_strCountSell3.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(41));		// 3차 매도건수
		m_strCountSell4.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(42));		// 4차 매도건수
		m_strCountSell5.Format(_T("%d"), (long)m_pOptionCur->GetHeaderValue(43));		// 5차 매도건수

		UpdateData(FALSE);
	}
	catch (_com_error e) {
		AfxMessageBox(e.ErrorMessage());
	}
}

CVCOptDlg::~CVCOptDlg()
{
	if (NULL != m_pOptionCur) {
		IUnknownPtr pUnk = m_pOptionCur;
		m_Handler.DispEventUnadvise(pUnk);
	}
}
//
