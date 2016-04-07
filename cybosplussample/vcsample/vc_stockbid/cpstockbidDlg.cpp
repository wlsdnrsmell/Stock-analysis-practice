// cpstockbidDlg.cpp : implementation file
//

#include "stdafx.h"
#include "cpstockbid.h"
#include "cpstockbidDlg.h"

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
// CCpstockbidDlg dialog

CCpstockbidDlg::CCpstockbidDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CCpstockbidDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CCpstockbidDlg)
	m_jongcode = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCpstockbidDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CCpstockbidDlg)
	DDX_Control(pDX, IDC_LIST1, m_hogaBox);
	DDX_Text(pDX, IDC_JONGMOK, m_jongcode);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CCpstockbidDlg, CDialog)
	//{{AFX_MSG_MAP(CCpstockbidDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_STOCKBID, OnStockbid)
	ON_WM_CTLCOLOR()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCpstockbidDlg message handlers

BOOL CCpstockbidDlg::OnInitDialog()
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
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CCpstockbidDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CCpstockbidDlg::OnPaint() 
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
HCURSOR CCpstockbidDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CCpstockbidDlg::OnStockbid() 
{
	     m_hogaBox.DeleteAllItems();
	     UpdateData(TRUE);
		 IDibPtr objip;
         objip.CreateInstance(CLSID_StockBid);
		 CRect rc;
		 int ta;
		 m_hogaBox.GetClientRect(&rc);
         ta = rc.right - rc.left;
		 m_hogaBox.InsertColumn(0,"시각",LVCFMT_LEFT,ta/3);
		 m_hogaBox.InsertColumn(1,"현재가",LVCFMT_LEFT,ta/3);
		 m_hogaBox.InsertColumn(2,"체결량",LVCFMT_LEFT,ta/3);
         
		 short sn = 75;
		 objip->SetInputValue(0, (LPCSTR)m_jongcode);
		 objip->SetInputValue(2, sn);
		 do
		 {
			 objip->BlockRequest();
			 long count = (long)(_variant_t)objip->GetHeaderValue(2);
			 long curval, sellval, buyval, vol,segan;
			 CString str1,str2,str3;
			 int icount = (int) count;
			 int index = m_hogaBox.GetItemCount();
			 for( int i =  0; i < icount - 1 ; i++)
			 {
				str1.Empty();
				str2.Empty();
				str3.Empty();
				segan = (long)(_variant_t)objip->GetDataValue(0, i); //시간
				curval = (long)(_variant_t)objip->GetDataValue(4, i); //현재가
				sellval =(long)(_variant_t) objip->GetDataValue(2, i);
				buyval = (long)(_variant_t)objip->GetDataValue(3, i);
				vol = (long)(_variant_t)objip->GetDataValue(6, i);  //체결량
				str1.Format("%d", segan);	
				m_hogaBox.InsertItem(index, str1);
				str2.Format("%d", curval);
				m_hogaBox.SetItemText(index , 1, str2);
				str3.Format("%d", vol);
				m_hogaBox.SetItemText(index , 2, str3); 
				index = index + 1;
			 }			 
		} while (objip->Continue);
	 AfxMessageBox("모든 데이터 수신이 끝났습니다");
       objip.Release();
}

HBRUSH CCpstockbidDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	CBrush m_brush;
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	if (pWnd->GetDlgCtrlID() == IDC_JONGMOK)
    {
      // Set the text color to red
      pDC->SetTextColor(RGB(0, 0, 255));
	  pDC->SetBkColor(RGB(255,255,255));

      // Set the background mode for text to transparent 
      // so background will show thru.
      pDC->SetBkMode(OPAQUE);

      // Return handle to our CBrush object
      hbr = m_brush;
   }

    	
	// TODO: Return a different brush if the default is not desired
	return hbr;
}
