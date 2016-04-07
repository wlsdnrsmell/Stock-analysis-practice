// vc_StockMemberDlg.h : header file
//

#if !defined(AFX_VC_STOCKMEMBERDLG_H__0D8A3712_F565_45A5_9A50_F8B67289AD82__INCLUDED_)
#define AFX_VC_STOCKMEMBERDLG_H__0D8A3712_F565_45A5_9A50_F8B67289AD82__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CVc_StockMemberDlg dialog

class CCpDibEvent;

class CVc_StockMemberDlg : public CDialog
{
// Construction
public:
	CVc_StockMemberDlg(CWnd* pParent = NULL);	// standard constructor
	~CVc_StockMemberDlg();	// <= 추가한 것임

// Dialog Data
	//{{AFX_DATA(CVc_StockMemberDlg)
	enum { IDD = IDD_VC_STOCKMEMBER_DIALOG };
	CString	m_strOut;
	CString	m_strJongMok;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CVc_StockMemberDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CMapStringToString m_mapStr;
	HICON m_hIcon;

	// <= 추가한 것임
	IDibPtr m_CpDibObj;
	CCpDibEvent* m_pEvent;
	//

	// Generated message map functions
	//{{AFX_MSG(CVc_StockMemberDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnBtnRequest();
	//}}AFX_MSG
    afx_msg LONG OnReceived(WPARAM wParam, LPARAM lParam); // 추가한 것임
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_VC_STOCKMEMBERDLG_H__0D8A3712_F565_45A5_9A50_F8B67289AD82__INCLUDED_)
