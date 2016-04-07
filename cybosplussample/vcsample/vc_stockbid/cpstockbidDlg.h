// cpstockbidDlg.h : header file
//

#if !defined(AFX_CPSTOCKBIDDLG_H__985995CA_1E2B_4E06_9DB3_7AE0A6B33B62__INCLUDED_)
#define AFX_CPSTOCKBIDDLG_H__985995CA_1E2B_4E06_9DB3_7AE0A6B33B62__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CCpstockbidDlg dialog

class CCpstockbidDlg : public CDialog
{
// Construction
public:
	CCpstockbidDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CCpstockbidDlg)
	enum { IDD = IDD_CPSTOCKBID_DIALOG };
	CListCtrl	m_hogaBox;
	CString	m_jongcode;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCpstockbidDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CCpstockbidDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnStockbid();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CPSTOCKBIDDLG_H__985995CA_1E2B_4E06_9DB3_7AE0A6B33B62__INCLUDED_)
