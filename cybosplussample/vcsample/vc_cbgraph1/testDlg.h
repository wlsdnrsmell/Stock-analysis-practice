// testDlg.h : header file
//

#if !defined(AFX_TESTDLG_H__1795C76F_479E_4414_B104_99D0EB61F6CD__INCLUDED_)
#define AFX_TESTDLG_H__1795C76F_479E_4414_B104_99D0EB61F6CD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include <afxtempl.h>


/////////////////////////////////////////////////////////////////////////////
// CTestDlg dialog

class CTestDlg : public CDialog
{
// Construction

private:
     CImageList* m_pImageList;
	 CMap< HTREEITEM,HTREEITEM&,CString,CString& > tooltipMap;
public:
	CTestDlg(CWnd* pParent = NULL);	// standard constructor
    ~CTestDlg();
// Dialog Data
	//{{AFX_DATA(CTestDlg)
	enum { IDD = IDD_TEST_DIALOG };
	CEdit	m_year;
	CButton	m_show;
	CToolTipCtrl m_tooltip;
	CTreeCtrl	m_tree;
	CStatic	m_time;
	CStatic	m_degree;
	CProgressCtrl	m_percent;
	CButton	m_stock;
	CButton	m_option;
	CButton	m_graph;
	CTabCtrl	m_tab;
	CButton	m_future;
	CListBox	m_list;
	CString	m_leap;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTestDlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CString converttime(DWORD a);
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CTestDlg)
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual BOOL OnInitDialog();
	afx_msg void OnOption();
	afx_msg void OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnGraph();
	afx_msg void OnStock();
	afx_msg void OnLeapyear();
	afx_msg void OnChangeLeap();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TESTDLG_H__1795C76F_479E_4414_B104_99D0EB61F6CD__INCLUDED_)
