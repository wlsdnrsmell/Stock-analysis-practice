// VCOptDlg.h : header file
//

#if !defined(AFX_VCOPTDLG_H__465FA300_46AD_4D2B_9A6D_F5416B2312A6__INCLUDED_)
#define AFX_VCOPTDLG_H__465FA300_46AD_4D2B_9A6D_F5416B2312A6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// 추가
#include "EventHandler.h"
//

/////////////////////////////////////////////////////////////////////////////
// CVCOptDlg dialog

// public IEventHanlder 추가
class CVCOptDlg : public CDialog, public IEventHandler
{
// Construction
public:
	CVCOptDlg(CWnd* pParent = NULL);	// standard constructor
	virtual ~CVCOptDlg();

// Dialog Data
	//{{AFX_DATA(CVCOptDlg)
	enum { IDD = IDD_VCOPT_DIALOG };
	CString	m_strCode;
	CString	m_strCountBuy1;
	CString	m_strCountBuy2;
	CString	m_strCountBuy3;
	CString	m_strCountBuy4;
	CString	m_strCountBuy5;
	CString	m_strCountSell1;
	CString	m_strCountSell2;
	CString	m_strCountSell3;
	CString	m_strCountSell4;
	CString	m_strCountSell5;
	CString	m_strPrice;
	CString	m_strPriceBuy1;
	CString	m_strPriceBuy2;
	CString	m_strPriceBuy3;
	CString	m_strPriceBuy4;
	CString	m_strPriceBuy5;
	CString	m_strPriceSell1;
	CString	m_strPriceSell2;
	CString	m_strPriceSell3;
	CString	m_strPriceSell4;
	CString	m_strPriceSell5;
	CString	m_strUnsettled;
	CString	m_strVolume;
	CString	m_strVolumeBuy1;
	CString	m_strVolumeBuy2;
	CString	m_strVolumeBuy3;
	CString	m_strVolumeBuy4;
	CString	m_strVolumeBuy5;
	CString	m_strVolumeSell1;
	CString	m_strVolumeSell2;
	CString	m_strVolumeSell3;
	CString	m_strVolumeSell4;
	CString	m_strVolumeSell5;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CVCOptDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CVCOptDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnRequest();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// 추가
protected:
	IDibPtr m_pOptionCur;
	CEventHandler m_Handler;
public:
	virtual void Received();
//
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_VCOPTDLG_H__465FA300_46AD_4D2B_9A6D_F5416B2312A6__INCLUDED_)
