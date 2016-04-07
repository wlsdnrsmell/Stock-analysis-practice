// StockCurDlgDlg.h : header file
//

#if !defined(AFX_STOCKCURDLGDLG_H__6F2E06E9_5CC3_4042_BDDF_77E9030A675C__INCLUDED_)
#define AFX_STOCKCURDLGDLG_H__6F2E06E9_5CC3_4042_BDDF_77E9030A675C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CStockCurDlgDlg dialog
//추가->
#define MYSTOCKCUR_EVENT_ID 1
#define MYCPSVR8092S_EVENT_ID 2
extern _ATL_FUNC_INFO ReceivedInfo;
//추가-<

class CStockCurDlgDlg : public CDialog
//추가->
	,
	public IDispEventSimpleImpl<MYSTOCKCUR_EVENT_ID, CStockCurDlgDlg, &__uuidof(DSCBO1Lib::_IDibEvents)>, //stockcur 이벤트 
	public IDispEventSimpleImpl<MYCPSVR8092S_EVENT_ID, CStockCurDlgDlg, &__uuidof(DSCBO1Lib::_IDibEvents)>	 //cpsvr8092s 이벤트 
//추가-<
{
// Construction
public:
	CStockCurDlgDlg(CWnd* pParent = NULL);	// standard constructor
	//추가->
	virtual ~CStockCurDlgDlg()
	{
		IDispEventSimpleImpl<MYSTOCKCUR_EVENT_ID, CStockCurDlgDlg, &__uuidof(DSCBO1Lib::_IDibEvents)>::DispEventUnadvise(m_pStockCur);
		IDispEventSimpleImpl<MYCPSVR8092S_EVENT_ID, CStockCurDlgDlg, &__uuidof(DSCBO1Lib::_IDibEvents)>::DispEventUnadvise(m_pCpSvr8092S);
	};
	//추가-<
// Dialog Data
	//{{AFX_DATA(CStockCurDlgDlg)
	enum { IDD = IDD_STOCKCURDLG_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CStockCurDlgDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CStockCurDlgDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnTest();
	afx_msg void OnTest2();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

//추가->
public:
	BEGIN_SINK_MAP(CStockCurDlgDlg)		
		SINK_ENTRY_INFO(MYSTOCKCUR_EVENT_ID, __uuidof(DSCBO1Lib::_IDibEvents), 1, OnMyStockCurReceived, &ReceivedInfo)
		SINK_ENTRY_INFO(MYCPSVR8092S_EVENT_ID, __uuidof(DSCBO1Lib::_IDibEvents), 1, OnMyCpSvr8092SReceived, &ReceivedInfo)
	END_SINK_MAP()

	void __stdcall OnMyStockCurReceived();
	void __stdcall OnMyCpSvr8092SReceived();    

	DSCBO1Lib::IDibPtr m_pStockCur;
	DSCBO1Lib::IDibPtr m_pCpSvr8092S;    
//추가-<
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STOCKCURDLGDLG_H__6F2E06E9_5CC3_4042_BDDF_77E9030A675C__INCLUDED_)
