// ExcelTestDlg.h : header file
//

#pragma once
#include "excel_defs.h" 
#include "afxwin.h"
#include "PictureEx.h"
#include "afxvslistbox.h"
#include "HScrollListBox.h"

// CExcelTestDlg dialog
class CExcelTestDlg : public CDialog
{
// Construction
public:

	CExcelTestDlg(CWnd* pParent = NULL);	// standard constructor

	CApplication app;  // app - это объект _Application, т.е. Excel

// Dialog Data
	enum { IDD = IDD_EXCELTEST_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedTest();
	afx_msg void OnBnClickedCancel();
	CListBox List1;
	int LInt;
	CListBox ListInt3;
	CPictureEx m_Picture;
	CPictureEx My_Name;
	CHScrollListBox m_list;
	static const int IntObjectMassiv = 250;
	CString ObjectMassiv[IntObjectMassiv];
	afx_msg void OnLbnSelchangeList1();
	afx_msg void OnLbnDblclkList1();
//	virtual HRESULT accDoDefaultAction(VARIANT varChild);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
	afx_msg void OnNcRButtonDblClk(UINT nHitTest, CPoint point);
	afx_msg void OnMenuSelect(UINT nItemID, UINT nFlags, HMENU hSysMenu);
	afx_msg void OnBnClickedCheck1();
	afx_msg void OnBnClickedTest2();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnElementCtrl();
	afx_msg void OnBnClickedTest3();
};
