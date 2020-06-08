// ExcelTestDlg.h : header file
//

#pragma once
#include "excel_defs.h" 

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
};
