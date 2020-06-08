// ReadExcelDlg.h : Header-Datei
//

#if !defined(AFX_READEXCELDLG_H__660FFF83_053E_11D3_A579_00105A59FE2F__INCLUDED_)
#define AFX_READEXCELDLG_H__660FFF83_053E_11D3_A579_00105A59FE2F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CReadExcelDlg Dialogfeld

class CReadExcelDlg : public CDialog
{
// Konstruktion
public:
	CReadExcelDlg(CWnd* pParent = NULL);	// Standard-Konstruktor

	CString GetExcelDriver( );

// Dialogfelddaten
	//{{AFX_DATA(CReadExcelDlg)
	enum { IDD = IDD_READEXCEL_DIALOG };
	CListBox	m_ctrlList;
	//}}AFX_DATA

	// Vom Klassenassistenten generierte Überladungen virtueller Funktionen
	//{{AFX_VIRTUAL(CReadExcelDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV-Unterstützung
	//}}AFX_VIRTUAL

// Implementierung
protected:
	HICON m_hIcon;

	// Generierte Message-Map-Funktionen
	//{{AFX_MSG(CReadExcelDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButton1();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ fügt unmittelbar vor der vorhergehenden Zeile zusätzliche Deklarationen ein.

#endif // !defined(AFX_READEXCELDLG_H__660FFF83_053E_11D3_A579_00105A59FE2F__INCLUDED_)
