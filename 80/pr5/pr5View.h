// pr5View.h : interface of the CPr5View class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_PR5VIEW_H__0B5293C0_C950_11D3_A3DB_000001260696__INCLUDED_)
#define AFX_PR5VIEW_H__0B5293C0_C950_11D3_A3DB_000001260696__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class CPr5View : public CView
{
protected: // create from serialization only
	CPr5View();
	DECLARE_DYNCREATE(CPr5View)

// Attributes
public:
	CPr5Doc* GetDocument();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPr5View)
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CPr5View();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CPr5View)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnFileOpen();
	CString m_string;
};

#ifndef _DEBUG  // debug version in pr5View.cpp
inline CPr5Doc* CPr5View::GetDocument()
   { return (CPr5Doc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PR5VIEW_H__0B5293C0_C950_11D3_A3DB_000001260696__INCLUDED_)
