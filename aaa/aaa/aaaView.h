
// aaaView.h : ��������� ������ CaaaView
//

#pragma once


class CaaaView : public CEditView
{
protected: // ������� ������ �� ������������
	CaaaView();
	DECLARE_DYNCREATE(CaaaView)

// ��������
public:
	CaaaDoc* GetDocument() const;

// ��������
public:

// ���������������
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// ����������
public:
	virtual ~CaaaView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// ��������� ������� ����� ���������
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // ���������� ������ � aaaView.cpp
inline CaaaDoc* CaaaView::GetDocument() const
   { return reinterpret_cast<CaaaDoc*>(m_pDocument); }
#endif

