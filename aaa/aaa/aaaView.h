
// aaaView.h : интерфейс класса CaaaView
//

#pragma once


class CaaaView : public CEditView
{
protected: // создать только из сериализации
	CaaaView();
	DECLARE_DYNCREATE(CaaaView)

// Атрибуты
public:
	CaaaDoc* GetDocument() const;

// Операции
public:

// Переопределение
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// Реализация
public:
	virtual ~CaaaView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Созданные функции схемы сообщений
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // отладочная версия в aaaView.cpp
inline CaaaDoc* CaaaView::GetDocument() const
   { return reinterpret_cast<CaaaDoc*>(m_pDocument); }
#endif

