
// aaaView.cpp : реализаци€ класса CaaaView
//

#include "stdafx.h"
// SHARED_HANDLERS можно определить в обработчиках фильтров просмотра реализации проекта ATL, эскизов
// и поиска; позвол€ет совместно использовать код документа в данным проекте.
#ifndef SHARED_HANDLERS
#include "aaa.h"
#endif

#include "aaaDoc.h"
#include "aaaView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CaaaView

IMPLEMENT_DYNCREATE(CaaaView, CEditView)

BEGIN_MESSAGE_MAP(CaaaView, CEditView)
	// —тандартные команды печати
	ON_COMMAND(ID_FILE_PRINT, &CEditView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CEditView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CaaaView::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

// создание/уничтожение CaaaView

CaaaView::CaaaView()
{
	// TODO: добавьте код создани€

}

CaaaView::~CaaaView()
{
}

BOOL CaaaView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: изменить класс Window или стили посредством изменени€
	//  CREATESTRUCT cs

	BOOL bPreCreated = CEditView::PreCreateWindow(cs);
	cs.style &= ~(ES_AUTOHSCROLL|WS_HSCROLL);	// –азрешить перенос слов

	return bPreCreated;
}


// печать CaaaView


void CaaaView::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

BOOL CaaaView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// подготовка CEditView по умолчанию
	return CEditView::OnPreparePrinting(pInfo);
}

void CaaaView::OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo)
{
	// Ќачало печати CEditView по умолчанию
	CEditView::OnBeginPrinting(pDC, pInfo);
}

void CaaaView::OnEndPrinting(CDC* pDC, CPrintInfo* pInfo)
{
	// «авершение печати CEditView по умолчанию
	CEditView::OnEndPrinting(pDC, pInfo);
}

void CaaaView::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void CaaaView::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}


// диагностика CaaaView

#ifdef _DEBUG
void CaaaView::AssertValid() const
{
	CEditView::AssertValid();
}

void CaaaView::Dump(CDumpContext& dc) const
{
	CEditView::Dump(dc);
}

CaaaDoc* CaaaView::GetDocument() const // встроена неотлаженна€ верси€
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CaaaDoc)));
	return (CaaaDoc*)m_pDocument;
}
#endif //_DEBUG


// обработчики сообщений CaaaView
