
// bbbView.cpp : реализаци€ класса CbbbView
//

#include "stdafx.h"
// SHARED_HANDLERS можно определить в обработчиках фильтров просмотра реализации проекта ATL, эскизов
// и поиска; позвол€ет совместно использовать код документа в данным проекте.
#ifndef SHARED_HANDLERS
#include "bbb.h"
#endif

#include "bbbDoc.h"
#include "bbbView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CbbbView

IMPLEMENT_DYNCREATE(CbbbView, CView)

BEGIN_MESSAGE_MAP(CbbbView, CView)
	// —тандартные команды печати
	ON_COMMAND(ID_FILE_PRINT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CbbbView::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

// создание/уничтожение CbbbView

CbbbView::CbbbView()
{
	// TODO: добавьте код создани€

}

CbbbView::~CbbbView()
{
}

BOOL CbbbView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: изменить класс Window или стили посредством изменени€
	//  CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

// рисование CbbbView

void CbbbView::OnDraw(CDC* pDC)
{
	CbbbDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;
	pDC->DrawText(pDoc->m_Doc,CRect(10,10,300,100),DT_CENTER); 
	// TODO: добавьте здесь код отрисовки дл€ собственных данных
}


// печать CbbbView


void CbbbView::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

BOOL CbbbView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// подготовка по умолчанию
	return DoPreparePrinting(pInfo);
}

void CbbbView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: добавьте дополнительную инициализацию перед печатью
}

void CbbbView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: добавьте очистку после печати
}

void CbbbView::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void CbbbView::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}


// диагностика CbbbView

#ifdef _DEBUG
void CbbbView::AssertValid() const
{
	CView::AssertValid();
}

void CbbbView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CbbbDoc* CbbbView::GetDocument() const // встроена неотлаженна€ верси€
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CbbbDoc)));
	return (CbbbDoc*)m_pDocument;
}
#endif //_DEBUG


// обработчики сообщений CbbbView
