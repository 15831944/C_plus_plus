
// bbbView.cpp : ���������� ������ CbbbView
//

#include "stdafx.h"
// SHARED_HANDLERS ����� ���������� � ������������ �������� ��������� ���������� ������� ATL, �������
// � ������; ��������� ��������� ������������ ��� ��������� � ������ �������.
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
	// ����������� ������� ������
	ON_COMMAND(ID_FILE_PRINT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CbbbView::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

// ��������/����������� CbbbView

CbbbView::CbbbView()
{
	// TODO: �������� ��� ��������

}

CbbbView::~CbbbView()
{
}

BOOL CbbbView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: �������� ����� Window ��� ����� ����������� ���������
	//  CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

// ��������� CbbbView

void CbbbView::OnDraw(CDC* pDC)
{
	CbbbDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;
	pDC->DrawText(pDoc->m_Doc,CRect(10,10,300,100),DT_CENTER); 
	// TODO: �������� ����� ��� ��������� ��� ����������� ������
}


// ������ CbbbView


void CbbbView::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

BOOL CbbbView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// ���������� �� ���������
	return DoPreparePrinting(pInfo);
}

void CbbbView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: �������� �������������� ������������� ����� �������
}

void CbbbView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: �������� ������� ����� ������
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


// ����������� CbbbView

#ifdef _DEBUG
void CbbbView::AssertValid() const
{
	CView::AssertValid();
}

void CbbbView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CbbbDoc* CbbbView::GetDocument() const // �������� ������������ ������
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CbbbDoc)));
	return (CbbbDoc*)m_pDocument;
}
#endif //_DEBUG


// ����������� ��������� CbbbView
