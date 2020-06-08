
// aaaView.cpp : ���������� ������ CaaaView
//

#include "stdafx.h"
// SHARED_HANDLERS ����� ���������� � ������������ �������� ��������� ���������� ������� ATL, �������
// � ������; ��������� ��������� ������������ ��� ��������� � ������ �������.
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
	// ����������� ������� ������
	ON_COMMAND(ID_FILE_PRINT, &CEditView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CEditView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CaaaView::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

// ��������/����������� CaaaView

CaaaView::CaaaView()
{
	// TODO: �������� ��� ��������

}

CaaaView::~CaaaView()
{
}

BOOL CaaaView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: �������� ����� Window ��� ����� ����������� ���������
	//  CREATESTRUCT cs

	BOOL bPreCreated = CEditView::PreCreateWindow(cs);
	cs.style &= ~(ES_AUTOHSCROLL|WS_HSCROLL);	// ��������� ������� ����

	return bPreCreated;
}


// ������ CaaaView


void CaaaView::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

BOOL CaaaView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// ���������� CEditView �� ���������
	return CEditView::OnPreparePrinting(pInfo);
}

void CaaaView::OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo)
{
	// ������ ������ CEditView �� ���������
	CEditView::OnBeginPrinting(pDC, pInfo);
}

void CaaaView::OnEndPrinting(CDC* pDC, CPrintInfo* pInfo)
{
	// ���������� ������ CEditView �� ���������
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


// ����������� CaaaView

#ifdef _DEBUG
void CaaaView::AssertValid() const
{
	CEditView::AssertValid();
}

void CaaaView::Dump(CDumpContext& dc) const
{
	CEditView::Dump(dc);
}

CaaaDoc* CaaaView::GetDocument() const // �������� ������������ ������
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CaaaDoc)));
	return (CaaaDoc*)m_pDocument;
}
#endif //_DEBUG


// ����������� ��������� CaaaView
