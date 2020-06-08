// pr5View.cpp : implementation of the CPr5View class
//

#include "stdafx.h"
#include "pr5.h"

#include "pr5Doc.h"
#include "pr5View.h"

#include "testdialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPr5View

IMPLEMENT_DYNCREATE(CPr5View, CView)

BEGIN_MESSAGE_MAP(CPr5View, CView)
	//{{AFX_MSG_MAP(CPr5View)
	ON_WM_LBUTTONDOWN()
	//}}AFX_MSG_MAP
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, CView::OnFilePrintPreview)
	ON_COMMAND(ID_FILE_OPEN, &CPr5View::OnFileOpen)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPr5View construction/destruction

CPr5View::CPr5View()
	: m_string(_T("fgbfgnffgnbgfbnfgbnfgb"))
{
	// TODO: add construction code here

}

CPr5View::~CPr5View()
{
}

BOOL CPr5View::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CPr5View drawing

void CPr5View::OnDraw(CDC* pDC)
{
	CPr5Doc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	
	pDC->TextOut(10,10,pDoc->stringData); 
}

/////////////////////////////////////////////////////////////////////////////
// CPr5View printing

BOOL CPr5View::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void CPr5View::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void CPr5View::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

/////////////////////////////////////////////////////////////////////////////
// CPr5View diagnostics

#ifdef _DEBUG
void CPr5View::AssertValid() const
{
	CView::AssertValid();
}

void CPr5View::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CPr5Doc* CPr5View::GetDocument() // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CPr5Doc)));
	return (CPr5Doc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CPr5View message handlers

void CPr5View::OnLButtonDown(UINT nFlags, CPoint point) 
{

	
}


void CPr5View::OnFileOpen()
{

	CPr5View cdialog; 

		CPr5Doc* pDoc = GetDocument();
		pDoc->stringData=cdialog.m_string;
		Invalidate();
}
