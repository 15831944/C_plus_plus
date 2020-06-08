// pr5Doc.cpp : implementation of the CPr5Doc class
//

#include "stdafx.h"
#include "pr5.h"

#include "pr5Doc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPr5Doc

IMPLEMENT_DYNCREATE(CPr5Doc, CDocument)

BEGIN_MESSAGE_MAP(CPr5Doc, CDocument)
	//{{AFX_MSG_MAP(CPr5Doc)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPr5Doc construction/destruction

CPr5Doc::CPr5Doc()
{
	// TODO: add one-time construction code here

}

CPr5Doc::~CPr5Doc()
{
}

BOOL CPr5Doc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	stringData="";

	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CPr5Doc serialization

void CPr5Doc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}

/////////////////////////////////////////////////////////////////////////////
// CPr5Doc diagnostics

#ifdef _DEBUG
void CPr5Doc::AssertValid() const
{
	CDocument::AssertValid();
}

void CPr5Doc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CPr5Doc commands
