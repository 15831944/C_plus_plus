// DragString.cpp: ���� ����������
//

#include "stdafx.h"
#include "ExcelTest.h"
#include "DragString.h"


// DragString

IMPLEMENT_DYNCREATE(DragString, CFormView)

DragString::DragString()
	: CFormView(DragString::IDD)
{

}

DragString::~DragString()
{
}

void DragString::DoDataExchange(CDataExchange* pDX)
{
	CFormView::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(DragString, CFormView)
END_MESSAGE_MAP()


// ����������� DragString

#ifdef _DEBUG
void DragString::AssertValid() const
{
	CFormView::AssertValid();
}

#ifndef _WIN32_WCE
void DragString::Dump(CDumpContext& dc) const
{
	CFormView::Dump(dc);
}
#endif
#endif //_DEBUG


// ����������� ��������� DragString
