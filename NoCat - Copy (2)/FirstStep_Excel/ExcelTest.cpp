// ExcelTest.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "ExcelTest.h"
#include "ExcelTestDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CExcelTestApp

BEGIN_MESSAGE_MAP(CExcelTestApp, CWinApp)
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()


// CExcelTestApp construction

CExcelTestApp::CExcelTestApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CExcelTestApp object

CExcelTestApp theApp;


// CExcelTestApp initialization

BOOL CExcelTestApp::InitInstance()
{
	CWinApp::InitInstance();

	if(!AfxOleInit()) // Your addition starts here
	{
		AfxMessageBox("Could not initialize COM dll");
		return FALSE;
	}               // End of your addition


	AfxEnableControlContainer();


	CExcelTestDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with OK
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
