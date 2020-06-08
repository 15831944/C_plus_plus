// ReadExcel.cpp : Legt das Klassenverhalten für die Anwendung fest.
//

#include "stdafx.h"
#include "ReadExcel.h"
#include "ReadExcelDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CReadExcelApp

BEGIN_MESSAGE_MAP(CReadExcelApp, CWinApp)
	//{{AFX_MSG_MAP(CReadExcelApp)
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CReadExcelApp Konstruktion

CReadExcelApp::CReadExcelApp()
{
}

/////////////////////////////////////////////////////////////////////////////
// Das einzige CReadExcelApp-Objekt

CReadExcelApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CReadExcelApp Initialisierung

BOOL CReadExcelApp::InitInstance()
{
	// Standardinitialisierung

#ifdef _AFXDLL
	Enable3dControls();			// Diese Funktion bei Verwendung von MFC in gemeinsam genutzten DLLs aufrufen
#else
	Enable3dControlsStatic();	// Diese Funktion bei statischen MFC-Anbindungen aufrufen
#endif

	CReadExcelDlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
	}
	else if (nResponse == IDCANCEL)
	{
	}

	// Da das Dialogfeld geschlossen wurde, FALSE zurückliefern, so dass wir die
	//  Anwendung verlassen, anstatt das Nachrichtensystem der Anwendung zu starten.
	return FALSE;
}
