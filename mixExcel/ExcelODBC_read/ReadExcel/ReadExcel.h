// ReadExcel.h : Haupt-Header-Datei f�r die Anwendung READEXCEL
//

#if !defined(AFX_READEXCEL_H__660FFF81_053E_11D3_A579_00105A59FE2F__INCLUDED_)
#define AFX_READEXCEL_H__660FFF81_053E_11D3_A579_00105A59FE2F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// Hauptsymbole

/////////////////////////////////////////////////////////////////////////////
// CReadExcelApp:
// Siehe ReadExcel.cpp f�r die Implementierung dieser Klasse
//

class CReadExcelApp : public CWinApp
{
public:
	CReadExcelApp();

// �berladungen
	// Vom Klassenassistenten generierte �berladungen virtueller Funktionen
	//{{AFX_VIRTUAL(CReadExcelApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementierung

	//{{AFX_MSG(CReadExcelApp)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ f�gt unmittelbar vor der vorhergehenden Zeile zus�tzliche Deklarationen ein.

#endif // !defined(AFX_READEXCEL_H__660FFF81_053E_11D3_A579_00105A59FE2F__INCLUDED_)
