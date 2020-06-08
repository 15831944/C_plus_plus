// ReadExcelDlg.cpp : Implementierungsdatei
#pragma comment(lib, "odbccp32.lib")

#include "stdafx.h"
#include "ReadExcel.h"
#include "ReadExcelDlg.h"
#include "odbcinst.h"

#include <iostream>

#undef _WIN32
#if !defined(_WIN32)
#define WIN32 must be defined   // C1189
#endif


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



/////////////////////////////////////////////////////////////////////////////
// CReadExcelDlg Dialogfeld

CReadExcelDlg::CReadExcelDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CReadExcelDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CReadExcelDlg)
		// HINWEIS: Der Klassenassistent f�gt hier Member-Initialisierung ein
	//}}AFX_DATA_INIT
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CReadExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CReadExcelDlg)
	DDX_Control(pDX, IDC_LIST1, m_ctrlList);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CReadExcelDlg, CDialog)
	//{{AFX_MSG_MAP(CReadExcelDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CReadExcelDlg Nachrichten-Handler

BOOL CReadExcelDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	SetIcon(m_hIcon, TRUE);			// Gro�es Symbol verwenden
	SetIcon(m_hIcon, FALSE);		// Kleines Symbol verwenden
	
	// ZU ERLEDIGEN: Hier zus�tzliche Initialisierung einf�gen
	
	return TRUE;  // Geben Sie TRUE zur�ck, au�er ein Steuerelement soll den Fokus erhalten
}

// Wollen Sie Ihrem Dialogfeld eine Schaltfl�che "Minimieren" hinzuf�gen, ben�tigen Sie 
//  den nachstehenden Code, um das Symbol zu zeichnen. F�r MFC-Anwendungen, die das 
//  Dokument/Ansicht-Modell verwenden, wird dies automatisch f�r Sie erledigt.

void CReadExcelDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // Ger�tekontext f�r Zeichnen

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Symbol in Client-Rechteck zentrieren
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Symbol zeichnen
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

HCURSOR CReadExcelDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}



//******************************************************************
// Read that Excel Sheet
//******************************************************************
// The method OnButton1() and GetExcelDriver() demonstrate how
// an Excel file can be read. Besides that two more interesting
// features are demonstrated: 
//		1) The use of ODBC without having a complete DSN 
//       installed in the ODBC manager
//    2) The use of CRecordset without having a class 
//       derived from it
//
// But there have to be preparations:
//    You must have an Excel ODBC Driver installed (you 
//    wouldn�t have guessed..). And there has to be database support,
//    so including <afxdb.h> is really not a bad idea. Last but
//    not least, if you want to determine the full name of that
//    Excel driver automagically (like I did in GetExcelDriver() )
//    you need "odbcinst.h" to be included also.
//
// And now for the drawbacks: 
//    Feature 1) only works with ODBC Admin V3.51 and higher. 
//    Earlier versions will not be able to use a DSN that actually
//    isn�t installed. 
//    Feature 2) needs to be a readonly, foreward only recset.
//    So any attempts to change the data or to move back will 
//    fail horribly. If you need to do something like that you�re
//    bound to use CRecordset the "usual" way. Another drawback is
//    that the tremendous overhead of CRecordset does in fact make
//    it rather slow. A solution to this would be using the class
//    CSQLDirect contributed by Dave Merner at codeguru�s
//    http://www.codeguru.com/mfc_database/direct_sql_with_odbc.shtml
//
// Corresponding articles:
//    For more stuff about writing into an Excel file or using a not
//    registered DSN please refer my article
//    http://www.codeguru.com/mfc_database/excel_sheets_using_odbc.shtml		
//
// There�s still work to do:
//    One unsolved mystery in reading those files is how to get the
//    data WITHOUT having a name defined for it. That means
//    how can the structure of the data be retrieved, how many 
//    "tables" are in there, and so on. If you have any idea about 
//    that I�d be glad to read it under almikula@EUnet.at (please 
//    make a CC to alexander.mikula@siemens.at)
//
//
// After my article at CodeGuru�s concerning how to write into an Excel 
// file I got tons of requests about how to read from such a file. 
// Well in fact I do hope this - however enhancable - example sorts 
// out the basic questions.
//
//	Have fun!
//			Alexander Mikula - The Famous CyberRat	
//******************************************************************
void CReadExcelDlg::OnButton1() 
{
	CDatabase database;
	CString sSql;
	CString sItem0, sItem1, sItem2;
	CString sDriver;
	CString sDsn;
	CString sFile = "ReadExcel.xls";		// ��� �����. ����� ����� ������������
                                  // ���-�� ����� C:\\Sheets\\WhatDoIKnow.xls
	 // ������� ���������� listbox
	m_ctrlList.ResetContent();
	
    // ���� ��� �������� Excel. ��� ����������,
    // ������ ��� Microsoft ����� ����������� ������������
    // ������������� ����� ���� "Microsoft Excel Driver (*.xls)" ������
    // "Microsoft Excel Treiber (*.xls)"
	sDriver = GetExcelDriver();
	if( sDriver.IsEmpty() )
	{
    // �� ���������� ����� ���� �������!
		AfxMessageBox("No Excel ODBC driver found");
		return;
	}
	
    // ������ ������ DSN, ������� ��� �������� � ����� Excel
    // ������ ��� �� ����������� ����� ����� DSN, ������������� �
    // ������ �������������� ODBC
	sDsn.Format("ODBC;DRIVER={%s};DSN='';DBQ=%s",sDriver,sFile);

	TRY
	{
        // ��������� ���� ������, ��������� �������������� ���������
        // ������ DSN
		database.Open(NULL,false,false,sDsn);
		
        // ������������ ������
		CRecordset recset( &database );

        // ������������ ������ SQL �������
        // ��������� ����� ������ ������ � Excel �������, ���������
        // "Insert->Names"(�������->���), ����� ����� ���� �������� �������
        // ��� � �������� � �������� ���� ������. � Excel ����� ��� ��
        // ����� ����������� ����� ����� �������.
		sSql = "SELECT ��, [����� � ������], [����� ������ � ��� ���] "		
			"FROM [WorkSheet$]"	;				
				 //"ORDER BY field_1";
	
        // ��������� ������ ������ (�������� �������� recordset)
		recset.Open(CRecordset::forwardOnly,sSql,CRecordset::readOnly);

        // �������� ����������
		while( !recset.IsEOF() )
		{
			
            // ������ ������ ����������
			recset.GetFieldValue("��",sItem0);
			recset.GetFieldValue("����� � ������",sItem1);
			recset.GetFieldValue("����� ������ � ��� ���",sItem2);

			if (StrToInt(sItem0)!=0) 
			{

            // ��������� ��������� � ������
				m_ctrlList.AddString(sItem0 + "-->" + sItem1 + " --> "+sItem2 );
			}
			// Skip to the next resultline

			recset.MoveNext();
		}

        // ��������� ���� ������
		database.Close();
							 
	}
	CATCH(CDBException, e)
	{
       // �������� ���� ������ ������� ����������, ����� ������ ������...
		AfxMessageBox("Database error: "+e->m_strError);
	}
	END_CATCH;
}


// �������� ��� Excel-ODBC �������� 
CString CReadExcelDlg::GetExcelDriver()
{
	char szBuf[2001];
	WORD cbBufMax = 2000;
	WORD cbBufOut;
	char *pszBuf = szBuf;
	CString sDriver;

    // �������� ����� ������������������� ���������
    // ("odbcinst.h" ������ ���� ������� � ������ )
   if(!SQLGetInstalledDrivers(szBuf,cbBufMax,& cbBufOut))
		return "";
	
    // ���� �������...
	do
	{
		if( strstr( pszBuf, "Excel" ) != 0 )
		{
            // ����� !
			sDriver = CString( pszBuf );
			break;
		}
		pszBuf = strchr( pszBuf, '\0' ) + 1;
	}
	while( pszBuf[1] != '\0' );

	return sDriver;
}

