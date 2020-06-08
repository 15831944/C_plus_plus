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
		// HINWEIS: Der Klassenassistent fьgt hier Member-Initialisierung ein
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

	SetIcon(m_hIcon, TRUE);			// GroЯes Symbol verwenden
	SetIcon(m_hIcon, FALSE);		// Kleines Symbol verwenden
	
	// ZU ERLEDIGEN: Hier zusдtzliche Initialisierung einfьgen
	
	return TRUE;  // Geben Sie TRUE zurьck, auЯer ein Steuerelement soll den Fokus erhalten
}

// Wollen Sie Ihrem Dialogfeld eine Schaltflдche "Minimieren" hinzufьgen, benцtigen Sie 
//  den nachstehenden Code, um das Symbol zu zeichnen. Fьr MFC-Anwendungen, die das 
//  Dokument/Ansicht-Modell verwenden, wird dies automatisch fьr Sie erledigt.

void CReadExcelDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // Gerдtekontext fьr Zeichnen

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
//    wouldnґt have guessed..). And there has to be database support,
//    so including <afxdb.h> is really not a bad idea. Last but
//    not least, if you want to determine the full name of that
//    Excel driver automagically (like I did in GetExcelDriver() )
//    you need "odbcinst.h" to be included also.
//
// And now for the drawbacks: 
//    Feature 1) only works with ODBC Admin V3.51 and higher. 
//    Earlier versions will not be able to use a DSN that actually
//    isnґt installed. 
//    Feature 2) needs to be a readonly, foreward only recset.
//    So any attempts to change the data or to move back will 
//    fail horribly. If you need to do something like that youґre
//    bound to use CRecordset the "usual" way. Another drawback is
//    that the tremendous overhead of CRecordset does in fact make
//    it rather slow. A solution to this would be using the class
//    CSQLDirect contributed by Dave Merner at codeguruґs
//    http://www.codeguru.com/mfc_database/direct_sql_with_odbc.shtml
//
// Corresponding articles:
//    For more stuff about writing into an Excel file or using a not
//    registered DSN please refer my article
//    http://www.codeguru.com/mfc_database/excel_sheets_using_odbc.shtml		
//
// Thereґs still work to do:
//    One unsolved mystery in reading those files is how to get the
//    data WITHOUT having a name defined for it. That means
//    how can the structure of the data be retrieved, how many 
//    "tables" are in there, and so on. If you have any idea about 
//    that Iґd be glad to read it under almikula@EUnet.at (please 
//    make a CC to alexander.mikula@siemens.at)
//
//
// After my article at CodeGuruґs concerning how to write into an Excel 
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
	CString sFile = "ReadExcel.xls";		// имя файла. также можно использовать
                                  // что-то вроде C:\\Sheets\\WhatDoIKnow.xls
	 // Очищаем содержимое listbox
	m_ctrlList.ResetContent();
	
    // Ищем имя драйвера Excel. Это необходимо,
    // потому что Microsoft имеет особенность использовать
    // специфицеские имена типа "Microsoft Excel Driver (*.xls)" вместо
    // "Microsoft Excel Treiber (*.xls)"
	sDriver = GetExcelDriver();
	if( sDriver.IsEmpty() )
	{
    // Не получается найти этот драйвер!
		AfxMessageBox("No Excel ODBC driver found");
		return;
	}
	
    // Создаём псевдо DSN, включая имя драйвера и файла Excel
    // теперь нам не понадобится иметь явный DSN, установленный в
    // нашего Администратора ODBC
	sDsn.Format("ODBC;DRIVER={%s};DSN='';DBQ=%s",sDriver,sFile);

	TRY
	{
        // Открываем базу данных, используя предварительно созданный
        // псевдо DSN
		database.Open(NULL,false,false,sDsn);
		
        // Распределяем записи
		CRecordset recset( &database );

        // Конструируем строку SQL запроса
        // Заполните имена секций данных в Excel таблице, используя
        // "Insert->Names"(Вставка->Имя), чтобы можно было работать данными
        // как с таблицей в реальной базе данных. В Excel файле так же
        // может содержаться более одной таблицы.
		sSql = "SELECT БС, [Актив в рублях], [Актив валюта в руб экв] "		
			"FROM [WorkSheet$]"	;				
				 //"ORDER BY field_1";
	
        // Выполняем данный запрос (косвенно открывая recordset)
		recset.Open(CRecordset::forwardOnly,sSql,CRecordset::readOnly);

        // получаем результаты
		while( !recset.IsEOF() )
		{
			
            // читаем строку результата
			recset.GetFieldValue("БС",sItem0);
			recset.GetFieldValue("Актив в рублях",sItem1);
			recset.GetFieldValue("Актив валюта в руб экв",sItem2);

			if (StrToInt(sItem0)!=0) 
			{

            // Вставляем результат в список
				m_ctrlList.AddString(sItem0 + "-->" + sItem1 + " --> "+sItem2 );
			}
			// Skip to the next resultline

			recset.MoveNext();
		}

        // Закрываем базу данных
		database.Close();
							 
	}
	CATCH(CDBException, e)
	{
       // Открытие базы данных вызвало исключение, проще говоря ошибку...
		AfxMessageBox("Database error: "+e->m_strError);
	}
	END_CATCH;
}


// Получаем имя Excel-ODBC драйвера 
CString CReadExcelDlg::GetExcelDriver()
{
	char szBuf[2001];
	WORD cbBufMax = 2000;
	WORD cbBufOut;
	char *pszBuf = szBuf;
	CString sDriver;

    // Получаем имена проинсталлированных драйверов
    // ("odbcinst.h" должен быть включён в проект )
   if(!SQLGetInstalledDrivers(szBuf,cbBufMax,& cbBufOut))
		return "";
	
    // Ищем драйвер...
	do
	{
		if( strstr( pszBuf, "Excel" ) != 0 )
		{
            // Нашли !
			sDriver = CString( pszBuf );
			break;
		}
		pszBuf = strchr( pszBuf, '\0' ) + 1;
	}
	while( pszBuf[1] != '\0' );

	return sDriver;
}

