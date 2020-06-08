
#include "stdafx.h"
#include "ExcelTest.h"
#include "ExcelTestDlg.h"
#include "Excel_enums.h"
#include <iostream>


#include <string>


#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define OL(x) (COleVariant(long(x)))
#define OS(x) (COleVariant(_T(x)))





CExcelTestDlg::CExcelTestDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CExcelTestDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelTestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CExcelTestDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_TEST, OnBnClickedTest)
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
END_MESSAGE_MAP()


// CExcelTestDlg message handlers

BOOL CExcelTestDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CExcelTestDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CExcelTestDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CExcelTestDlg::OnBnClickedTest()
{
	//
	//CString m_Text;	// создание стандартной панели выбора файла Open
 //
 // 	CFileDialog DlgOpen(TRUE,(LPCSTR)"xls",NULL,
 //
 //		OFN_HIDEREADONLY,(LPCSTR)" Text Files (*.xls) |*.xls||");
 //
 //  	// отображение стандартной панели выбора файла Open
 //
 //	if(DlgOpen.DoModal()==IDOK)
 //
 //   	{	// создание объекта и открытие файла для чтения
	//
 //
 // 	}
 

//--------------------------чтение----------------------------------
	if(!app.CreateDispatch(_T("Excel.Application"))) //запустить сервер
	{
		AfxMessageBox(_T("Ошибка при старте Excel!"));
		return;
	}
	else

	app.put_Visible(FALSE); //и сделать его Невидимым

	CWorkbooks oBooks;
 	CWorkbook oBook;
	CWorksheets oSheets;
	CWorksheet oSheet;
	CRange oRange1;	
	CRange oRangeEmpty;
	oBooks = app.get_Workbooks();

	
	COleVariant Resalt0(long(0));




	

	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant covTrue((short)TRUE, VT_BOOL), covFalse((short)FALSE, VT_BOOL);
	COleVariant covBOOL((short)FALSE, VT_BOOL);

	oBooks.Open(_T("C:\\1.xls"),
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional );
	
	oBook= oBooks.get_Item(COleVariant(long(1)));
	oSheets = oBook.get_Worksheets();
            
	oSheet=oSheets.get_Item(COleVariant(long(1)));

	oRangeEmpty = oSheet.get_Range(COleVariant(_T("F380")),COleVariant(_T("F380")));
	COleVariant ResaltEmpty = oRangeEmpty.get_Value2();
	

	//COleVariant Resalt;

	//int n = 1;
	//char *c;
	//do
	////{
	

	//int i = 1;
	//std::string str;

	//itoa(i,str,10);
	CString w;


	oRange1 = oSheet.get_Range(COleVariant(_T("F22")),COleVariant(_T("F22")));
	COleVariant Resalt1 = oRange1.get_Value2();
	
	if(Resalt1==ResaltEmpty)
	{
		AfxMessageBox(_T("Вижу пусто"));
	}


	//} while (Resalt.lVal!=ResaltEmpty.lVal);
	


	//if (Resalt.lVal!=Resalt0.lVal)
	//{
	//	AfxMessageBox("НЕРавно");
	//}
	//else
	//{
	//	AfxMessageBox("Равно");
	//}


	oBook.Close(covFalse, covOptional, covOptional); 
	oBook.ReleaseDispatch(); 
	oBooks.Close(); 
	oBooks.ReleaseDispatch();

	//--------------------------чтение----------------------------------
	


	//--------------------------запись----------------------------------

	app.put_Visible(TRUE); //и сделать его видимым

	CWorkbooks Books;
 	CWorkbook Book;
	CWorksheets Sheets;
	CWorksheet Sheet;
	CRange Range2;	
	Books = app.get_Workbooks();

	Books.Open(_T("C:\\2.xls"),
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional );
	
	Book= Books.get_Item(COleVariant(long(1)));
	Sheets = Book.get_Worksheets();
            
	Sheet=Sheets.get_Item(COleVariant(long(3)));
	
	Range2 = Sheet.get_Range(COleVariant(_T("D12")),COleVariant(_T("D15")));
	Range2.put_Value2(COleVariant(_T(ResaltEmpty)));
//--------------------------запись----------------------------------

}


void CExcelTestDlg::OnBnClickedCancel()
{
		
	// --- Выход из программы без закрытия Excel, просто отсоединяемся от него
	app.DetachDispatch();

	OnCancel();
}
