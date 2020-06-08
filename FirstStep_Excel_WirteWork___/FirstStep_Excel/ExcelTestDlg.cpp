
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
	, LInt(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelTestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);



	DDX_Control(pDX, IDC_GIFFIRST, m_Picture);

	DDX_Control(pDX, IDB_MYNAME, My_Name);
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

	if (m_Picture.Load(MAKEINTRESOURCE(IDR_GIF2),_T("GIF")))
		m_Picture.Draw();
	if (My_Name.Load(MAKEINTRESOURCE(IDR_GIF3),_T("GIF")))
		My_Name.Draw();
	
	
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
	CRange oRange2;
	CRange oRange3;
	CRange oRange4;
	CRange oRange5;
	CRange oRange6;
	CRange oRange7;
	CRange oRange8;

	CRange oRangeEmpty;
	oBooks = app.get_Workbooks();

	
	COleVariant Resalt0(long(0));



	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant covTrue((short)TRUE, VT_BOOL), covFalse((short)FALSE, VT_BOOL);
	COleVariant covBOOL((short)FALSE, VT_BOOL);

	oBooks.Open(_T("\\1.xls"),
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional );
	
	oBook= oBooks.get_Item(COleVariant(long(1)));
	oSheets = oBook.get_Worksheets();
            
	oSheet=oSheets.get_Item(COleVariant(long(1)));

	oRange1 = oSheet.get_Range(COleVariant(_T("D2")),COleVariant(_T("D378")));
	COleSafeArray SchetaFromExcel(oRange1.get_Value(covOptional));
	oRange2 = oSheet.get_Range(COleVariant(_T("CA378")),COleVariant(_T("CA378")));
	oRangeEmpty = oSheet.get_Range(COleVariant(_T("CA378")),COleVariant(_T("CA378")));
	COleVariant ResaltEmpty = oRangeEmpty.get_Value2();

	oRange2 = oSheet.get_Range(COleVariant(_T("CA2")),COleVariant(_T("CA378")));
	COleSafeArray SchetaFromExcelNew(oRange2.get_Value(covOptional));

	oRange3 = oSheet.get_Range(COleVariant(_T("F2")),COleVariant(_T("F378")));
	COleSafeArray AktivRubFromExcel(oRange3.get_Value(covOptional));
	oRange4 = oSheet.get_Range(COleVariant(_T("I2")),COleVariant(_T("I378")));
	COleSafeArray PasivRubFromExcel(oRange4.get_Value(covOptional));
	oRange5 = oSheet.get_Range(COleVariant(_T("CA2")),COleVariant(_T("CA378")));
	COleSafeArray RubFromExcelNew(oRange5.get_Value(covOptional));

	oRange6 = oSheet.get_Range(COleVariant(_T("G2")),COleVariant(_T("G378")));
	COleSafeArray AktivValutFromExcel(oRange6.get_Value(covOptional));
	oRange7 = oSheet.get_Range(COleVariant(_T("J2")),COleVariant(_T("J378")));
	COleSafeArray PasivValutFromExcel(oRange7.get_Value(covOptional));
	oRange8 = oSheet.get_Range(COleVariant(_T("CA2")),COleVariant(_T("CA378")));
	COleSafeArray ValutFromExcelNew(oRange8.get_Value(covOptional));



	COleVariant vData;
	COleVariant vData0;
	COleVariant vData1;
	long iRows;
	long iRows2;
    SchetaFromExcel.GetUBound(1, &iRows);
	long index[2];
	long index1[2];
	int i = 1;
	index1[1]=i;
	index[1]=1;

	for (int rowCounter = 1; rowCounter <= iRows; rowCounter++) {

			index[0]=rowCounter;
			SchetaFromExcel.GetElement(index,vData);
			CString szdata(vData);
			if (szdata=="0") {
			}
			else {
				index1[0]=i;
				SchetaFromExcelNew.PutElement(index1,vData);

				AktivRubFromExcel.GetElement(index,vData0);
				CString szdata0(vData0);
				if (szdata0=="0") {
					PasivRubFromExcel.GetElement(index,vData0);
					RubFromExcelNew.PutElement(index1,vData0);
				}
				else {
					RubFromExcelNew.PutElement(index1,vData0);
				}
				AktivValutFromExcel.GetElement(index,vData1);
				CString szdata1(vData1);
				if (szdata1=="0") {
					PasivValutFromExcel.GetElement(index,vData1);
					ValutFromExcelNew.PutElement(index1,vData1);
				}
				else {
					ValutFromExcelNew.PutElement(index1,vData1);
				}

				++i;
			}
	}
	



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
	CRange Range3;
	CRange Range2;
	CRange Range1;
	Books = app.get_Workbooks();

	Books.Open(_T("\\2.xls"),
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional );
	
	Book= Books.get_Item(COleVariant(long(1)));
	Sheets = Book.get_Worksheets();
            
	Sheet=Sheets.get_Item(COleVariant(long(3)));
	
	Range1 = Sheet.get_Range(COleVariant(_T("D6")),COleVariant(_T("D1390")));
	Range1.put_Value2(ResaltEmpty);
	COleSafeArray RubFromExcelToo(Range1.get_Value(covOptional));

	Range3 = Sheet.get_Range(COleVariant(_T("E6")),COleVariant(_T("E1390")));
	Range3.put_Value2(ResaltEmpty);
	COleSafeArray ValutFromExcelToo(Range3.get_Value(covOptional));

	Range2 = Sheet.get_Range(COleVariant(_T("B6")),COleVariant(_T("B1390")));
	COleSafeArray SchetaFromExcelToo(Range2.get_Value(covOptional));

    SchetaFromExcelNew.GetUBound(1, &iRows);
	SchetaFromExcelToo.GetUBound(1, &iRows2);
	index[1]=1;
	index1[1]=1;
	COleVariant vData2;
	COleVariant vData3;
	for (int rowCounter = 1; rowCounter <= 171; rowCounter++) {
			index[0]=rowCounter;
			SchetaFromExcelNew.GetElement(index,vData);
			int yes = 0;
			for (int rowCounter2 = 1; rowCounter2 <= iRows2; rowCounter2++) {
				index1[0]=rowCounter2;
				SchetaFromExcelToo.GetElement(index1,vData2);
				if (vData==vData2) {
				yes = 1;
				RubFromExcelNew.GetElement(index,vData3);
				RubFromExcelToo.PutElement(index1,vData3);
				ValutFromExcelNew.GetElement(index,vData3);
				ValutFromExcelToo.PutElement(index1,vData3);
				}
				if ((yes==0)&(rowCounter2==iRows2)) {
					CString w = CString(vData);
					AfxMessageBox(_T(w));
			}

	}
	Range1.put_Value2(RubFromExcelToo);
	Range3.put_Value2(ValutFromExcelToo);
//--------------------------запись----------------------------------

}


void CExcelTestDlg::OnBnClickedCancel()
{
		
	// --- Выход из программы без закрытия Excel, просто отсоединяемся от него
	app.DetachDispatch();

	OnCancel();
}



