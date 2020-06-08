// ExcelTestDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ExcelTest.h"
#include "ExcelTestDlg.h"
#include "Excel_enums.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define OL(x) (COleVariant(long(x)))
#define OS(x) (COleVariant(_T(x)))

/*******************************************************************************************
*	Автор: Чернопятов Е.А. mail-to: blackheel@mail.ru
*	Исходные тексты программ для статьи "Автоматизация приложений MS Office (часть 4)"
*
*	Опубликовано на сайтах 
*	http://blackheel.ru
*	http://firststeps.ru
*   
*	Данные исходные тексты могут быть свободно воспроизведены в любом ИНТЕРНЕТ-издании
*   ТОЛЬКО ПРИ УСЛОВИИ УКАЗАНИЯ АВТОРА (Чернопятов Е.А.)
*
*	Запрещается использовать данный исходный текст в любых ПЕЧАТНЫХ изданиях
*	без предварительного согласования с автором.
*******************************************************************************************/



// CExcelTestDlg dialog



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
	
	if(!app.CreateDispatch(_T("Excel.Application"))) //запустить сервер
	{
		AfxMessageBox(_T("Ошибка при старте Excel!"));
		return;
	}
	else
		app.put_Visible(TRUE); //и сделать его видимым
	
	CWorkbooks oBooks;
	CWorkbook oBook;

	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant covTrue((short)TRUE, VT_BOOL), covFalse((short)FALSE, VT_BOOL);
	COleVariant covBOOL((short)FALSE, VT_BOOL);

	//наша коллекция раБочиХ книг
	oBooks = app.get_Workbooks();
	//добавить к ней новую книгу с шаблоном по умолчанию	
	oBooks.Add(covOptional);	//для Office XP
	//и получить его как экзепляр коллекции с номером 1
	oBook = oBooks.get_Item(COleVariant(long(1)));
	//активизировать документ	
	oBook.Activate();
	


	CWorksheets oSheets;
	CWorksheet oSheet;


	//наша коллекция рабочих листов
	oSheets = oBook.get_Worksheets();	
	// добавить еще один рабочий лист (добавляем в начало, 1 лист, тип "Рабочий лист")
	oSheet = oSheets.Add(covOptional,covOptional,
		COleVariant(long(1)),COleVariant(long(xlWorksheet)));	//для Office XP		
	//активизировать лист
	oSheet.put_Name(_T("Мой рабочий лист"));
	oSheet.Activate();

	// РАБОТА С ОДИНОЧНЫМИ ЯЧЕЙКАМИ И ДИАПАЗОНОМ
	
	// данные и формулы занесть (пример работы с отдельными ячейками)....
	CRange oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A1")));
	oRange.put_Formula(COleVariant(_T("123")));
	oRange = oSheet.get_Range(COleVariant(_T("A2")),COleVariant(_T("A2")));
	oRange.put_Formula(COleVariant(_T("456")));
	oRange = oSheet.get_Range(COleVariant(_T("A3")),COleVariant(_T("A3")));
	oRange.put_Formula(COleVariant(_T("=SUM(A1:A2)")));

	// диапазон ячеек А1:А3
	oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A3")));
	// границы диапазона поменять на двойные....
	CBorders oBorders;
	CBorder  oBorder;
	oBorders = oRange.get_Borders();	
	oBorders.put_LineStyle(COleVariant(long(xlDouble)));
	// ... и внутрение убрать 
	oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A3")));
	oBorders = oRange.get_Borders();	 
	oBorder = oBorders.get_Item(long(xlInsideHorizontal));	
	oBorder.put_LineStyle(COleVariant(long(xlLineStyleNone)));

	// слить ячейки с С2 по D4 в одну
	oRange = oSheet.get_Range(COleVariant(_T("C2")),COleVariant(_T("D4")));
	oRange.Merge(covFalse);
	// добавить длинный текст, сделать перенос 
	oRange.put_WrapText(covTrue);
	oRange.put_Formula(COleVariant(_T("Длинный текст с переносом строк при достижении правой границы")));

	
	// ПРИМЕР РАБОТЫ С ДВУМЕРНЫМ МАССИВОМ ДАННЫХ (МАТРИЦЕЙ)
	// будем заносить и считывать матрицу 10*10 в/из диапазона ячеек A10:J19
	CRange oRange1;	
	oRange1 = oSheet.get_Range(COleVariant(_T("A10")),COleVariant(_T("J19")));
	
	COleSafeArray saMatrixToExcel;
	DWORD numElements[] = {10, 10};

	
	// --- запись значений из массива в ячейки, входящие в oRange1 ....

	// создать вариантный массив, тип double (64-bit)
	saMatrixToExcel.Create(VT_R8, 2, numElements);
	ASSERT(saMatrixToExcel.GetDim() == 2);
	
	long index[2];
	double val;
	for(index[0]=0; index[0]<10; index[0]++)
	{
		for(index[1]=0; index[1]<10; index[1]++)
		{
			val = index[0] + index[1]*10 + 0.1;
			// заполнить его какими-то данными
			saMatrixToExcel.PutElement(index, &val);
		}
	}

	// поместить данные массива в желаемый диапазон ячеек
	oRange1.put_Value2(saMatrixToExcel); 

	// поменять число 99.1 в правой нижней ячейке диапазона
	// на строку (нужно для следующего примера на чтение)
	oRange1 = oSheet.get_Range(COleVariant(_T("J19")),COleVariant(_T("J19")));
	oRange1.put_Value2(COleVariant(_T("Тест")));


	
	// --- чтение значений из диапазона ячеек в массив ....

	// ! ЭТА ЧАСТЬ КОДА ПРЕДНАЗНАЧЕНА ДЛЯ ПРОСМОТРА В ОТЛАДЧИКЕ, 
	// ПОТОМУ ЧТО ПОЛУЧЕННЫЕ ЗНАЧЕНИЯ ДАМПЯТСЯ В ОКНО OUTPUT/DEBUG		
	oRange1 = oSheet.get_Range(COleVariant(_T("A10")),COleVariant(_T("J19")));
	
	// get_Value|get_Value2 вернет нам либо 1 значение (если диапазон содержит тольrо одну ячейку),
	// либо SAFEARRAY, если oRange1 представляет собой прямоугольную область		
	VARIANT va;
	va = oRange1.get_Value2();
	COleSafeArray saMatrixFromExcel(va); // в нашем примере будет массив
	VARIANT el;
	long row_start, col_start, row_count, col_count;
	
	ASSERT(saMatrixFromExcel.GetDim() == 2); // у нас должна быть матрица
	saMatrixFromExcel.GetLBound(1,&row_start);
	saMatrixFromExcel.GetLBound(2,&col_start);
	saMatrixFromExcel.GetUBound(1,&row_count);
	saMatrixFromExcel.GetUBound(2,&col_count);


	for(index[0]=row_start; index[0]<=row_count; index[0]++)
	{
		for(index[1]=col_start; index[1]<=col_count; index[1]++)
		{
			// содержимое saMatrixFromExcel представляет собой матрицу
			// элементов типа VARIANT, именно так ее передает нам Excel
			saMatrixFromExcel.GetElement(index,&el);
			// кто в теремочке живёт?
			switch (el.vt)
			{
			case VT_R8:				
				TRACE("%g ", el.dblVal);
				break;
			case VT_BSTR: 
				{
					CString s(el);
					TRACE("%s ", s);
				}
				break;
			default: // прочие типы в этом примере не используются
				break;
			}
		}
		TRACE("\n", el.bstrVal);
	}

	
	// ПРИМЕР РАБОТЫ С ДИАГРАММОЙ

	// диаграмма
	CCharts oCharts;
	CChart  oChart;
	oCharts = oBook.get_Charts();
	oChart = oCharts.Add(covOptional,covOptional,COleVariant(long(1)));
	oChart.put_ChartType(long(xlBarClustered));
	oChart.put_Name(_T("Моя диаграмма"));
	// данные для диаграммы
	oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A3")));	
	oChart.SetSourceData(oRange,covOptional);
	oChart.put_HasTitle(TRUE);
	

	CChartTitle oChartTitle;
	oChartTitle = oChart.get_ChartTitle();
	oChartTitle.put_Caption(_T("Заголовок моей диаграммы"));
	
	oChart.Activate();	

		

	// --- Пример корректного закрытия  Excel-я программно
	// --- отпускаем все интерфейсы, закрываем приложение 
	// oBook.Close(covFalse, covOptional, covOptional); 
	// oBook.ReleaseDispatch(); 
	// oBooks.Close(); 
	// oBooks.ReleaseDispatch();
	// app.Quit();	
	// app.ReleaseDispatch();
}


void CExcelTestDlg::OnBnClickedCancel()
{
		
	// --- Выход из программы без закрытия Excel, просто отсоединяемся от него
	app.DetachDispatch();

	OnCancel();
}
