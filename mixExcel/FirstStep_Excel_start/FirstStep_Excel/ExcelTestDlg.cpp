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
*	�����: ���������� �.�. mail-to: blackheel@mail.ru
*	�������� ������ �������� ��� ������ "������������� ���������� MS Office (����� 4)"
*
*	������������ �� ������ 
*	http://blackheel.ru
*	http://firststeps.ru
*   
*	������ �������� ������ ����� ���� �������� �������������� � ����� ��������-�������
*   ������ ��� ������� �������� ������ (���������� �.�.)
*
*	����������� ������������ ������ �������� ����� � ����� �������� ��������
*	��� ���������������� ������������ � �������.
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
	
	if(!app.CreateDispatch(_T("Excel.Application"))) //��������� ������
	{
		AfxMessageBox(_T("������ ��� ������ Excel!"));
		return;
	}
	else
		app.put_Visible(TRUE); //� ������� ��� �������
	
	CWorkbooks oBooks;
	CWorkbook oBook;

	COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant covTrue((short)TRUE, VT_BOOL), covFalse((short)FALSE, VT_BOOL);
	COleVariant covBOOL((short)FALSE, VT_BOOL);

	//���� ��������� ������� ����
	oBooks = app.get_Workbooks();
	//�������� � ��� ����� ����� � �������� �� ���������	
	oBooks.Add(covOptional);	//��� Office XP
	//� �������� ��� ��� �������� ��������� � ������� 1
	oBook = oBooks.get_Item(COleVariant(long(1)));
	//�������������� ��������	
	oBook.Activate();
	


	CWorksheets oSheets;
	CWorksheet oSheet;


	//���� ��������� ������� ������
	oSheets = oBook.get_Worksheets();	
	// �������� ��� ���� ������� ���� (��������� � ������, 1 ����, ��� "������� ����")
	oSheet = oSheets.Add(covOptional,covOptional,
		COleVariant(long(1)),COleVariant(long(xlWorksheet)));	//��� Office XP		
	//�������������� ����
	oSheet.put_Name(_T("��� ������� ����"));
	oSheet.Activate();

	// ������ � ���������� �������� � ����������
	
	// ������ � ������� ������� (������ ������ � ���������� ��������)....
	CRange oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A1")));
	oRange.put_Formula(COleVariant(_T("123")));
	oRange = oSheet.get_Range(COleVariant(_T("A2")),COleVariant(_T("A2")));
	oRange.put_Formula(COleVariant(_T("456")));
	oRange = oSheet.get_Range(COleVariant(_T("A3")),COleVariant(_T("A3")));
	oRange.put_Formula(COleVariant(_T("=SUM(A1:A2)")));

	// �������� ����� �1:�3
	oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A3")));
	// ������� ��������� �������� �� �������....
	CBorders oBorders;
	CBorder  oBorder;
	oBorders = oRange.get_Borders();	
	oBorders.put_LineStyle(COleVariant(long(xlDouble)));
	// ... � ��������� ������ 
	oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A3")));
	oBorders = oRange.get_Borders();	 
	oBorder = oBorders.get_Item(long(xlInsideHorizontal));	
	oBorder.put_LineStyle(COleVariant(long(xlLineStyleNone)));

	// ����� ������ � �2 �� D4 � ����
	oRange = oSheet.get_Range(COleVariant(_T("C2")),COleVariant(_T("D4")));
	oRange.Merge(covFalse);
	// �������� ������� �����, ������� ������� 
	oRange.put_WrapText(covTrue);
	oRange.put_Formula(COleVariant(_T("������� ����� � ��������� ����� ��� ���������� ������ �������")));

	
	// ������ ������ � ��������� �������� ������ (��������)
	// ����� �������� � ��������� ������� 10*10 �/�� ��������� ����� A10:J19
	CRange oRange1;	
	oRange1 = oSheet.get_Range(COleVariant(_T("A10")),COleVariant(_T("J19")));
	
	COleSafeArray saMatrixToExcel;
	DWORD numElements[] = {10, 10};

	
	// --- ������ �������� �� ������� � ������, �������� � oRange1 ....

	// ������� ���������� ������, ��� double (64-bit)
	saMatrixToExcel.Create(VT_R8, 2, numElements);
	ASSERT(saMatrixToExcel.GetDim() == 2);
	
	long index[2];
	double val;
	for(index[0]=0; index[0]<10; index[0]++)
	{
		for(index[1]=0; index[1]<10; index[1]++)
		{
			val = index[0] + index[1]*10 + 0.1;
			// ��������� ��� ������-�� �������
			saMatrixToExcel.PutElement(index, &val);
		}
	}

	// ��������� ������ ������� � �������� �������� �����
	oRange1.put_Value2(saMatrixToExcel); 

	// �������� ����� 99.1 � ������ ������ ������ ���������
	// �� ������ (����� ��� ���������� ������� �� ������)
	oRange1 = oSheet.get_Range(COleVariant(_T("J19")),COleVariant(_T("J19")));
	oRange1.put_Value2(COleVariant(_T("����")));


	
	// --- ������ �������� �� ��������� ����� � ������ ....

	// ! ��� ����� ���� ������������� ��� ��������� � ���������, 
	// ������ ��� ���������� �������� �������� � ���� OUTPUT/DEBUG		
	oRange1 = oSheet.get_Range(COleVariant(_T("A10")),COleVariant(_T("J19")));
	
	// get_Value|get_Value2 ������ ��� ���� 1 �������� (���� �������� �������� ����r� ���� ������),
	// ���� SAFEARRAY, ���� oRange1 ������������ ����� ������������� �������		
	VARIANT va;
	va = oRange1.get_Value2();
	COleSafeArray saMatrixFromExcel(va); // � ����� ������� ����� ������
	VARIANT el;
	long row_start, col_start, row_count, col_count;
	
	ASSERT(saMatrixFromExcel.GetDim() == 2); // � ��� ������ ���� �������
	saMatrixFromExcel.GetLBound(1,&row_start);
	saMatrixFromExcel.GetLBound(2,&col_start);
	saMatrixFromExcel.GetUBound(1,&row_count);
	saMatrixFromExcel.GetUBound(2,&col_count);


	for(index[0]=row_start; index[0]<=row_count; index[0]++)
	{
		for(index[1]=col_start; index[1]<=col_count; index[1]++)
		{
			// ���������� saMatrixFromExcel ������������ ����� �������
			// ��������� ���� VARIANT, ������ ��� �� �������� ��� Excel
			saMatrixFromExcel.GetElement(index,&el);
			// ��� � ��������� ����?
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
			default: // ������ ���� � ���� ������� �� ������������
				break;
			}
		}
		TRACE("\n", el.bstrVal);
	}

	
	// ������ ������ � ����������

	// ���������
	CCharts oCharts;
	CChart  oChart;
	oCharts = oBook.get_Charts();
	oChart = oCharts.Add(covOptional,covOptional,COleVariant(long(1)));
	oChart.put_ChartType(long(xlBarClustered));
	oChart.put_Name(_T("��� ���������"));
	// ������ ��� ���������
	oRange = oSheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A3")));	
	oChart.SetSourceData(oRange,covOptional);
	oChart.put_HasTitle(TRUE);
	

	CChartTitle oChartTitle;
	oChartTitle = oChart.get_ChartTitle();
	oChartTitle.put_Caption(_T("��������� ���� ���������"));
	
	oChart.Activate();	

		

	// --- ������ ����������� ��������  Excel-� ����������
	// --- ��������� ��� ����������, ��������� ���������� 
	// oBook.Close(covFalse, covOptional, covOptional); 
	// oBook.ReleaseDispatch(); 
	// oBooks.Close(); 
	// oBooks.ReleaseDispatch();
	// app.Quit();	
	// app.ReleaseDispatch();
}


void CExcelTestDlg::OnBnClickedCancel()
{
		
	// --- ����� �� ��������� ��� �������� Excel, ������ ������������� �� ����
	app.DetachDispatch();

	OnCancel();
}
