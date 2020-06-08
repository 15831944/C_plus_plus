
#include "stdafx.h"
#include "ExcelTest.h"
#include "ExcelTestDlg.h"
#include "Excel_enums.h"
#include <iostream>
#include <string>
#include "cstringt.h"

#include <afx.h>
#include <afxwin.h>
//
#include <sstream>
//

#include <stdlib.h>


#include "atlbase.h"
#include "atlstr.h"
#include "comutil.h"



#include "DispatchHelper.h"






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
	DDX_Control(pDX, IDC_LIST1, m_list);
	
}


BEGIN_MESSAGE_MAP(CExcelTestDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_TEST, OnBnClickedTest)
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)

	ON_LBN_SELCHANGE(IDC_LIST1, &CExcelTestDlg::OnLbnSelchangeList1)
	ON_LBN_DBLCLK(IDC_LIST1, &CExcelTestDlg::OnLbnDblclkList1)
	ON_WM_MENUSELECT()
	ON_BN_CLICKED(IDC_TEST2, &CExcelTestDlg::OnBnClickedTest2)
	ON_BN_CLICKED(IDC_CHECK1, &CExcelTestDlg::OnBnClickedCheck1)
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

	if( v8COMBase ) // существет ли у нас COM соединение?
	{
		// чтобы процесс не висел в ТМ, вызываем функцию выхода из 1С
		v8COMBase.Invoke(L"Exit", false);
		v8COMBase.Release(); // уничтожаем COM соединение
	}
}



void CExcelTestDlg::OnBnClickedTest2()
{

	HRESULT hr;
	VARIANT varRet;
 
	hr = v8COMBase.CreateInstance(L"V8.Application"); // создаем COM соединение с 1С
	if( FAILED(hr) ) // проверка на присутствие 1С Automation Sever-а
	{
		AfxMessageBox(_T("Переустановите 1С Предприятие!"));
		return;
	}

	CString m_strPath = _T("C:\\base");
	CString m_Usr = _T("Администратор");
	CString m_Pwd = _T("");

	varRet = v8COMBase.Invoke(L"Connect", L"File=C:\\base;Usr=Администратор;Pwd=;");

	if( (varRet.vt != VT_BOOL)||(varRet.boolVal == 0) ) // проверка удачно/не удачно
	{
		AfxMessageBox(_T("Невозможно подключиться к базе 1С Предприятия"));
		v8COMBase.Release();
		return;
	}
	CDispatchHelper v8SpravPol; // наш справочник
 
	v8SpravPol = v8COMBase.Invoke(L"Метаданные.ОбщиеМодули", (IDispatch*)NULL, (IDispatch*)NULL);
	if( !v8SpravPol )
	{
		AfxMessageBox(_T("Невозможно подключиться к [Метаданные.ОбщиеМодули]"));
		v8SpravPol.Release();
		return;
	}




	v8SpravPol.Release();
	//CDispatchHelper v8Admin; // наш Админ :)
 //
	//v8Admin = v8SpravPol.Invoke(L"НайтиПоНаименованию", L"Администратор", false, (IDispatch*)NULL, (IDispatch*)NULL);
	//if( !v8Admin )
	//{
	//	AfxMessageBox(_T("Пользователь Администратор не найден"));
	//	v8SpravPol.Release();
	//	v8COMBase.Release();
	//	return;
	//}
	//VARIANT varRet1; // переменная для хранения результата
 //
	//varRet1 = v8Admin.Get(L"Наименование"); // получаем значение поля
	//AfxMessageBox(_T(CString(varRet1.bstrVal))); // выводим сообщение о прочитанном имени пользователя

	//v8Admin.Release(); // Админ
	//v8SpravPol.Release(); // справочник
}



void CExcelTestDlg::OnBnClickedCancel()
{
	if( v8COMBase ) // существет ли у нас COM соединение?
	{
		// чтобы процесс не висел в ТМ, вызываем функцию выхода из 1С
		v8COMBase.Invoke(L"Exit", false);
		v8COMBase.Release(); // уничтожаем COM соединение
	}
	OnCancel();
}




void CExcelTestDlg::OnLbnSelchangeList1()
{
	
}



void CExcelTestDlg::OnLbnDblclkList1()
{

	
}





void CExcelTestDlg::OnMenuSelect(UINT nItemID, UINT nFlags, HMENU hSysMenu)
{
	CDialog::OnMenuSelect(IDR_MENU1, nFlags, hSysMenu);

	// TODO: добавьте свой код обработчика сообщений
	MessageBox("sdsl;ajslkdj");
}




void CExcelTestDlg::OnBnClickedCheck1()
{
	// TODO: добавьте свой код обработчика уведомлений
}
