
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
	pv8->Release();
}



void CExcelTestDlg::OnBnClickedTest2()
{

	//для начала инициализируем COM
	HRESULT hr = CoInitialize(NULL);
	if(FAILED(hr))
	{
	  AfxMessageBox(_T("Невозможно инициализировать COM!"));
	  return;
	}

	CLSID   cls8;

	hr = CLSIDFromProgID(L"V82.Application", &cls8); 
	if(FAILED(hr))
	{
	  AfxMessageBox(_T("Переустановите 1С Предприятие!"));
	  CoUninitialize();
	  return;
	}


	hr = CoCreateInstance(cls8, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pv8);
 
	if(FAILED(hr) || !pv8)
	{
		AfxMessageBox("Невозможно инициализировать интерфейс 1С Предприятия"); 
		CoUninitialize();
		return;
	}

	DISPID dispIDCRM, dispIDInitialize = 0;

	BSTR init = L"Connect";
	hr = pv8->GetIDsOfNames(IID_NULL, &init, 1, 0, &dispIDInitialize);
	if (FAILED(hr))
	{
		AfxMessageBox("Не удалось получить ID от Initialize");
		if (pv8) pv8->Release();
		CoUninitialize();
		return;
	}

	// Connect принимает только 1 параметр
	DISPPARAMS args = {0, 0, 0, 0};
	VARIANT vars[3];  // Параметры для вызова Initialize
	CString m_strPath = _T("C:\\base");

	args.cArgs = 1;
	args.rgvarg = vars;
	ZeroMemory(vars, sizeof(vars));
	CString stroka;
	stroka = "File=\""
		 + m_strPath  // Путь к информационной базе
		 + "\\\";Usr=\""+ "Администратор" // Имя пользователя
		 + "\";Psw=\"" + "" // Пароль для входа
		 + "\";";
	// Переменные m_* типа CString
	  _bstr_t bstrS =(_T(stroka));
	  vars[0].vt = VT_BSTR;
	  vars[0].bstrVal = bstrS.GetBSTR();

	  // Ну а теперь коннектимся
	VARIANT vRet;
  hr = pv8->Invoke(dispIDInitialize, IID_NULL, 0, DISPATCH_PROPERTYGET, &args,
    &vRet, NULL, NULL);


  if(FAILED(hr) ||  (vRet.vt ==  VT_BOOL && vRet.bstrVal == 0x00))
  {
    AfxMessageBox("Невозможно запустить 1С Предприятие");
    if (pv8) pv8->Release();
    CoUninitialize();
    return;
  }


}



void CExcelTestDlg::OnBnClickedCancel()
{
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
