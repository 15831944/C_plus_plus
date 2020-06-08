
#include "stdafx.h"
#include "ExcelTest.h"
#include "ExcelTestDlg.h"
#include "Excel_enums.h"


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
	
	try
      {
      CApplication app;     // app is an _Application object.
      CWorkbook book;       // More object declarations.
      CWorksheet sheet;
      CWorkbooks books;
      CWorksheets sheets;

      CRange range;          // Used for Microsoft Excel 97 components.
      LPDISPATCH lpDisp;    // Often reused variable.

      // Common OLE variants. Easy variants to use for calling arguments.
      COleVariant
        covTrue((short)TRUE),
        covFalse((short)FALSE),
        covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

      // Start Microsoft Excel, get _Application object,
      // and attach to app object.
      if(!app.CreateDispatch("Excel.Application"))
       {
        AfxMessageBox("Couldn't CreateDispatch() for Excel");
        return;
       }


      // Set visible.
   	  app.put_Visible(TRUE);

      // Register the Analysis ToolPak.
      CString sAppPath;

	  //sAppPath.Format ("%s\\Analysis\\Analys32.xll", app.get_LibraryPath());

   //   if(!app.RegisterXLL(sAppPath))
   //     AfxMessageBox("Didn't register the Analys32.xll");

      // Get the Workbooks collection.
	  lpDisp = app.get_Workbooks();     // Get an IDispatch pointer.
      ASSERT(lpDisp);
      books.AttachDispatch(lpDisp);    // Attach the IDispatch pointer
                                       // to the books object.

      // Open a new workbook and attach that IDispatch pointer to the
      // Workbook object.
      lpDisp = books.Add( covOptional );
      ASSERT(lpDisp);
      book.AttachDispatch( lpDisp );

         // To open an existing workbook, you need to provide all
         // arguments for the Open member function. In the case of 
         // Excel 2002 you must provide 16 arguments.
         // However in Excel 2003 you must provide 15 arguments.
         // The code below opens a workbook and adds it to the Workbook's
         // Collection object. It shows 13 arguments, required for Excel
         // 2000.
         // You need to modify the path and file name for your own
         // workbook.

      // 
       lpDisp = books.Open("C:\\1",     // Test.xls is a workbook.
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional, covOptional,
       covOptional, covOptional, covOptional, covOptional );   // Return Workbook's IDispatchpointer.
	   book = books.get_Item(COleVariant(long(1)));
	//активизировать документ	
	   book.Activate();

      // Get the Sheets collection and attach the IDispatch pointer to your
      // sheets object.
	  lpDisp = book.get_Sheets();
      ASSERT(lpDisp);
      sheets.AttachDispatch(lpDisp);

      // Get sheet #1 and attach the IDispatch pointer to your sheet
      // object.
	  lpDisp = sheets.get_Item( COleVariant((short)(1)) );
                                        //GetItem(const VARIANT &index)
      ASSERT(lpDisp);
      sheet.AttachDispatch(lpDisp);

      lpDisp = sheet.get_Range(COleVariant("C3"), COleVariant("C3"));
      range.AttachDispatch(lpDisp);
      range.put_Formula(COleVariant(_T("123")));


      // Release dispatch pointers.
      range.ReleaseDispatch();
      sheet.ReleaseDispatch();
      // This is not really necessary because
      // the default second parameter of AttachDispatch releases
      // when the current scope is lost.

      } // End of processing.

        catch(COleException *e)
      {
        char buf[1024];     // For the Try...Catch error message.
        sprintf(buf, "COleException. SCODE: %08lx.", (long)e->m_sc);
        ::MessageBox(NULL, buf, "COleException", MB_SETFOREGROUND | MB_OK);
      }

      catch(COleDispatchException *e)
      {
        char buf[1024];     // For the Try...Catch error message.
        sprintf(buf,
               "COleDispatchException. SCODE: %08lx, Description: \"%s\".",
               (long)e->m_wCode,(LPSTR)e->m_strDescription.GetBuffer(512));
        ::MessageBox(NULL, buf, "COleDispatchException",
                           MB_SETFOREGROUND | MB_OK);
      }

      catch(...)
      {
        ::MessageBox(NULL, "General Exception caught.", "Catch-All",
                           MB_SETFOREGROUND | MB_OK);
      }
}


void CExcelTestDlg::OnBnClickedCancel()
{
		
	// --- ¬ыход из программы без закрыти€ Excel, просто отсоедин€емс€ от него
	app.DetachDispatch();

	OnCancel();
}
