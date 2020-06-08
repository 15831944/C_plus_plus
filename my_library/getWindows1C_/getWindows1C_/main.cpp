#define _AFXDLL
#include "afxwin.h"
#include <iostream>
#using <System.Drawing.dll>
#using <System.Windows.Forms.dll>
#using <System.dll>

using namespace std;
using namespace System;
using namespace System::Runtime::InteropServices;
using namespace System::Drawing;
using namespace System::Windows::Forms;


void main()
{
    HWND hwnd;
    hwnd=FindWindow("V8TopLevelFrame",0);
	hwnd = FindWindowEx(hwnd, NULL, "V8AutoHideLayouter", NULL);
	hwnd = FindWindowEx(hwnd, NULL, "V8NewLocalFrameBaseWnd", NULL);
	hwnd = FindWindowEx(hwnd, NULL, "V8LayouterTabsWindow", NULL);
	hwnd = FindWindowEx(hwnd, NULL, "V8FormElement", NULL);
	hwnd = FindWindowEx(hwnd, NULL, "V8Grid", NULL);
	
    if (hwnd!=NULL)
    {
		SetForegroundWindow(hwnd);
	
		//SendKeys::SendWait("Константы");
		SendKeys::SendWait("^{INSERT}");

		//		char * buffer = NULL;
		//if ( OpenClipboard() ) 
		//{
		//	HANDLE hData = GetClipboardData( CF_TEXT );
		//	buffer = (char*)GlobalLock( hData );
		//	char* psubstr = strstr(buffer, "Константы");
		//	if(psubstr)
		//	{
		//		cout << "Bingo!" << endl;
		//	}
		//	GlobalUnlock(hData);
		//	CloseClipboard();
		//}


		//SendKeys::SendWait("{RIGHT}");
		//SendKeys::SendWait("{DOWN}");
		//SendKeys::SendWait("{DOWN}");
		cout << "OK" << endl;
		cin.get();
    }
    else 
		{
			cout << "Error Find Window" << endl;
			cin.get();
		}

	
 
}