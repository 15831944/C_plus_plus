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
	//hwnd = FindWindowEx(hwnd, NULL, "V8Dockbar", NULL);

	//hwnd = FindWindowEx(hwnd, NULL, "V8CommandBarDockFrame", NULL);
	//hwnd = FindWindowEx(hwnd, NULL, "V8CommandBar", NULL);
	hwnd = GetNextWindow(hwnd,2);
    if (hwnd!=NULL)
    {
		SetForegroundWindow(hwnd);
		SendKeys::SendWait("{ENTER}");

		HMENU hMenu;
		hMenu=GetMenu(hwnd);		
		if (hMenu!=NULL) 
		{
			cout << "Menu YES " << endl;
		}
		else
		{
		  cout << "Error Find Menu " << endl;		
		}

		////SendKeys::SendWait("Общие");
		////SendKeys::SendWait("{DOWN}");
		////SendKeys::SendWait("{DOWN}");
		cout << "OK" << endl;
		cin.get();

    }
    else 
		{
			cout << "Error Find Window" << endl;
			cin.get();
		}
	
 
}

//#define _AFXDLL
//#include "afxwin.h"
//#include <iostream>
//#using <System.Drawing.dll>
//#using <System.Windows.Forms.dll>
//#using <System.dll>
//#include "windows.h"
//using namespace std;
//using namespace System;
//using namespace System::Runtime::InteropServices;
//using namespace System::Drawing;
//using namespace System::Windows::Forms;
//
//void main()
//{
//  HWND hwnd;
//  hwnd=FindWindow("CalcFrame","Calculator");
//  if (hwnd!=NULL) 
//  {
//	  HMENU hMenu;
//	  hMenu=GetMenu(hwnd);
//	  if (hMenu!=NULL)
//	  {
//		  int iCount;
//		  iCount=GetMenuItemCount(hMenu);
//		  cout << "Menu Item - " << iCount << endl;
//		  cin.get();
//	  }
//	  else cout << " Error Loading Menu" << endl;
//	  cin.get();
//  }
//  else cout << " Error Find Windows" << endl;
//  cin.get();
//}