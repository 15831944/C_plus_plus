#include "iostream"

#include "conio.h"
#include "windows.h"

using namespace std;

void main()
{
DEVMODE dm; 
dm.dmSize = sizeof(DEVMODE); 
int index = 0; 
while (EnumDisplaySettings(NULL, index, &dm)) 
{ 
  if (dm.dmPelsWidth == 1024 && dm.dmPelsHeight == 768)
  { 
    dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT; 
    LONG result = ChangeDisplaySettings(&dm, CDS_TEST); 
    if (result == DISP_CHANGE_SUCCESSFUL) 
    { 
      ChangeDisplaySettings(&dm, 0); 
      break; 
    } 
    else if (result == DISP_CHANGE_RESTART) 
    { 
      cout << "Требуется перезагрузка" << endl;
      break; 
    } 
    else 
    { 
      cout << "Установка не поддерживается монитором" << endl; 
      break; 
    } 
  } 
  index++; 
}
}