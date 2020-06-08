
#include <iostream>


#include <ole2.h> // OLE2 Definitions

// AutoWrap() - Automation helper function...
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...) {
    // ������� ���������� ������ ����������...
    va_list marker;
    va_start(marker, cArgs);

    if(!pDisp) {
        MessageBox(NULL, "NULL IDispatch passed to AutoWrap()", "Error", 0x10010);
        _exit(0);
    }

    // ���������� ������������...
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;
    char buf[200];
    char szName[200];

    
    // ������������ ���� � ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);
    
    //  �������� ������������� �������� ������ � ���������� dispID
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if(FAILED(hr)) {
        sprintf_s(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
        MessageBox(NULL, buf, "AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }
    
	//���� ������������� ������ ������, �������������� ������ ��� ��� ������:
    // �������� ������ ��� ����������...
    VARIANT *pArgs = new VARIANT[cArgs+1];
    // ��������� ��������� �� �������
    for(int i=0; i<cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }
    
    // �������� ��������� � ���������-��������� DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;
    
    // Handle special-case for property-puts!
    if(autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }
    
    // � �������-�� �������� ������ ��� �����:
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if(FAILED(hr)) {
        sprintf_s(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, hr);
        MessageBox(NULL, buf, "AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }
    // ����� ���� ������� �� �����:
    va_end(marker);
    
    delete [] pArgs;
    
    return hr;
}

int main()
{
	 // ��������������� COM ��� ����� ������...
   CoInitialize(NULL);

   // �������� CLSID ��� ������ �������...
   CLSID clsid;
   HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

   if(FAILED(hr)) {

      ::MessageBox(NULL, "CLSIDFromProgID() failed", "Error", 0x10010);
      return -1;
   }

   // ���� �� ����� (� ���������� clsid) ������������� CLSID... 
   //�� ����� ��������� ������ � �������� IDispatch...
   IDispatch *pXlApp;
   hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pXlApp);
   if(FAILED(hr)) {
      ::MessageBox(NULL, "Excel �� ��������������� ���������", "Error", 0x10010);
      return -2;
   }



   //// �������� COM ������ ������� (�.�. app.visible = 1)
   //{

   //   VARIANT x;
   //   x.vt = VT_I4;
   //   x.lVal = 1;
   //   AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x);
   //}



      // �������� ������������...
   ::MessageBox(NULL, "All done.", "Notice", 0x10000);



//----------------------------END-----------------------

   //// ������� Excel �������� (�.�. ����������. �����)
   //AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);

   // ������ �������...
   //pXlRange->Release();
   //pXlSheet->Release();
   //pXlBook->Release();
   //pXlBooks->Release();
   pXlApp->Release();
   //VariantClear(&arr);

   // ����������������� COM ��� ����� ������...
   CoUninitialize();
   

}