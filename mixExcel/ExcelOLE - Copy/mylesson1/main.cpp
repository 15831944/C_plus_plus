
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

   // �������� COM ������ ������� (�.�. app.visible = 1)
   {

      VARIANT x;
      x.vt = VT_I4;
      x.lVal = 1;
      AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x);
   }

   // �������� ��������� ���� Excel
   IDispatch *pXlBooks;
   {
      VARIANT result;
      VariantInit(&result);
      AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"Workbooks", 0);
      pXlBooks = result.pdispVal;
   }

   // �������� ����� �����
   IDispatch *pXlBook;
   {
      VARIANT result;
      VariantInit(&result);
      AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, L"Add", 0);
      pXlBook = result.pdispVal;
   }
//------------------------------------------------------------------------
   // ��� ������������ ���������� ������� ���������� ������� ������� 15�15
   VARIANT arr;
   arr.vt = VT_ARRAY | VT_VARIANT;
   {
      SAFEARRAYBOUND sab[2];
      sab[0].lLbound = 1; sab[0].cElements = 15; // ���������� �����
      sab[1].lLbound = 1; sab[1].cElements = 15;// ���������� �������
      arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
   }

   // ��������� safearray ���������� ����������...
   for(int i=1; i<=15; i++) {
      for(int j=1; j<=15; j++) {
         // �������� ������� �������� ��� (�, j)
         VARIANT tmp;
         tmp.vt = VT_I4;
         tmp.lVal = i*j;
         // �������� � safearray...
         long indices[] = {i,j};
         SafeArrayPutElement(arr.parray, indices, (void *)&tmp);
      }
   }
//-----------------------------------------------------------------------
      // �������� ������������...
   ::MessageBox(NULL, "All done.", "Notice", 0x10000);
   //�������� �������� ���� ����� Excel
   IDispatch *pXlSheet;
   {
      VARIANT result;
      VariantInit(&result);
      AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0);
      pXlSheet = result.pdispVal;
   }
   //����� ��������� ������� � �xcel ���������� ���������� �������� �������
   // �������� ������ ��������� ��� ��������� A1:O15...
   IDispatch *pXlRange;
   {
      VARIANT parm;
      parm.vt = VT_BSTR;
      parm.bstrVal = ::SysAllocString(L"A1:O15");

      VARIANT result;
      VariantInit(&result);
      AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm);
      VariantClear(&parm);

      pXlRange = result.pdispVal;
   }

   // ����� �����, �������� � ������� ��������� ������:
   AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlRange, L"Value", 1, arr);

   // �������� ������������...
   ::MessageBox(NULL, "All done.", "Notice", 0x10000);

// Set .Saved �������� ������� ����� � ������, ����� �������, ��� �� ����������
   // ���������, ����� �� ������� Excel ��������...
   {
      VARIANT x;
      x.vt = VT_I4;
      x.lVal = 1;
      AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlBook, L"Saved", 1, x);
   }

   // ������� Excel �������� (�.�. ����������. �����)
   AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);

   // ������ �������...
   pXlRange->Release();
   pXlSheet->Release();
   pXlBook->Release();
   pXlBooks->Release();
   pXlApp->Release();
   VariantClear(&arr);

   // ����������������� COM ��� ����� ������...
   CoUninitialize();
   

}