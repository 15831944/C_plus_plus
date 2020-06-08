
#include <iostream>


#include <ole2.h> // OLE2 Definitions

// AutoWrap() - Automation helper function...
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...) {
    // Начните переменный список параметров...
    va_list marker;
    va_start(marker, cArgs);

    if(!pDisp) {
        MessageBox(NULL, "NULL IDispatch passed to AutoWrap()", "Error", 0x10010);
        _exit(0);
    }

    // Переменные используются...
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;
    char buf[200];
    char szName[200];

    
    // Преобразуйте вниз в ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);
    
    //  получаем идентификатор искомого метода в переменную dispID
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if(FAILED(hr)) {
        sprintf_s(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
        MessageBox(NULL, buf, "AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }
    
	//Если идентификатор метода найден, подготавливаем данные для его вызова:
    // Выделите память для параметров...
    VARIANT *pArgs = new VARIANT[cArgs+1];
    // Извлекаем аргументы из массива
    for(int i=0; i<cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }
    
    // Помещаем аргументы в структуру-контейнер DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;
    
    // Handle special-case for property-puts!
    if(autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }
    
    // и наконец-то вызываем нужный нам метод:
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if(FAILED(hr)) {
        sprintf_s(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, hr);
        MessageBox(NULL, buf, "AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }
    // после чего убираем за собой:
    va_end(marker);
    
    delete [] pArgs;
    
    return hr;
}

int main()
{
	 // Инициализируйте COM для этого потока...
   CoInitialize(NULL);

   // Получите CLSID для нашего сервера...
   CLSID clsid;
   HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

   if(FAILED(hr)) {

      ::MessageBox(NULL, "CLSIDFromProgID() failed", "Error", 0x10010);
      return -1;
   }

   // Имея на руках (в переменной clsid) идентификатор CLSID... 
   //мы можем запустить сервер и получить IDispatch...
   IDispatch *pXlApp;
   hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pXlApp);
   if(FAILED(hr)) {
      ::MessageBox(NULL, "Excel не зарегистрирован корректно", "Error", 0x10010);
      return -2;
   }



   //// Сделайте COM объект видимым (т.е. app.visible = 1)
   //{

   //   VARIANT x;
   //   x.vt = VT_I4;
   //   x.lVal = 1;
   //   AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x);
   //}



      // Ожидайте пользователя...
   ::MessageBox(NULL, "All done.", "Notice", 0x10000);



//----------------------------END-----------------------

   //// Скажите Excel выходить (т.е. Приложение. Выход)
   //AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);

   // Ссылки выпуска...
   //pXlRange->Release();
   //pXlSheet->Release();
   //pXlBook->Release();
   //pXlBooks->Release();
   pXlApp->Release();
   //VariantClear(&arr);

   // Деинициализируйте COM для этого потока...
   CoUninitialize();
   

}