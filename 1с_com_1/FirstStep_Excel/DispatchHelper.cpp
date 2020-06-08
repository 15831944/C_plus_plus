#include "StdAfx.h"
#include "DispatchHelper.h"

CDispatchHelper::CDispatchHelper(void)
{
	m_pDispatch = NULL;
}

CDispatchHelper::~CDispatchHelper(void)
{
	if( m_pDispatch ) m_pDispatch->Release();
	m_pDispatch = NULL;
}

VARIANT CDispatchHelper::Get(LPOLESTR sParam)
{
	HRESULT hr;
	DISPID dispid;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sParam, 1, 0, &dispid);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;

	dispidParams.cArgs = 0;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = NULL;

	hr = m_pDispatch->Invoke(dispid, IID_NULL, 0, DISPATCH_PROPERTYGET, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Put(LPOLESTR sParam, _variant_t varValue)
{
	HRESULT hr;
	DISPID dispid;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sParam, 1, 0, &dispid);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[1];
	DISPID dispParams[1];

	dispidParams.cArgs = 1;
	dispidParams.cNamedArgs = 1;
	dispidParams.rgdispidNamedArgs = dispParams;
	dispidParams.rgvarg = varParams;
	varParams[0] = varValue;
	dispParams[0] = DISPID_PROPERTYPUT;

	hr = m_pDispatch->Invoke(dispid, IID_NULL, 0, DISPATCH_PROPERTYPUT, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;

	dispidParams.cArgs = 0;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = NULL;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[1];  

	dispidParams.cArgs = 1;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[2];  

	dispidParams.cArgs = 2;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam2;
	varParams[1] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[3];  

	dispidParams.cArgs = 3;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam3;
	varParams[1] = varParam2;
	varParams[2] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[4];  

	dispidParams.cArgs = 4;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam4;
	varParams[1] = varParam3;
	varParams[2] = varParam2;
	varParams[3] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[5];  

	dispidParams.cArgs = 5;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam5;
	varParams[1] = varParam4;
	varParams[2] = varParam3;
	varParams[3] = varParam2;
	varParams[4] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[6];  

	dispidParams.cArgs = 6;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam6;
	varParams[1] = varParam5;
	varParams[2] = varParam4;
	varParams[3] = varParam3;
	varParams[4] = varParam2;
	varParams[5] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6, _variant_t varParam7)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[7];

	dispidParams.cArgs = 7;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam7;
	varParams[1] = varParam6;
	varParams[2] = varParam5;
	varParams[3] = varParam4;
	varParams[4] = varParam3;
	varParams[5] = varParam2;
	varParams[6] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6, _variant_t varParam7, _variant_t varParam8)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[8];

	dispidParams.cArgs = 8;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam8;
	varParams[1] = varParam7;
	varParams[2] = varParam6;
	varParams[3] = varParam5;
	varParams[4] = varParam4;
	varParams[5] = varParam3;
	varParams[6] = varParam2;
	varParams[7] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

VARIANT CDispatchHelper::Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6, _variant_t varParam7, _variant_t varParam8, _variant_t varParam9)
{
	HRESULT hr;
	DISPID dispidFunc;
	VARIANT varRet;

	hr = m_pDispatch->GetIDsOfNames(IID_NULL, &sFunc, 1, 0, &dispidFunc);    

	if( FAILED(hr) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_PARAMNOTFOUND;
		return varRet;
	}

	DISPPARAMS dispidParams;
	VARIANT varParams[9];

	dispidParams.cArgs = 9;
	dispidParams.cNamedArgs = 0;
	dispidParams.rgdispidNamedArgs = NULL;
	dispidParams.rgvarg = varParams;
	varParams[0] = varParam9;
	varParams[1] = varParam8;
	varParams[2] = varParam7;
	varParams[3] = varParam6;
	varParams[4] = varParam5;
	varParams[5] = varParam4;
	varParams[6] = varParam3;
	varParams[7] = varParam2;
	varParams[8] = varParam1;

	hr = m_pDispatch->Invoke(dispidFunc, IID_NULL, 0, DISPATCH_METHOD, &dispidParams, &varRet, NULL, NULL);

	if( FAILED(hr) && (varRet.vt == VT_EMPTY) )
	{
		varRet.vt = VT_ERROR;
		varRet.scode = DISP_E_EXCEPTION;
	}

	return varRet;
}

HRESULT CDispatchHelper::CreateInstance(LPCOLESTR sProgID, DWORD dwClsContext)
{
	HRESULT hr;
	CLSID clsid;

	if( m_pDispatch )
	{
		m_pDispatch->Release();
		m_pDispatch = NULL;
	}

	hr = ::CLSIDFromProgID(sProgID, &clsid);
	if( FAILED(hr) ) return hr;

	return ::CoCreateInstance(clsid, NULL, dwClsContext, IID_IDispatch, (void**)&m_pDispatch);
}

IDispatch* CDispatchHelper::GetDispatch(void)
{
	return m_pDispatch;
}

ULONG CDispatchHelper::Release(void)
{
	ULONG ret = 0;

	if( m_pDispatch )
	{
		ret = m_pDispatch->Release();
		m_pDispatch = NULL;
	}

	return ret;
}
