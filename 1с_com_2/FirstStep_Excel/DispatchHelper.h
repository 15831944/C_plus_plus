#pragma once

#include <comdef.h>

class CDispatchHelper
{
public:
	CDispatchHelper(void);
	~CDispatchHelper(void);

	HRESULT CreateInstance(LPCOLESTR sProgID, DWORD dwClsContext = CLSCTX_LOCAL_SERVER);
	IDispatch* GetDispatch(void);
	ULONG Release(void);

	VARIANT Get(LPOLESTR sParam);
	VARIANT Put(LPOLESTR sParam, _variant_t varValue);
	VARIANT Invoke(LPOLESTR sFunc);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6, _variant_t varParam7);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6, _variant_t varParam7, _variant_t varParam8);
	VARIANT Invoke(LPOLESTR sFunc, _variant_t varParam1, _variant_t varParam2, _variant_t varParam3, _variant_t varParam4, _variant_t varParam5, _variant_t varParam6, _variant_t varParam7, _variant_t varParam8, _variant_t varParam9);

	CDispatchHelper& operator=(VARIANT& v)
	{
		if( (v.vt == VT_DISPATCH)&&(v.pdispVal != NULL) )
		{
			if( m_pDispatch ) m_pDispatch->Release();
			m_pDispatch = NULL;
			m_pDispatch = v.pdispVal;
		}
		return *this;
	}

	IDispatch* operator->()
	{
		return m_pDispatch;
	}

	operator bool() const
	{
		return (m_pDispatch != NULL);
	}
protected:
	IDispatch* m_pDispatch;
};
