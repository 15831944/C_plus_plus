
// bbb.h : ������� ���� ��������� ��� ���������� bbb
//
#pragma once

#ifndef __AFXWIN_H__
	#error "�������� stdafx.h �� ��������� ����� ����� � PCH"
#endif

#include "resource.h"       // �������� �������


// CbbbApp:
// � ���������� ������� ������ ��. bbb.cpp
//

class CbbbApp : public CWinAppEx
{
public:
	CbbbApp();


// ���������������
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

// ����������
	BOOL  m_bHiColorIcons;

	virtual void PreLoadState();
	virtual void LoadCustomState();
	virtual void SaveCustomState();

	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CbbbApp theApp;
