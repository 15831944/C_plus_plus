
// aaa.h : ������� ���� ��������� ��� ���������� aaa
//
#pragma once

#ifndef __AFXWIN_H__
	#error "�������� stdafx.h �� ��������� ����� ����� � PCH"
#endif

#include "resource.h"       // �������� �������


// CaaaApp:
// � ���������� ������� ������ ��. aaa.cpp
//

class CaaaApp : public CWinAppEx
{
public:
	CaaaApp();


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

extern CaaaApp theApp;
