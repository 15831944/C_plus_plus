#pragma once



// представление формы DragString

class DragString : public CFormView
{
	DECLARE_DYNCREATE(DragString)

protected:
	DragString();           // защищенный конструктор, используемый при динамическом создании
	virtual ~DragString();

public:
	enum { IDD = IDD_FORMVIEW };
#ifdef _DEBUG
	virtual void AssertValid() const;
#ifndef _WIN32_WCE
	virtual void Dump(CDumpContext& dc) const;
#endif
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // поддержка DDX/DDV

	DECLARE_MESSAGE_MAP()
};


