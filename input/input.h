
// input.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CinputApp:
// �йش����ʵ�֣������ input.cpp
//

class CinputApp : public CWinApp
{
public:
	CinputApp();

// ��д
public:
	virtual BOOL InitInstance();





// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CinputApp theApp;