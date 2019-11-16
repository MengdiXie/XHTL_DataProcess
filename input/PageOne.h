#pragma once


// PageOne 对话框
#include "inputDlg.h"
class PageOne : public CDialogEx
{
	DECLARE_DYNAMIC(PageOne)

public:
	PageOne(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~PageOne();

// 对话框数据
	enum { IDD = IDD_SHEET1 };

	CString m_boardnum;
	CString m_opname;
	CString m_tempture;
	CString m_humidity;
	CString m_multimeter;
	CString m_constant_volt_source;
	CString m_signal_source;
	CString m_测试工装;

	CString m_测试阶段;



	BOOL m_dui1;
	BOOL m_dui2;
	BOOL m_dui3;
	BOOL m_dui4;
	BOOL m_dui5;
	BOOL m_dui6;
	BOOL m_dui7;
	BOOL m_dui8;
	BOOL m_dui9;
	BOOL m_dui10;
	BOOL m_dui11;
	BOOL m_dui12;
	BOOL m_cuo1;
	BOOL m_cuo2;
	BOOL m_cuo3;
	BOOL m_cuo4;
	BOOL m_cuo5;
	BOOL m_cuo6;
	BOOL m_cuo7;
	BOOL m_cuo8;
	BOOL m_cuo9;
	BOOL m_cuo10;
	BOOL m_cuo11;
	BOOL m_cuo12;


	CApplication app;
    CWorkbooks books;
    CWorkbook book;
    CWorksheets sheets;
    CWorksheet sheet;
    CRange range;
    CRange iCell;
    LPDISPATCH lpDisp;

	CString m_inputstr;

	CString m_inputstr2;
	CString m_inputstr3;
	CString m_inputstr4;
	CString m_inputstr5;
	CString m_inputstr6;
	CString m_inputstr7;
	CString m_inputstr8;

    CFont my_Font;


	CBrush m_Brush;


public:
	std::vector<float> CStringtoFloat(CString m_s);

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedtwo();
	afx_msg void OnBnClickedthree();
	afx_msg void OnBnClickedstop();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedCancel();
};
