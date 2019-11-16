#pragma once

#include "inputDlg.h"
// DO 对话框

class DO : public CDialogEx
{
	DECLARE_DYNAMIC(DO)

public:
	DO(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~DO();


	CApplication app;
    CWorkbooks books;
    CWorkbook book;
    CWorksheets sheets;
    CWorksheet sheet;
    CRange range;
    CRange iCell;
    LPDISPATCH lpDisp;





	CString m_boardnum;
	CString m_opname;
	CString m_tempture;
	CString m_humidity;
	CString m_multimeter;
	CString m_constant_volt_source;
	CString m_signal_source;
	CString m_测试工装;

	CString m_strtime;


	CString m_edit12;
	CString m_edit13;
	CString m_edit14;
	CString m_edit15;
	CString m_OKorNot;

	CString m_edit16;
	CString m_edit17;
	CString m_edit18;
	CString m_edit19;

	CFont my_Font;


	CBrush m_Brush;


	BOOL m_click;

// 对话框数据
	enum { IDD = IDD_DO };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	static DWORD Button1Thread(LPVOID lpParam);
	static DWORD Button2Thread(LPVOID lpParam);
	static DWORD Button3Thread(LPVOID lpParam);

	static DWORD Button11Thread(LPVOID lpParam);
	static DWORD Button12Thread(LPVOID lpParam);


	std::vector<float> CStringtoFloat(CString m_s);

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();


	afx_msg LRESULT OnUpdateTrueMessage(WPARAM wParam, LPARAM lParam); 
	afx_msg LRESULT OnUpdateFalseMessage(WPARAM wParam, LPARAM lParam); 
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedButton2();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedButton10();
	afx_msg void OnBnClickedButton11();
	afx_msg void OnBnClickedButton12();
	afx_msg void OnBnClickedCancel();
};
